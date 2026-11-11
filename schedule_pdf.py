# -*- coding: utf-8 -*-
"""「職員別スケジュール表(実績)」PDF の読み取り。

訪問看護ソフトが出力する PDF（1ファイル＝職員1人）から、
職員名・対象期間・訪問の明細を取り出す。

PDF の文字を行ごとに読むと、サービス内容や備考の折り返し（例「精神基本療養費Ⅰ・正／看」）
が次の行に分かれて列がずれる。そこで文字の位置（x, y 座標）で列と行を決める。
  - 列: 見出し「日付／分／ご利用者名／サービス内容／備考」の x 位置で区切る
  - 行: 日付の文字がある高さを1訪問とし、折り返した文字はいちばん近い訪問へ寄せる
"""
import re

import pypdf

import payroll_calc as calc

_DATE = re.compile(r"^(\d{1,2})月(\d{1,2})日\s*[（(]([月火水木金土日])[）)]")
_TIME = re.compile(r"^\d{1,2}:\d{2}$")
_ERA = re.compile(r"令和\s*(\d+)\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")

# 列の境界（実物の PDF で計測。各列の文字は列の左端から書き始められる）
#   日付 40 / 時刻 111 / 分 198 / 利用者 224 / サービス内容 324 / 備考 424
_BASE = {"time": 100.0, "dur": 190.0, "user": 215.0, "service": 310.0, "note": 405.0}
# そのときの見出し文字の x 位置。実際の見出しがずれていれば、その差だけ境界を動かす。
_BASE_HEAD = {"dur": ("分", 197.8), "user": ("ご利用者名", 248.8),
              "service": ("サービス内容", 344.3), "note": ("備考", 482.3)}


class ScheduleError(Exception):
    pass


def _fragments(page):
    out = []

    def visit(text, cm, tm, _fd, _fs):
        if not text or not text.strip():
            return
        x = tm[4] * cm[0] + tm[5] * cm[2] + cm[4]
        y = tm[4] * cm[1] + tm[5] * cm[3] + cm[5]
        out.append((x, y, text.strip()))

    page.extract_text(visitor_text=visit)
    return out


def _bounds(frags):
    """列境界と表の上端(y: 見出しの高さ)を求める。"""
    b = dict(_BASE)
    head_y = []
    for x, y, t in frags:
        if t == "日付":
            head_y.append(y)
        for col, (label, base_x) in _BASE_HEAD.items():
            if t == label:
                b[col] = _BASE[col] + (x - base_x)
                head_y.append(y)
    top = min(head_y) if head_y else None
    return b, top


def _column(x, b):
    if x < b["time"]:
        return "date"
    if x < b["dur"]:
        return "time"
    if x < b["user"]:
        return "dur"
    if x < b["service"]:
        return "user"
    if x < b["note"]:
        return "service"
    return "note"


def _header_info(frags):
    """職員名・事業所名・対象期間の文字列を取り出す。"""
    lines = {}
    for x, y, t in frags:
        lines.setdefault(round(y), []).append((x, t))
    joined = [" ".join(t for _x, t in sorted(v)) for _y, v in sorted(lines.items(), reverse=True)]
    person = office = period = ""
    for s in joined:
        if s.startswith("職員名") and not person:
            person = re.sub(r"^職員名\s*[:：]\s*", "", s).strip()
        elif s.startswith("事業所名") and not office:
            office = re.sub(r"^事業所名\s*[:：]\s*", "", s).strip()
        elif s.startswith("対象期間") and not period:
            period = re.sub(r"^対象期間\s*[:：]\s*", "", s).strip()
    return person, office, period


def _start_year(period_text):
    m = _ERA.search(period_text or "")
    if m:
        return 2018 + int(m.group(1)), int(m.group(2))
    return None, None


def read(path):
    """PDF を読み、dict(person, office, period_text, rows) を返す。

    rows は payroll_calc の8列（日付は 'YYYY-MM-DD'）。
    """
    try:
        reader = pypdf.PdfReader(path)
    except Exception as ex:  # 壊れたファイル・PDF 以外
        raise ScheduleError(f"PDF を開けませんでした: {ex}") from ex

    person = office = period_text = ""
    rows = []
    for page in reader.pages:
        frags = _fragments(page)
        if not frags:
            continue
        p, o, per = _header_info(frags)
        person = person or p
        office = office or o
        period_text = period_text or per
        b, top = _bounds(frags)

        # 1) 日付の文字がある高さ = 1訪問
        anchors = []
        for x, y, t in frags:
            if x < b["time"] and _DATE.match(t) and (top is None or y < top - 3):
                anchors.append([y, t, {"time": [], "dur": [], "user": [],
                                       "service": [], "note": []}])
        if not anchors:
            continue
        ys = sorted(a[0] for a in anchors)
        gap = min((ys[i + 1] - ys[i] for i in range(len(ys) - 1)), default=27.0)
        lowest = ys[0] - gap / 2

        # 2) 表の中の文字を、いちばん近い訪問の該当列へ入れる
        for x, y, t in frags:
            if top is not None and y >= top - 3:
                continue
            if y < lowest:
                continue                    # 表の下（「★は同行を…」・ページ番号）
            col = _column(x, b)
            if col == "date":
                continue
            a = min(anchors, key=lambda a: abs(a[0] - y))
            if abs(a[0] - y) > gap / 2 + 1:
                continue
            a[2][col].append((-y, x, t))

        for y, date_text, cols in sorted(anchors, key=lambda a: -a[0]):
            def joined(col, sep=""):
                return sep.join(t for _ny, _x, t in sorted(cols[col]))
            times = [t for _ny, _x, t in sorted(cols["time"], key=lambda v: v[1])
                     if _TIME.match(t)]
            m = _DATE.match(date_text)
            rows.append({
                "month": int(m.group(1)), "day": int(m.group(2)), "dow": m.group(3),
                "start": times[0] if times else "",
                "end": times[1] if len(times) > 1 else "",
                "dur": joined("dur"),
                "user": re.sub(r"様$", "", joined("user", " ")).strip(),
                "service": joined("service"),
                "note": joined("note"),
            })

    if not rows:
        raise ScheduleError("訪問の明細が見つかりませんでした。"
                            "「職員別スケジュール表(実績)」のPDFか確認してください。")

    # 3) 年を付ける。対象期間の開始年から、月が戻ったら翌年とみなす。
    year, start_month = _start_year(period_text)
    if year is None:
        import datetime
        year = datetime.date.today().year
        start_month = rows[0]["month"]
    out = []
    cur_year, prev_month = year, start_month
    for r in rows:
        if r["month"] < prev_month:
            cur_year += 1
        prev_month = r["month"]
        date = f"{cur_year:04d}-{r['month']:02d}-{r['day']:02d}"
        out.append(calc.clean_row((date, r["dow"], r["start"], r["end"], r["dur"],
                                   r["user"], r["service"], r["note"])))
    return {"person": person, "office": office, "period_text": period_text,
            "rows": out}
