# -*- coding: utf-8 -*-
"""訪問看護の実績の集計・給与計算ロジック。

1行 = 1訪問。列の並びは
  0 日付 / 1 曜日 / 2 開始 / 3 終了 / 4 分 / 5 利用者名 / 6 サービス内容 / 7 備考

計算のきまり（判定はこの順）:
  1. サービス内容に「★」が付く訪問 → 同行。実時間 × 研修時給
  2. 対応表にあるサービス内容     → 対応表の規定時間（0 は実時間）× 訪問時給
     「計上する時間＝実時間」の設定なら、対応表にあっても実時間で数える
  3. 対応表にないサービス内容     → エラー（計算に含めない）
"""
import re
import datetime
from collections import defaultdict, Counter

COL_DATE, COL_DOW, COL_START, COL_END, COL_DUR, COL_USER, COL_SERVICE, COL_NOTE = range(8)
NCOLS = 8
HEADER_ROWS = 1
DOUKOU_MARK = "★"

_SPACES = re.compile(r"[\s　]+")


# ---------------------------------------------------------------- 値の解釈
def _to_minutes_of_day(cell):
    """時刻セルを 0時からの分に変換。datetime.time / 'H:MM' 文字列に対応。"""
    if cell is None:
        return None
    hour = getattr(cell, "hour", None)
    if hour is not None:
        return hour * 60 + getattr(cell, "minute", 0)
    m = re.match(r"\s*(\d{1,2})\s*[:：]\s*(\d{1,2})", str(cell))
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    return None


def parse_minutes(dur_cell, start_cell, end_cell):
    """「分」列から分。無ければ開始/終了から算出。"""
    if dur_cell is not None and str(dur_cell).strip() != "":
        m = re.search(r"(\d+)", str(dur_cell))
        if m:
            return int(m.group(1))
    s = _to_minutes_of_day(start_cell)
    e = _to_minutes_of_day(end_cell)
    if s is not None and e is not None:
        if e < s:
            e += 24 * 60            # 日をまたぐ訪問
        return e - s
    return 0


def parse_date(v):
    """日付セル → datetime.date（年が分からなければ None）。

    取込時は 'YYYY-MM-DD' で保存している。Excel の日付型にも対応。
    """
    if v is None:
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    m = re.search(r"(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})", str(v))
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    return None


def month_of(v):
    """日付セルから (年, 月)。年が無い '10月1日' 形式は (None, 月)。不明は None。"""
    d = parse_date(v)
    if d:
        return d.year, d.month
    m = re.search(r"(\d{1,2})\s*月\s*\d{1,2}\s*日", str(v or ""))
    if m:
        return None, int(m.group(1))
    return None


def normalize_service(service):
    """サービス内容の表記をそろえる（★・空白・改行を除く）。対応表の照合に使う。"""
    s = str(service or "").replace(DOUKOU_MARK, "")
    return _SPACES.sub("", s)


def is_doukou(service):
    return DOUKOU_MARK in str(service or "")


def clean_row(row):
    """行を 8 列の文字列タプルにそろえる（不足は空、余分は捨てる）。"""
    cells = list(row or [])[:NCOLS]
    cells += [""] * (NCOLS - len(cells))
    return tuple("" if c is None else c for c in cells)


def is_data_row(row):
    """日付に数字があり、開始/終了/分のどれかが読める行をデータ行とみなす。"""
    if not row or len(row) <= COL_SERVICE:
        return False
    date = row[COL_DATE]
    if date is None or not re.search(r"\d", str(date)):
        return False
    return (_to_minutes_of_day(row[COL_START]) is not None
            or _to_minutes_of_day(row[COL_END]) is not None
            or parse_minutes(row[COL_DUR], row[COL_START], row[COL_END]) > 0)


def filter_month(rows, year, month):
    """対象月の行だけを残す。戻り: (残した行, 除いた行数)。

    年の分からない行（'10月1日' 形式）は月だけで判定する。
    日付が読めない行は残す（計算側でエラーとして見えるようにするため）。
    """
    kept, dropped = [], 0
    for r in rows:
        ym = month_of(r[COL_DATE]) if r else None
        if ym is None:
            kept.append(r)
            continue
        y, m = ym
        if m == month and (y is None or y == year):
            kept.append(r)
        else:
            dropped += 1
    return kept, dropped


def main_month(rows):
    """行の中でいちばん件数の多い (年, 月)。読めなければ None。"""
    cnt = Counter()
    for r in rows:
        ym = month_of(r[COL_DATE]) if r else None
        if ym and ym[0]:
            cnt[ym] += 1
    return cnt.most_common(1)[0][0] if cnt else None


# ---------------------------------------------------------------- 1訪問の分類
def classify(row, mapping, use_kitei=True):
    """1訪問を分類して時間を決める。空行なら None。

    mapping: {正規化したサービス内容: 規定分}（0 は実時間）
    戻り: dict(minutes=実時間, pay=訪問時給で数える分, training=研修分,
               kitei=規定分 or None, error=bool, kind=表示用の区分)
    """
    if not row or len(row) <= COL_SERVICE or row[COL_DATE] in (None, ""):
        return None
    service = row[COL_SERVICE]
    minutes = parse_minutes(row[COL_DUR], row[COL_START], row[COL_END])
    out = {"minutes": minutes, "pay": 0, "training": 0, "kitei": None,
           "error": False, "kind": ""}
    if is_doukou(service):
        out["training"] = minutes
        out["kind"] = "同行"
        return out
    key = normalize_service(service)
    if key in mapping:
        kitei = int(mapping[key] or 0)
        out["kitei"] = kitei
        if kitei > 0 and use_kitei:
            out["pay"] = kitei
            out["kind"] = f"{kitei}分"
        else:
            out["pay"] = minutes
            out["kind"] = "実時間"
        return out
    out["error"] = True
    out["kind"] = "対応表外"
    return out


# ---------------------------------------------------------------- 交通費
def _kotsu_auto_flags(visit_times, first_included):
    """各訪問が交通費対象かどうかの自動判定（訪問と同じ並びの bool リスト）。

    - 直前の訪問の終了時刻と開始時刻が同じなら、移動が無いものとして対象外。
    - first_included=False のとき、各日の最初の1件を対象外。
    """
    by_day = defaultdict(list)
    for order, (day, s, e) in enumerate(visit_times):
        by_day[day].append((s, e, order))
    flags = [False] * len(visit_times)
    for items in by_day.values():
        items.sort(key=lambda t: (t[0] is None, t[0] if t[0] is not None else 0, t[2]))
        prev_end = None
        for i, (s, e, order) in enumerate(items):
            chargeable = True
            if i == 0 and not first_included:
                chargeable = False
            elif i > 0 and s is not None and prev_end is not None and s == prev_end:
                chargeable = False
            flags[order] = chargeable
            prev_end = e
    return flags


def apply_kotsu_overrides(auto_flags, overrides):
    """自動判定に手動の交通費あり/なしを上書きする（None は自動のまま）。"""
    result = []
    for i, auto in enumerate(auto_flags):
        ov = overrides[i] if overrides and i < len(overrides) else None
        result.append(auto if ov is None else bool(ov))
    return result


def _visit_time(row):
    return (str(row[COL_DATE]), _to_minutes_of_day(row[COL_START]),
            _to_minutes_of_day(row[COL_END]))


# ---------------------------------------------------------------- 集計
def sheet_rows(ws):
    """Excel 1シートから見出し行を除いたデータ行を返す（見出しの行数は自動判定）。"""
    rows = list(ws.iter_rows(values_only=True))
    for i, row in enumerate(rows):
        if is_data_row(row):
            return [clean_row(r) for r in rows[i:]]
    return [clean_row(r) for r in rows[HEADER_ROWS:]]


def calc_rows(rows, mapping, kotsu_first_included=True, kotsu_overrides=None,
              use_kitei=True, unit_min=None):
    """1人分の行を集計する。

    戻り: visits / kotsu_count / visit_min / training_min / total_min /
          slots(規定分→件数) / real_count(実時間で数えた件数) / units / errors
    """
    visit_min = training_min = 0
    visits = 0
    errors = []
    slots = Counter()
    real_count = 0
    units = 0
    visit_times = []
    for row in rows:
        d = classify(row, mapping, use_kitei)
        if d is None:
            continue
        visits += 1
        visit_times.append(_visit_time(row))
        if d["error"]:
            errors.append(normalize_service(row[COL_SERVICE]) or "（空欄）")
            continue
        training_min += d["training"]
        visit_min += d["pay"]
        if d["kind"] == "実時間":
            real_count += 1             # 区分の無いサービス（医療保険等）は単位に数えない
        elif d["pay"]:
            slots[d["pay"]] += 1
            if unit_min:
                units += int(round(d["pay"] / unit_min))
    flags = apply_kotsu_overrides(_kotsu_auto_flags(visit_times, kotsu_first_included),
                                  kotsu_overrides)
    return {
        "visits": visits,
        "kotsu_count": sum(1 for f in flags if f),
        "visit_min": visit_min,
        "training_min": training_min,
        "total_min": visit_min + training_min,
        "slots": dict(sorted(slots.items())),
        "real_count": real_count,
        "units": units,
        "errors": errors,
    }


def visit_details(rows, mapping, kotsu_first_included=True, kotsu_overrides=None,
                  use_kitei=True):
    """訪問を1件ずつ明細化する（労働記録タブ・勤務実績PDF用）。

    classify の各キーに date/dow/user/service/start/end/note と
    visit_index / kotsu_auto / kotsu を付けて返す。
    """
    details, visit_times = [], []
    for row in rows:
        d = classify(row, mapping, use_kitei)
        if d is None:
            continue
        d.update(date=row[COL_DATE], dow=row[COL_DOW], user=row[COL_USER],
                 service=row[COL_SERVICE], start=row[COL_START], end=row[COL_END],
                 note=row[COL_NOTE])
        visit_times.append(_visit_time(row))
        details.append(d)
    auto = _kotsu_auto_flags(visit_times, kotsu_first_included)
    final = apply_kotsu_overrides(auto, kotsu_overrides)
    for i, d in enumerate(details):
        d["visit_index"] = i
        d["kotsu_auto"] = auto[i]
        d["kotsu"] = final[i]
    return details


# ---------------------------------------------------------------- 貼り付け
_RANGE = re.compile(r"(\d{1,2}[:：]\d{2})\s*[～~〜\-－]\s*(\d{1,2}[:：]\d{2})")


def parse_pasted(text):
    """表計算ソフトから貼り付けたタブ区切りテキストを行に変換する。

    次の2つの並びに対応する（見出し行・空行は読み飛ばす）。
      8列: 日付 曜日 開始 終了 分 利用者名 サービス内容 備考
      6列: 日付 サービス提供時間(08:40～09:00) 分 利用者名 サービス内容 備考
    """
    rows = []
    for line in text.splitlines():
        if not line.strip():
            continue
        cells = [c.strip() for c in line.split("\t")]
        if not cells or not re.search(r"\d", cells[0]):
            continue
        if len(cells) > 1 and _RANGE.search(cells[1]):
            m = _RANGE.search(cells[1])
            dow = ""
            dm = re.search(r"[（(]([月火水木金土日])[）)]", cells[0])
            if dm:
                dow = dm.group(1)
            rest = cells[2:] + [""] * 4
            cells = [cells[0], dow, m.group(1), m.group(2)] + rest[:4]
        rows.append(clean_row(cells))
    return rows


# ---------------------------------------------------------------- 金額
def yen(minutes, rate):
    """分 ÷ 60 × 時給（四捨五入）。"""
    return int(round(minutes / 60.0 * float(rate or 0)))


def calc_kaizen(res, rate):
    """処遇改善手当 = 訪問時間 × 円/時（同行の時間は対象外）。"""
    return yen(res["visit_min"], rate)


def calc_shikaku(res, rate):
    """資格手当 = 訪問時間 × 円/時（同行の時間は対象外）。"""
    return yen(res["visit_min"], rate)


def slot_text(res, unit_min=None):
    """時間の内訳を「20分×5 / 30分×12 / 実時間×3」の形で返す。"""
    parts = []
    for m, n in (res.get("slots") or {}).items():
        m = int(m)
        if unit_min:
            parts.append(f"{int(round(m / unit_min))}単位({m}分)×{n}")
        else:
            parts.append(f"{m}分×{n}")
    if res.get("real_count"):
        parts.append(f"実時間×{res['real_count']}")
    return " / ".join(parts)


def build_record(person, period, res, rates, kotsu_unit,
                 shikaku_rate=0, other1=0, other2=0):
    """集計結果と単価から1人分の給与レコードを組み立てる。"""
    amt_training = yen(res["training_min"], rates["training"])
    amt_visit = yen(res["visit_min"], rates["visit"])
    amt_kaizen = calc_kaizen(res, rates.get("kaizen", 0))
    kotsu_count = res.get("kotsu_count", res["visits"])
    kotsu = int(round(kotsu_count * float(kotsu_unit or 0)))
    shikaku_rate = float(shikaku_rate or 0)
    shikaku = calc_shikaku(res, shikaku_rate)
    other1 = int(other1 or 0)
    other2 = int(other2 or 0)
    total = amt_training + amt_visit + amt_kaizen + kotsu + shikaku + other1 + other2
    return {
        "person": person, "period": period,
        "visits": res["visits"], "kotsu_count": kotsu_count,
        "total_min": res["total_min"], "visit_min": res["visit_min"],
        "training_min": res["training_min"], "units": res.get("units", 0),
        "rate_training": float(rates["training"]), "rate_visit": float(rates["visit"]),
        "rate_kaizen": float(rates.get("kaizen", 0)),
        "amt_training": amt_training, "amt_visit": amt_visit, "amt_kaizen": amt_kaizen,
        "kotsu": kotsu, "rate_shikaku": shikaku_rate, "shikaku": shikaku,
        "other1": other1, "other2": other2,
        "total_amount": total,
    }
