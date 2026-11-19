# -*- coding: utf-8 -*-
"""データ保管 (SQLite) と設定・対応表の管理。

DB は職種ごとに別（%APPDATA%\\訪問看護給与_<職種>\\payroll.db）。
3つのアプリは互いのデータに触れない。
"""
import sqlite3
import datetime
import json

import app_paths
import app_profile as prof
import payroll_calc as calc

DB_PATH = app_paths.db_path()

# サービス内容 → 規定時間（分）。0 は「実時間で数える」。
# 職種ごとに別の表を使う（他職種の行は入れない）。算定の根拠:
#   介護保険 訪問看護費（訪看Ⅰ／予防は予訪看Ⅰ）
#     Ⅰ１=20分未満 Ⅰ２=30分未満 Ⅰ３=30分以上1時間未満 Ⅰ４=1時間以上1時間30分未満
#     → 看護師等は区分の上限（20/30/60/90分）で数える
#     「・准」= 准看護師が行う場合（所定の90/100）。時間区分は看護師と同じ
#     Ⅰ５ = 理学療法士・作業療法士・言語聴覚士の訪問。1回20分（20分＝1単位）
#           (×２)(×３) は続けて2回・3回、「・２超」は1日2回を超える場合（所定の90/100）
#   定期巡回・随時対応型との連携（定期巡回訪看）は月単位の報酬 → 実時間
#   医療保険 訪問看護基本療養費(Ⅰ)・精神科訪問看護基本療養費(Ⅰ)
#     職種（正看／准看／理学療法士等）で区分され、時間では分けない → 実時間
_KUBUN = {"１": 20, "２": 30, "３": 60, "４": 90}
_PREFIXES = ("訪看Ⅰ", "予訪看Ⅰ")


def _nurse_mapping(suffix, iryo):
    m = {}
    for pre in _PREFIXES:
        for no, minutes in _KUBUN.items():
            m[f"{pre}{no}{suffix}"] = minutes
    for s in iryo:
        m[s] = 0
    return m


def _pt_mapping():
    m = {}
    for pre in _PREFIXES:
        m[f"{pre}５"] = 20
        m[f"{pre}５(×２)"] = 40
        m[f"{pre}５(×３)"] = 60
        m[f"{pre}５・２超"] = 20
    m["基本療養費Ⅰ・理学"] = 0
    return m


DEFAULT_MAPPINGS = {
    "nurse": _nurse_mapping("", ("基本療養費Ⅰ・正看", "精神基本療養費Ⅰ・正看",
                                 "定期巡回訪看", "定期巡回訪看・介５")),
    "lpn": _nurse_mapping("・准", ("基本療養費Ⅰ・准看", "精神基本療養費Ⅰ・准看",
                                   "定期巡回訪看", "定期巡回訪看・介５")),
    "pt": _pt_mapping(),
}
DEFAULT_MAPPING = DEFAULT_MAPPINGS[prof.KEY]

# v1.0.0 で3職種に共通で入れていた表（入れ替えのときに「手を加えていない行」を見分ける）
LEGACY_MAPPING = {}
for _pre in _PREFIXES:
    for _no, _m in _KUBUN.items():
        LEGACY_MAPPING[f"{_pre}{_no}"] = _m
        LEGACY_MAPPING[f"{_pre}{_no}・准"] = _m
    LEGACY_MAPPING[f"{_pre}５"] = 20
    LEGACY_MAPPING[f"{_pre}５(×２)"] = 40
    LEGACY_MAPPING[f"{_pre}５(×３)"] = 60
for _s in ("基本療養費Ⅰ・正看", "基本療養費Ⅰ・准看", "基本療養費Ⅰ・理学",
           "精神基本療養費Ⅰ・正看", "精神基本療養費Ⅰ・准看",
           "定期巡回訪看", "定期巡回訪看・介５"):
    LEGACY_MAPPING[_s] = 0

# 対応表の版。上げたときだけ _migrate_mapping が一度動く（手で直した行は残す）。
MAPPING_VERSION = 2

QUALIFICATIONS = list(prof.QUALIFICATIONS)


def shikaku_key(i):
    """資格の並び順 → 設定キー（資格名を変えてもキーは変わらない）。"""
    return f"shikaku_q{i}"


DEFAULT_SETTINGS = {
    "rate_training": "1100",        # 研修時給（★同行）
    "rate_visit": prof.P["rate_visit"],  # 訪問時給
    "kotsu_unit": "100",            # 交通費 1件あたり
    "kotsu_first_included": "1",    # 1日の最初の1件を交通費に含める(1)/含めない(0)
    "rate_kaizen": "0",             # 処遇改善手当 円/時（訪問時間に対して）
    "default_shikaku": "0",         # 資格手当 初期値・未登録者 円/時
    "use_kitei": "1",               # 計上する時間: 1=対応表の規定時間 / 0=実時間
    "attend_start": "8:30",
    "attend_end": "17:30",
    "attend_shotei": "8:00",
    "attend_post": prof.P["attend_post"],
    "attend_employ": "パートスタッフ",
}
for _i, _q in enumerate(QUALIFICATIONS):
    DEFAULT_SETTINGS[shikaku_key(_i)] = "0"

RECORD_COLS = ["person", "period", "visits", "kotsu_count", "total_min", "visit_min",
               "training_min", "units", "rate_training", "rate_visit", "rate_kaizen",
               "amt_training", "amt_visit", "amt_kaizen", "kotsu", "rate_shikaku",
               "shikaku", "other1", "other2", "total_amount", "created_at"]


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def _now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def init_db():
    conn = connect()
    c = conn.cursor()
    c.execute("CREATE TABLE IF NOT EXISTS settings(key TEXT PRIMARY KEY, value TEXT)")
    c.execute("CREATE TABLE IF NOT EXISTS mapping(service TEXT PRIMARY KEY, minutes INTEGER)")
    c.execute("""CREATE TABLE IF NOT EXISTS employees(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        mgmt_no TEXT, name TEXT NOT NULL, qualification TEXT, created_at TEXT)""")
    c.execute("""CREATE TABLE IF NOT EXISTS work_data(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person TEXT, period TEXT, rows_json TEXT, kotsu_flags TEXT,
        created_at TEXT, updated_at TEXT, UNIQUE(person, period))""")
    c.execute(f"""CREATE TABLE IF NOT EXISTS records(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        person TEXT, period TEXT, visits INTEGER, kotsu_count INTEGER,
        total_min INTEGER, visit_min INTEGER, training_min INTEGER, units INTEGER,
        rate_training REAL, rate_visit REAL, rate_kaizen REAL,
        amt_training INTEGER, amt_visit INTEGER, amt_kaizen INTEGER,
        kotsu INTEGER, rate_shikaku REAL, shikaku INTEGER,
        other1 INTEGER, other2 INTEGER, total_amount INTEGER, created_at TEXT)""")
    for k, v in DEFAULT_SETTINGS.items():
        c.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (k, v))
    if c.execute("SELECT COUNT(*) FROM mapping").fetchone()[0] == 0:
        for s, m in DEFAULT_MAPPING.items():
            c.execute("INSERT INTO mapping(service,minutes) VALUES(?,?)", (s, m))
        _set_mapping_version(c)
    else:
        _migrate_mapping(c)
    conn.commit()
    conn.close()


def _mapping_version(c):
    r = c.execute("SELECT value FROM settings WHERE key='mapping_version'").fetchone()
    try:
        return int(r[0]) if r else 1
    except (TypeError, ValueError):
        return 1


def _set_mapping_version(c):
    c.execute("INSERT INTO settings(key,value) VALUES('mapping_version',?) "
              "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (str(MAPPING_VERSION),))


def _migrate_mapping(c):
    """v1.0.0 の共通表 → 職種別の表へ一度だけ入れ替える。

    - 旧既定のまま（名前も分数も同じ）で、この職種の表に無い行 → 消す
    - この職種の表にあって、まだ無い行 → 足す
    - 手で分数を変えた行・手で足した行 → そのまま残す
    版を記録するので、2回目以降の起動・次の更新では何もしない。
    """
    if _mapping_version(c) >= MAPPING_VERSION:
        return
    have = {r[0]: int(r[1] or 0) for r in
            c.execute("SELECT service,minutes FROM mapping").fetchall()}
    for s, m in have.items():
        if s not in DEFAULT_MAPPING and LEGACY_MAPPING.get(s) == m:
            c.execute("DELETE FROM mapping WHERE service=?", (s,))
    for s, m in DEFAULT_MAPPING.items():
        if s not in have:
            c.execute("INSERT INTO mapping(service,minutes) VALUES(?,?)", (s, m))
    _set_mapping_version(c)


# --- 設定 ---
def get_settings():
    conn = connect()
    rows = conn.execute("SELECT key,value FROM settings").fetchall()
    conn.close()
    d = dict(DEFAULT_SETTINGS)
    d.update({r["key"]: r["value"] for r in rows})
    return d


def save_settings(d):
    conn = connect()
    for k, v in d.items():
        conn.execute("INSERT INTO settings(key,value) VALUES(?,?) "
                     "ON CONFLICT(key) DO UPDATE SET value=excluded.value", (k, str(v)))
    conn.commit()
    conn.close()


# --- 対応表 ---
def get_mapping():
    """{正規化したサービス内容: 規定分}。照合は normalize_service 後の文字列で行う。"""
    conn = connect()
    rows = conn.execute("SELECT service,minutes FROM mapping ORDER BY rowid").fetchall()
    conn.close()
    return {calc.normalize_service(r["service"]): int(r["minutes"] or 0) for r in rows}


def get_mapping_rows():
    """画面表示用（登録した表記のまま）。"""
    conn = connect()
    rows = conn.execute("SELECT service,minutes FROM mapping ORDER BY rowid").fetchall()
    conn.close()
    return [(r["service"], int(r["minutes"] or 0)) for r in rows]


def save_mapping(mapping):
    """mapping: {service: minutes} で全置換。"""
    conn = connect()
    conn.execute("DELETE FROM mapping")
    for s, m in mapping.items():
        s = calc.normalize_service(s)
        if s:
            conn.execute("INSERT OR REPLACE INTO mapping(service,minutes) VALUES(?,?)",
                         (s, int(m or 0)))
    conn.commit()
    conn.close()


# --- 記録 ---
def save_record(rec):
    rec = dict(rec)
    rec.setdefault("created_at", _now())
    conn = connect()
    conn.execute(
        f"INSERT INTO records({','.join(RECORD_COLS)}) "
        f"VALUES({','.join('?' * len(RECORD_COLS))})",
        [rec.get(c) for c in RECORD_COLS])
    conn.commit()
    rid = conn.execute("SELECT last_insert_rowid()").fetchone()[0]
    conn.close()
    return rid


def list_records():
    conn = connect()
    rows = conn.execute("SELECT * FROM records ORDER BY created_at DESC, id DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_record(rid):
    conn = connect()
    r = conn.execute("SELECT * FROM records WHERE id=?", (rid,)).fetchone()
    conn.close()
    return dict(r) if r else None


def delete_record(rid):
    conn = connect()
    conn.execute("DELETE FROM records WHERE id=?", (rid,))
    conn.commit()
    conn.close()


# --- 従業員 ---
def list_employees():
    conn = connect()
    rows = conn.execute("SELECT id,mgmt_no,name,qualification FROM employees "
                        "ORDER BY mgmt_no, id").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_employee(mgmt_no, name, qualification):
    conn = connect()
    conn.execute("INSERT INTO employees(mgmt_no,name,qualification,created_at) "
                 "VALUES(?,?,?,?)", (mgmt_no, name, qualification, _now()))
    conn.commit()
    conn.close()


def update_employee(eid, mgmt_no, name, qualification):
    conn = connect()
    conn.execute("UPDATE employees SET mgmt_no=?, name=?, qualification=? WHERE id=?",
                 (mgmt_no, name, qualification, eid))
    conn.commit()
    conn.close()


def delete_employee(eid):
    conn = connect()
    conn.execute("DELETE FROM employees WHERE id=?", (eid,))
    conn.commit()
    conn.close()


def _name_key(name):
    return "".join(str(name or "").split())


def get_employee_by_name(name):
    """氏名で従業員を1件取得（空白の有無は無視。PDF は「山田 太郎」と空白入り）。"""
    key = _name_key(name)
    if not key:
        return None
    for e in list_employees():
        if _name_key(e["name"]) == key:
            return e
    return None


def _to_float(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


# --- 取込データ（読み込んだ実績の生データ） ---
def save_work_data(person, period, rows):
    """取込データを person+period 単位で保存（上書き）。交通費の手動設定は残す。"""
    norm = [["" if c is None else str(c) for c in row] for row in rows]
    payload = json.dumps(norm, ensure_ascii=False)
    now = _now()
    conn = connect()
    conn.execute(
        "INSERT INTO work_data(person,period,rows_json,created_at,updated_at) "
        "VALUES(?,?,?,?,?) ON CONFLICT(person,period) DO UPDATE SET "
        "rows_json=excluded.rows_json, updated_at=excluded.updated_at",
        (person, period, payload, now, now))
    conn.commit()
    conn.close()


def _row_to_work_data(r):
    d = dict(r)
    try:
        d["rows"] = [calc.clean_row(x) for x in json.loads(d.pop("rows_json") or "[]")]
    except (ValueError, TypeError):
        d["rows"] = []
    try:
        flags = json.loads(d.pop("kotsu_flags") or "null")
    except (ValueError, TypeError):
        flags = None
    d["kotsu_flags"] = flags if isinstance(flags, list) else None
    return d


def list_work_data():
    conn = connect()
    rows = conn.execute(
        "SELECT id,person,period,rows_json,kotsu_flags,updated_at FROM work_data "
        "ORDER BY updated_at DESC, id DESC").fetchall()
    conn.close()
    return [_row_to_work_data(r) for r in rows]


def get_work_data(wid):
    conn = connect()
    r = conn.execute(
        "SELECT id,person,period,rows_json,kotsu_flags,updated_at FROM work_data "
        "WHERE id=?", (wid,)).fetchone()
    conn.close()
    return _row_to_work_data(r) if r else None


def get_kotsu_flags(person, period):
    conn = connect()
    r = conn.execute("SELECT kotsu_flags FROM work_data WHERE person=? AND period=?",
                     (person, period)).fetchone()
    conn.close()
    if not r or not r["kotsu_flags"]:
        return None
    try:
        flags = json.loads(r["kotsu_flags"])
    except (ValueError, TypeError):
        return None
    return flags if isinstance(flags, list) else None


def save_kotsu_flags(person, period, flags):
    conn = connect()
    conn.execute("UPDATE work_data SET kotsu_flags=? WHERE person=? AND period=?",
                 (json.dumps(flags, ensure_ascii=False), person, period))
    conn.commit()
    conn.close()


def delete_work_data(wid):
    conn = connect()
    conn.execute("DELETE FROM work_data WHERE id=?", (wid,))
    conn.commit()
    conn.close()


def shikaku_rate_for(qualification, settings=None):
    """資格に応じた資格手当(円/時)。未設定(0)や未登録なら初期値を使う。"""
    settings = settings or get_settings()
    if qualification in QUALIFICATIONS:
        rate = _to_float(settings.get(shikaku_key(QUALIFICATIONS.index(qualification))))
        if rate > 0:
            return rate
    return _to_float(settings.get("default_shikaku", 0))
