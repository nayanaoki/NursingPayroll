# -*- coding: utf-8 -*-
"""職種ごとの違い（アプリ名・時間の区切り・資格・既定単価）を1か所にまとめる。

3つのアプリ（理学療法士／看護師／准看護師）は同じコードから作り、
どの職種として動くかだけをここで切り替える。

どの職種で動くかの決め方（上から順に優先）:
  1. ビルド時に生成される _build_profile.py の PROFILE
     （配布exeはこれで固定される。build_all.py が書き出す）
  2. 環境変数 PAYROLL_PROFILE（pt / nurse / lpn）
  3. 起動引数 --profile=pt など（開発時の確認用）
  4. 既定は nurse
"""
import os
import sys

PROFILES = {
    "pt": {
        "key": "pt",
        "job": "理学療法士",
        "app_name": "訪問看護給与_理学療法士",     # %APPDATA% のフォルダ名・インストール先
        "exe_name": "訪問看護給与_理学療法士",     # exe のファイル名（拡張子なし）
        "asset": "NursingPayroll_PT",              # 配布物のASCII名（GitHub が日本語名を削るため）
        "app_id": "{{4C7B1E20-5A31-4D8E-9B02-NURSEPAY0PT1}",
        "subtitle": "訪問リハビリ（20分＝1単位）の実績から給与を組み立てる",
        "qualifications": ["理学療法士", "作業療法士", "言語聴覚士"],
        "unit_min": 20,                            # 1単位の分数（単位数を表示する）
        "time_rule": "20分＝1単位。2単位（40分）が基本です。",
        "rate_visit": "2000",
        "attend_post": "訪問看護（理学療法士）",
    },
    "nurse": {
        "key": "nurse",
        "job": "看護師",
        "app_name": "訪問看護給与_看護師",
        "exe_name": "訪問看護給与_看護師",
        "asset": "NursingPayroll_NS",
        "app_id": "{{4C7B1E20-5A31-4D8E-9B02-NURSEPAY0NS1}",
        "subtitle": "訪問看護（20・30・60・90分）の実績から給与を組み立てる",
        "qualifications": ["看護師", "保健師", "助産師"],
        "unit_min": None,
        "time_rule": "20分・30分・60分・90分の区分で数えます。",
        "rate_visit": "1800",
        "attend_post": "訪問看護（看護師）",
    },
    "lpn": {
        "key": "lpn",
        "job": "准看護師",
        "app_name": "訪問看護給与_准看護師",
        "exe_name": "訪問看護給与_准看護師",
        "asset": "NursingPayroll_LPN",
        "app_id": "{{4C7B1E20-5A31-4D8E-9B02-NURSEPAY0LP1}",
        "subtitle": "訪問看護（20・30・60・90分）の実績から給与を組み立てる",
        "qualifications": ["准看護師"],
        "unit_min": None,
        "time_rule": "20分・30分・60分・90分の区分で数えます。",
        "rate_visit": "1600",
        "attend_post": "訪問看護（准看護師）",
    },
}

DEFAULT_KEY = "nurse"


def _detect():
    try:
        import _build_profile  # noqa: WPS433  ビルド時にだけ存在する
        key = getattr(_build_profile, "PROFILE", None)
        if key in PROFILES:
            return key
    except ImportError:
        pass
    key = os.environ.get("PAYROLL_PROFILE", "").strip().lower()
    if key in PROFILES:
        return key
    for a in sys.argv[1:]:
        if a.startswith("--profile="):
            key = a.split("=", 1)[1].strip().lower()
            if key in PROFILES:
                return key
    return DEFAULT_KEY


KEY = _detect()
P = PROFILES[KEY]

JOB = P["job"]
APP_NAME = P["app_name"]
QUALIFICATIONS = list(P["qualifications"])
UNIT_MIN = P["unit_min"]
