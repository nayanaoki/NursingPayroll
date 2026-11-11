# -*- coding: utf-8 -*-
"""アプリのバージョン情報と更新配信設定。

3つのアプリ（理学療法士／看護師／准看護師）は同じバージョン番号で同時に出す。
version.json は1つで、アプリごとの配布URLを apps に持つ:

  {"version": "1.0.0",
   "apps": {"pt":    {"download_url": ".../NursingPayroll_PT.zip"},
            "nurse": {"download_url": ".../NursingPayroll_NS.zip"},
            "lpn":   {"download_url": ".../NursingPayroll_LPN.zip"}}}

リリースのたびに __version__ を上げ、version.json も同じ番号に更新する。
"""

__version__ = "1.0.0"

# 起動時に取得する version.json の URL（3アプリ共通）
UPDATE_URL = "https://raw.githubusercontent.com/nayanaoki/NursingPayroll/main/version.json"
