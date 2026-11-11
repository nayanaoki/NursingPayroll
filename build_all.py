# -*- coding: utf-8 -*-
"""3職種（理学療法士／看護師／准看護師）の exe・zip・インストーラを一度に作る。

使い方:  python build_all.py            … 3つとも
         python build_all.py pt lpn     … 指定した職種だけ

出力（release フォルダ。ファイル名は ASCII。GitHub が日本語名を削るため）:
  NursingPayroll_PT.zip / NursingPayroll_PT_setup.exe
  NursingPayroll_NS.zip / NursingPayroll_NS_setup.exe
  NursingPayroll_LPN.zip / NursingPayroll_LPN_setup.exe
あわせて version.json（3アプリ共通の更新情報）を現在のバージョンで書き直す。

zip の中身は exe 1つだけ（ルート直下）。自動更新(updater)はこれを展開して上書きする。
"""
import os
import sys
import json
import shutil
import zipfile
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
sys.path.insert(0, HERE)

import app_profile                      # noqa: E402
from version import __version__         # noqa: E402

REPO = "nayanaoki/NursingPayroll"
RELEASE_DIR = os.path.join(HERE, "release")
ISCC = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "Inno Setup 6", "ISCC.exe")
# exe は Dropbox の外で作る（Dropbox が同期のために exe を掴み、PyInstaller の書き込みが失敗するため）
WORK = os.path.join(os.environ.get("LOCALAPPDATA", HERE), "NursingPayrollBuild")
DIST = os.path.join(WORK, "dist")


def download_url(asset):
    return f"https://github.com/{REPO}/releases/download/v{__version__}/{asset}.zip"


def build_one(key):
    p = app_profile.PROFILES[key]
    exe_name = p["exe_name"]
    print(f"\n===== {p['job']}（{key}） v{__version__} =====", flush=True)

    with open("_build_profile.py", "w", encoding="utf-8") as f:
        f.write(f'# build_all.py が生成（コミットしない）\nPROFILE = "{key}"\n')
    try:
        env = dict(os.environ, NURSEPAY_EXE_NAME=exe_name)
        subprocess.run([sys.executable, "-m", "PyInstaller", "--noconfirm",
                        "--workpath", os.path.join(WORK, "build", key),
                        "--distpath", DIST, "nursing.spec"],
                       check=True, env=env)
    finally:
        os.remove("_build_profile.py")

    exe = os.path.join(DIST, exe_name + ".exe")
    zpath = os.path.join(RELEASE_DIR, p["asset"] + ".zip")
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        z.write(exe, arcname=exe_name + ".exe")
    print("zip:", zpath, flush=True)

    if os.path.exists(ISCC):
        subprocess.run([ISCC, "/Q",
                        f"/DAppName={p['app_name']}",
                        f"/DAppExe={exe_name}.exe",
                        f"/DAppIdStr={p['app_id']}",
                        f"/DAppVersion={__version__}",
                        f"/DOutName={p['asset']}_setup",
                        f"/DDistDir={DIST}",
                        f"/O{WORK}",                 # 出力も Dropbox の外 → 後でコピー
                        "installer.iss"], check=True)
        setup = p["asset"] + "_setup.exe"
        shutil.copy2(os.path.join(WORK, setup), os.path.join(RELEASE_DIR, setup))
        print("setup:", os.path.join(RELEASE_DIR, setup), flush=True)
    else:
        print("Inno Setup が見つからないためインストーラは作りません:", ISCC)


def write_version_json():
    data = {"version": __version__,
            "apps": {k: {"name": p["job"], "download_url": download_url(p["asset"])}
                     for k, p in app_profile.PROFILES.items()}}
    with open("version.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print("version.json:", data["version"])


def main(keys):
    os.makedirs(RELEASE_DIR, exist_ok=True)
    for key in keys:
        build_one(key)
    write_version_json()
    shutil.rmtree(os.path.join(WORK, "build"), ignore_errors=True)


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a in app_profile.PROFILES]
    main(args or list(app_profile.PROFILES))
