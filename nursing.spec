# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller ビルド定義（onefile / GUI）。3職種で共通。

直接は使わず build_all.py から呼ぶ。build_all.py が
  - _build_profile.py（どの職種のexeか）を書き出し
  - 環境変数 NURSEPAY_EXE_NAME に exe 名を入れて
このファイルでビルドする。
"""
import os
from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = [], [], ["_build_profile"]

# reportlab(フォント)・openpyxl・certifi(CA証明書)・pypdf(実績PDFの読取) を同梱
for pkg in ("reportlab", "openpyxl", "certifi", "pypdf"):
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    ["payroll_app.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name=os.environ["NURSEPAY_EXE_NAME"],
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
