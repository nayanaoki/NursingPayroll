# -*- coding: utf-8 -*-
"""起動時の最新版確認とワンクリック更新。

構成:
  - 配信側に version.json を置く  {"version": "1.1.0", "download_url": "...zip"}
  - 起動時に version.json を取得し、現在バージョンと比較
  - 新しければ「新しいバージョンがあります」を表示
  - 「更新する」で zip をダウンロード → 終了待ち → 上書き → 再起動

更新の上書き処理は日本語パスでも確実に動くよう PowerShell ヘルパーで行う。
配布版（PyInstaller onefile）でのみ自動更新を実行できる。
"""
import os
import re
import ssl
import sys
import json
import tempfile
import threading
import subprocess
import urllib.request
import tkinter as tk
from tkinter import ttk, messagebox

import app_paths
import app_profile
from version import __version__, UPDATE_URL

_USER_AGENT = "NursingPayroll-Updater/1.0"


# ----------------------------------------------------------- SSL
def _ssl_context():
    """検証用 SSL コンテキスト。OS 証明書ストア＋certifi バンドルを併用する。

    端末のルート証明書が古い/不足していても検証できるよう certifi を追加し、
    社内プロキシ等の独自ルートは OS ストア(既定)側でカバーする。
    どちらも欠けて CERTIFICATE_VERIFY_FAILED になる配布端末への対策。
    """
    ctx = ssl.create_default_context()  # OS 既定のルート証明書を読み込む
    try:
        import certifi
        ctx.load_verify_locations(cafile=certifi.where())  # 公開ルートを補強
    except Exception:
        pass
    return ctx


# ----------------------------------------------------------- バージョン比較
def _parse(v):
    """'1.10.2' → (1, 10, 2)。数値以外は無視して安全に比較。"""
    return tuple(int(x) for x in re.findall(r"\d+", str(v))) or (0,)


def is_newer(latest, current):
    return _parse(latest) > _parse(current)


# ----------------------------------------------------------- 取得
def fetch_latest(timeout=8):
    """version.json を取得して dict で返す。"""
    req = urllib.request.Request(UPDATE_URL, headers={"User-Agent": _USER_AGENT})
    with urllib.request.urlopen(req, timeout=timeout, context=_ssl_context()) as r:
        return json.loads(r.read().decode("utf-8"))


def download_url_for(data, key=None):
    """version.json から、このアプリ（職種）の配布URLを取り出す。

    3アプリで version.json を共有するため、URL は apps.<職種キー> に分けて置く。
    """
    key = key or app_profile.KEY
    app = (data.get("apps") or {}).get(key) or {}
    return str(app.get("download_url", "")).strip()


# ----------------------------------------------------------- 起動時チェック
def check_for_updates(parent, silent=True):
    """非同期で最新版を確認。silent=True なら更新がある時だけ通知する。"""
    def work():
        try:
            data = fetch_latest()
        except Exception as ex:
            if not silent:
                parent.after(0, lambda e=ex: messagebox.showwarning(
                    "更新確認", f"最新版情報を取得できませんでした。\n\n{e}"))
            return
        latest = str(data.get("version", "")).strip()
        url = download_url_for(data)
        if latest and is_newer(latest, __version__):
            parent.after(0, lambda: _prompt(parent, latest, url))
        elif not silent:
            parent.after(0, lambda: messagebox.showinfo(
                "更新確認", f"最新版を利用中です。（v{__version__}）"))

    threading.Thread(target=work, daemon=True).start()


def _prompt(parent, latest, url):
    msg = ("新しいバージョンがあります\n\n"
           f"現在：{__version__}\n"
           f"最新：{latest}\n\n"
           "今すぐ更新しますか？")
    if not messagebox.askyesno("更新の確認", msg):
        return
    if not url:
        messagebox.showwarning("更新", "更新ファイルのURLが設定されていません。")
        return
    apply_update(parent, url)


# ----------------------------------------------------------- 更新適用
def apply_update(parent, download_url):
    """zip をダウンロードし、再起動ヘルパー経由で上書き更新する。"""
    if not app_paths.is_frozen():
        messagebox.showinfo(
            "更新", "開発モード（スクリプト実行）では自動更新を行いません。\n"
                    "配布版（インストール済みアプリ）でご利用ください。")
        return

    tmp = tempfile.mkdtemp(prefix=f"nursepay_{app_profile.KEY}_upd_")
    zip_path = os.path.join(tmp, "update.zip")

    win, bar, lbl = _progress_window(parent)

    def work():
        try:
            _download(download_url, zip_path,
                      lambda p: parent.after(0, lambda: _set_progress(bar, lbl, p)))
            extract_dir = os.path.join(tmp, "extract")
            _extract(zip_path, extract_dir)
            src = _flatten_single_dir(extract_dir)
            _launch_replacer(src, app_paths.app_dir(), tmp)
        except Exception as ex:
            parent.after(0, lambda e=ex: (_safe_destroy(win),
                         messagebox.showerror("更新エラー", f"更新に失敗しました。\n\n{e}")))
            return
        # ヘルパー起動に成功 → 本体を終了（ヘルパーが上書き後に再起動）
        parent.after(0, lambda: _quit_for_update(parent, win))

    threading.Thread(target=work, daemon=True).start()


def _download(url, dest, on_progress):
    req = urllib.request.Request(url, headers={"User-Agent": _USER_AGENT})
    with urllib.request.urlopen(req, timeout=30, context=_ssl_context()) as r, open(dest, "wb") as f:
        total = int(r.headers.get("Content-Length", 0) or 0)
        read = 0
        while True:
            chunk = r.read(64 * 1024)
            if not chunk:
                break
            f.write(chunk)
            read += len(chunk)
            on_progress(read / total * 100 if total else -1)


def _extract(zip_path, dest):
    import zipfile
    with zipfile.ZipFile(zip_path) as z:
        z.extractall(dest)


def _flatten_single_dir(extract_dir):
    """zip 内が単一フォルダで包まれている場合はその中身を実体とする。"""
    entries = os.listdir(extract_dir)
    if len(entries) == 1:
        only = os.path.join(extract_dir, entries[0])
        if os.path.isdir(only):
            return only
    return extract_dir


def _launch_replacer(src, install_dir, tmp_root):
    """PowerShell ヘルパーを起動：本体終了待ち → 上書き(リトライ) → 再起動。

    配布版(onefile)は終了直後も一瞬 exe がロックされるため、解放されるまで
    リトライしながら上書きする。ヘルパーは本体のジョブから離脱させて、本体終了で
    巻き込み終了しないようにする。経過は %APPDATA%\\給与自動計算\\update_log.txt に記録。
    """
    exe = os.path.basename(sys.executable)
    pid = os.getpid()
    log_path = os.path.join(app_paths.data_dir(), "update_log.txt")
    ps_path = os.path.join(tmp_root, "apply_update.ps1")
    script = f"""$ErrorActionPreference = 'Stop'
$log = '{log_path}'
function Log($m) {{ ("{{0:o}}  {{1}}" -f (Get-Date), $m) | Out-File -FilePath $log -Append -Encoding utf8 }}
try {{
    Log "ヘルパー開始 (本体PID={pid} の終了を待機)"
    try {{ Wait-Process -Id {pid} -Timeout 120 }} catch {{}}
    $dest = Join-Path '{install_dir}' '{exe}'
    $copied = $false
    for ($i = 0; $i -lt 60; $i++) {{
        try {{
            Copy-Item -Path (Join-Path '{src}' '*') -Destination '{install_dir}' -Recurse -Force
            $copied = $true
            Log "上書きコピー成功 (試行 $i 回目)"
            break
        }} catch {{
            Start-Sleep -Milliseconds 500
        }}
    }}
    if (-not $copied) {{ Log "上書き失敗: exe が解放されませんでした(60回リトライ)" }}
    Start-Process -FilePath $dest
    Log "再起動を実行: $dest"
    Start-Sleep -Seconds 2
    Remove-Item -Path '{tmp_root}' -Recurse -Force -ErrorAction SilentlyContinue
}} catch {{
    Log ("致命的エラー: " + $_.Exception.Message)
}}
"""
    # PowerShell が Unicode として読めるよう BOM 付き UTF-8 で書き出す
    with open(ps_path, "w", encoding="utf-8-sig") as f:
        f.write(script)

    # 【重要】DETACHED_PROCESS(0x00000008) は使わない。配布版(onefile)から
    # DETACHED_PROCESS で powershell を起動すると、環境によってはヘルパーが
    # まったく起動・実行されず(ログの1行目すら書かれない)、exe が置換されない
    # まま本体だけ終了する不具合が出る。窓を出さずに確実に動く CREATE_NO_WINDOW
    # を使う(コンソールは持つが非表示)。BREAKAWAY で本体終了に巻き込まれない。
    NO_WINDOW = 0x08000000         # CREATE_NO_WINDOW (窓を出さずコンソールは持つ)
    NEW_GROUP = 0x00000200         # CREATE_NEW_PROCESS_GROUP
    BREAKAWAY = 0x01000000         # CREATE_BREAKAWAY_FROM_JOB (本体終了で巻き込まれない)
    args = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
            "-WindowStyle", "Hidden", "-File", ps_path]
    try:
        subprocess.Popen(args, creationflags=NO_WINDOW | NEW_GROUP | BREAKAWAY,
                         close_fds=True)
    except OSError:
        # ジョブが離脱を許可しない環境ではフラグ無しで再試行
        subprocess.Popen(args, creationflags=NO_WINDOW | NEW_GROUP, close_fds=True)


# ----------------------------------------------------------- 進捗UI
def _progress_window(parent):
    win = tk.Toplevel(parent)
    win.title("更新中")
    win.geometry("360x120")
    win.transient(parent)
    win.resizable(False, False)
    win.grab_set()
    lbl = ttk.Label(win, text="ダウンロード中...")
    lbl.pack(pady=(18, 6))
    bar = ttk.Progressbar(win, length=300, mode="indeterminate")
    bar.pack(pady=6)
    bar.start(12)
    return win, bar, lbl


def _set_progress(bar, lbl, percent):
    if percent is None or percent < 0:
        return
    if str(bar["mode"]) != "determinate":
        bar.stop()
        bar.config(mode="determinate", maximum=100)
    bar["value"] = percent
    lbl.config(text=f"ダウンロード中... {percent:.0f}%")


def _safe_destroy(win):
    try:
        win.grab_release()
        win.destroy()
    except tk.TclError:
        pass


def _quit_for_update(parent, win):
    _safe_destroy(win)
    messagebox.showinfo("更新", "更新を適用します。アプリを再起動します。")
    try:
        parent.destroy()
    finally:
        sys.exit(0)
