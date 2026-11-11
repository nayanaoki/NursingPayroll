# -*- coding: utf-8 -*-
"""実行パス・データ保存先の解決。

PyInstaller でフリーズした配布版でも、開発時のスクリプト実行でも
同じ呼び出しで「書込可能なデータ領域」「同梱リソース」を得られるようにする。
"""
import os
import sys
import shutil

import app_profile

# 職種ごとに別のフォルダ（%APPDATA%\\訪問看護給与_看護師 など）
APP_NAME = app_profile.APP_NAME


def is_frozen():
    """PyInstaller などで exe 化されているか。"""
    return getattr(sys, "frozen", False)


def app_dir():
    """実行ファイル（または本スクリプト）のあるディレクトリ。

    配布版では exe の置かれたインストール先を指す（更新の上書き先）。
    """
    if is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def data_dir():
    """書込可能なユーザーデータ用ディレクトリ（%APPDATA%\\<アプリ名>）。

    インストール先が読取専用でも、更新で上書きされても消えない場所に
    DB を置くために使用する。
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    d = os.path.join(base, APP_NAME)
    os.makedirs(d, exist_ok=True)
    return d


def db_path():
    """payroll.db の正式パス（%APPDATA%）。

    既存の旧配置（実行ファイル同階層）に DB があり、新配置に未作成なら
    自動で移行コピーして既存データを引き継ぐ。
    """
    target = os.path.join(data_dir(), "payroll.db")
    if not os.path.exists(target):
        legacy = os.path.join(app_dir(), "payroll.db")
        if os.path.exists(legacy):
            try:
                shutil.copy2(legacy, target)
            except OSError:
                pass
    return target


def resource_path(rel):
    """同梱リソースへのパス（onefile 展開先 _MEIPASS に対応）。"""
    base = getattr(sys, "_MEIPASS", None) or app_dir()
    return os.path.join(base, rel)
