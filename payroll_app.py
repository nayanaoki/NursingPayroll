# -*- coding: utf-8 -*-
"""
訪問看護 給与計算（理学療法士／看護師／准看護師 の3アプリ共通コード）
  - 職員別スケジュール表(実績)PDF・Excel・貼り付けから読込 → 人物ごとに給与計算
  - 計算記録を人物ごとに保管 / 全員の総計
  - 結果を PDF 出力
  - 対応表(サービス内容 → 規定時間)編集画面
  - その他項目(時給・交通費単価)設定画面
  - 交通費は 交通費単価 × 交通費対象件数（初回除外・連続提供の条件を反映）
どの職種として動くかは app_profile が決める。
"""
import os
import re
import sys
import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import openpyxl
    import reportlab  # noqa: F401
except ImportError as ex:
    root = tk.Tk(); root.withdraw()
    messagebox.showerror("ライブラリ不足",
                         f"必要なライブラリが不足しています: {ex.name}\n\n"
                         "コマンドプロンプトで次を実行してください:\n"
                         "    pip install openpyxl reportlab")
    sys.exit(1)

import app_profile as prof
import payroll_db as db
import payroll_calc as calc
import payroll_pdf as pdf
import schedule_pdf
import theme
import updater
from version import __version__

APP_TITLE = f"訪問看護給与　{prof.JOB}"

FONT = theme.FONT
FONT_B = theme.FONT_B


def parse_int(s, default=0):
    try:
        return int(round(float(str(s).replace(",", "").strip())))
    except (ValueError, AttributeError):
        return default


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        db.init_db()
        self.title(f"{APP_TITLE}  v{__version__}")
        self.geometry("1140x780")
        self.minsize(960, 640)
        theme.apply(self)   # 帳簿・数字型のテーマ（地=クリーム/面=白/主=墨）

        # 読み込んだ実績（人物ごと・対象月で絞る前の全行）: [(氏名, 行の並び), ...]
        self.sources = []
        self.person_tabs = {}   # person -> dict(state)
        self._tab_outer = {}    # inner frame -> outer(tab) frame

        # 対象月はアプリ全体で1つ。ヘッダーで選び、計算・保存の両方が参照する。
        today = datetime.date.today()
        self._period_lock = False
        self.var_period = tk.StringVar(value=today.strftime("%Y年%m月"))
        self.var_year = tk.StringVar(value=str(today.year))
        self.var_month = tk.StringVar(value=f"{today.month:02d}")
        self.var_period.trace_add("write", self._period_to_widgets)

        self._build_header()

        # 操作結果は画面下端の薄い文字ではなく、常に見える帯で返す。
        self.banner = theme.Banner(self)
        self.banner.pack(fill="x", padx=12, pady=(10, 0))

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=12, pady=(8, 12))
        self.main_nb = nb

        # 各タブは縦横スクロール可能な内側フレーム。ウィンドウが小さくても
        # 内容が見切れずスクロールで参照できる。
        self.tab_employees = self._scrollable_tab(nb, "従業員管理")
        self.tab_calc = self._scrollable_tab(nb, "給与計算")
        self.tab_workdata = self._scrollable_tab(nb, "労働記録")
        self.tab_records = self._scrollable_tab(nb, "計算記録")
        self.tab_totals = self._scrollable_tab(nb, "全員総計")
        self.tab_mapping = self._scrollable_tab(nb, "対応表編集")
        self.tab_settings = self._scrollable_tab(nb, "その他項目設定")

        self._build_employees()
        self._build_calc()
        self._build_workdata()
        self._build_records()
        self._build_totals()
        self._build_mapping()
        self._build_settings()

        nb.bind("<<NotebookTabChanged>>", self._on_tab_change)
        self._bind_shortcuts()

        # 起動時に最新版を確認（更新がある時だけ通知）
        self.after(1000, lambda: updater.check_for_updates(self, silent=True))

    # ============================================================ ヘッダー
    def _build_header(self):
        """墨のヘッダーバー。アプリ名と、全画面が参照する「対象月」を置く。

        対象月はここでしか変えられない。以前は給与計算タブの単価欄に自由入力で
        紛れていたため、表記ゆれで別レコードとして保存される事故があった。
        """
        header = tk.Frame(self, bg=theme.INK, height=62)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        tk.Frame(header, bg=theme.GOLD, width=4).pack(side="left", fill="y")

        left = tk.Frame(header, bg=theme.INK)
        left.pack(side="left", padx=16)
        tk.Label(left, text=APP_TITLE, bg=theme.INK, fg=theme.ON_INK,
                 font=theme.FONT_TITLE).pack(anchor="w", pady=(11, 0))
        tk.Label(left, text=prof.P["subtitle"], bg=theme.INK,
                 fg=theme.ON_INK_SUB, font=theme.FONT_SM).pack(anchor="w")

        tk.Label(header, text=f"v{__version__}", bg=theme.INK,
                 fg=theme.ON_INK_SUB, font=theme.FONT_SM).pack(
            side="right", padx=(0, 18))

        per = tk.Frame(header, bg=theme.INK)
        per.pack(side="right", padx=18)
        tk.Label(per, text="対象月", bg=theme.INK, fg=theme.ON_INK_SUB,
                 font=theme.FONT_SM).pack(anchor="w", pady=(9, 1))
        row = tk.Frame(per, bg=theme.INK)
        row.pack(anchor="w")
        y = datetime.date.today().year
        cb_y = ttk.Combobox(row, textvariable=self.var_year, width=6,
                            state="readonly", justify="center",
                            values=[str(v) for v in range(y - 3, y + 2)])
        cb_y.pack(side="left")
        tk.Label(row, text="年", bg=theme.INK, fg=theme.ON_INK_SUB,
                 font=theme.FONT_SM).pack(side="left", padx=(3, 6))
        cb_m = ttk.Combobox(row, textvariable=self.var_month, width=4,
                            state="readonly", justify="center",
                            values=[f"{m:02d}" for m in range(1, 13)])
        cb_m.pack(side="left")
        tk.Label(row, text="月", bg=theme.INK, fg=theme.ON_INK_SUB,
                 font=theme.FONT_SM).pack(side="left", padx=(3, 0))
        for cb in (cb_y, cb_m):
            cb.bind("<<ComboboxSelected>>", self._period_from_widgets)

    # ------------------------------------------------------------ 対象月
    def _period_from_widgets(self, _e=None):
        """ヘッダーの年・月コンボから var_period（"YYYY年MM月"）を組み立てる。

        実績を読み込み済みなら、選んだ月の行だけで計算し直す。
        """
        if self._period_lock:
            return
        self._period_lock = True
        try:
            self.var_period.set(f"{self.var_year.get()}年{self.var_month.get()}月")
        finally:
            self._period_lock = False
        if self.sources:
            self.calculate()

    def _period_to_widgets(self, *_a):
        """他画面から var_period が書き換えられたときにコンボへ反映する。"""
        if self._period_lock:
            return
        text = self.var_period.get()
        m = re.match(r"^\s*(\d{4})\D+(\d{1,2})", text)
        if not m:
            return
        self._period_lock = True
        try:
            self.var_year.set(m.group(1))
            self.var_month.set(f"{int(m.group(2)):02d}")
        finally:
            self._period_lock = False

    # ------------------------------------------------------------ 通知
    def notify(self, text, kind="ok"):
        """操作結果をヘッダー直下の帯で返す（ok/info は自動で消える）。"""
        self.banner.show(text, kind)

    # ------------------------------------------------------------ ショートカット
    def _bind_shortcuts(self):
        """毎日使う操作はキーボードだけで回せるようにする。"""
        self.bind_all("<Control-o>", lambda _e: self.open_file())
        self.bind_all("<F5>", lambda _e: self.calculate())
        self.bind_all("<Control-s>", lambda _e: self.save_all_records())
        self.bind_all("<Control-Key-1>", lambda _e: self._select_tab(self.tab_employees))
        self.bind_all("<Control-Key-2>", lambda _e: self._select_tab(self.tab_calc))
        self.bind_all("<Control-Key-3>", lambda _e: self._select_tab(self.tab_workdata))
        self.bind_all("<Control-Key-4>", lambda _e: self._select_tab(self.tab_records))
        self.bind_all("<Control-Key-5>", lambda _e: self._select_tab(self.tab_totals))
        self.bind_all("<Escape>", lambda _e: self.banner.hide())

    # ------------------------------------------------------------ 共通レイアウト
    @staticmethod
    def _toolbar(parent, pady=(0, 0)):
        """ページ（クリーム地）に置く操作列。左＝主要、右＝補助で並べる。"""
        bar = ttk.Frame(parent, style="Page.TFrame")
        bar.pack(fill="x", pady=pady)
        return bar

    @staticmethod
    def _sortable(tree, numeric_cols=()):
        """見出しクリックで並び替えできるようにする。"""
        state = {"col": None, "desc": False}

        def sort_by(col):
            desc = not state["desc"] if state["col"] == col else False
            state.update(col=col, desc=desc)

            def key(iid):
                v = tree.set(iid, col)
                if col in numeric_cols:
                    digits = re.sub(r"[^0-9.\-]", "", str(v))
                    try:
                        return (0, float(digits))
                    except ValueError:
                        return (1, 0.0)
                return (0, str(v))

            for i, iid in enumerate(sorted(tree.get_children(""), key=key,
                                           reverse=desc)):
                tree.move(iid, "", i)
            theme.stripe(tree)

        for c in tree["columns"]:
            tree.heading(c, command=lambda c=c: sort_by(c))

    # ============================================================ スクロール枠
    def _scrollable_tab(self, nb, text):
        """縦横スクロール可能な内側フレームを持つタブを追加し、内側を返す。"""
        outer = ttk.Frame(nb)
        nb.add(outer, text=text)
        canvas = tk.Canvas(outer, highlightthickness=0, bg=theme.CREAM)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=canvas.xview)

        def _auto(bar, side, fill):
            """必要なときだけスクロールバーを出す（常時表示は視界の雑音になる）。"""
            def setter(lo, hi):
                if float(lo) <= 0.0 and float(hi) >= 1.0:
                    bar.pack_forget()
                elif not bar.winfo_ismapped():
                    bar.pack(side=side, fill=fill, before=canvas)
                bar.set(lo, hi)
            return setter

        canvas.configure(yscrollcommand=_auto(vsb, "right", "y"),
                         xscrollcommand=_auto(hsb, "bottom", "x"))
        canvas.pack(side="left", fill="both", expand=True)

        inner = ttk.Frame(canvas, padding=(18, 16), style="Page.TFrame")
        win = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner(_e):
            canvas.configure(scrollregion=canvas.bbox("all"))
        inner.bind("<Configure>", _on_inner)

        def _on_canvas(e):
            # 内容がキャンバスより小さいときは幅を合わせて余白を無くす
            canvas.itemconfigure(win, width=max(e.width, inner.winfo_reqwidth()))
        canvas.bind("<Configure>", _on_canvas)

        # マウスホイールはカーソルが乗っているキャンバスに対して効かせる
        def _wheel(e):
            canvas.yview_scroll(int(-e.delta / 120), "units")
        canvas.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _wheel))
        canvas.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        self._tab_outer[inner] = outer
        return inner

    def _select_tab(self, inner):
        self.main_nb.select(self._tab_outer[inner])

    # ============================================================ 従業員管理
    def _build_employees(self):
        t = self.tab_employees
        theme.section(
            t, "従業員",
            "氏名は実績PDFの「職員名」と一致させてください（空白の有無は問いません）。"
            "一致すると資格に応じた資格手当が自動で適用されます。").pack(
            anchor="w", fill="x")

        cols = ("id", "mgmt_no", "name", "qualification")
        heads = {"id": "ID", "mgmt_no": "管理番号", "name": "氏名",
                 "qualification": "資格"}
        widths = {"id": 56, "mgmt_no": 130, "name": 220, "qualification": 160}
        box, tv = theme.scrolled_tree(t, cols, height=10)
        theme.setup_tree(tv, numeric=("id",), center=("qualification",),
                         widths=widths, heads=heads, stretch="name")
        box.pack(fill="both", expand=True, pady=(12, 0))
        tv.bind("<<TreeviewSelect>>", self._emp_on_select)
        self.emp_tree = tv
        self._sortable(tv, numeric_cols=("id",))

        card, form = theme.card(t)
        card.pack(fill="x", pady=(12, 0))
        theme.field_head(form, "選択中の従業員を編集、または新しく登録").grid(
            row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        def field(col, label, make):
            """ラベルを小さく上、入力を下に置く1項目分の枠。"""
            box = ttk.Frame(form)
            box.grid(row=1, column=col, sticky="w", padx=(0, 20))
            theme.field_head(box, label).pack(anchor="w")
            make(box).pack(anchor="w", pady=(3, 0))

        self.emp_mgmt = tk.StringVar()
        self.emp_name = tk.StringVar()
        self.emp_qual = tk.StringVar()
        field(0, "管理番号",
              lambda p: ttk.Entry(p, textvariable=self.emp_mgmt, width=14))
        field(1, "氏名（必須）",
              lambda p: ttk.Entry(p, textvariable=self.emp_name, width=20))
        field(2, "資格（必須）",
              lambda p: ttk.Combobox(p, textvariable=self.emp_qual,
                                     values=db.QUALIFICATIONS, width=14,
                                     state="readonly"))

        form.columnconfigure(3, weight=1)      # 余白を右端に寄せる
        bar = ttk.Frame(form)
        bar.grid(row=2, column=0, columnspan=6, sticky="ew", pady=(16, 0))
        ttk.Button(bar, text="この内容で登録", command=self._emp_add,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="選択中を上書き", command=self._emp_update).pack(
            side="left", padx=8)
        ttk.Button(bar, text="入力をクリア", command=self._emp_clear,
                   style="Small.TButton").pack(side="left")
        ttk.Button(bar, text="選択中を削除", command=self._emp_delete,
                   style="Danger.TButton").pack(side="right")
        self._refresh_employees()

    def _refresh_employees(self):
        self.emp_tree.delete(*self.emp_tree.get_children())
        for e in db.list_employees():
            self.emp_tree.insert("", "end", values=(
                e["id"], e["mgmt_no"] or "", e["name"], e["qualification"] or ""))
        theme.stripe(self.emp_tree)

    def _emp_on_select(self, _e=None):
        sel = self.emp_tree.selection()
        if not sel:
            return
        _id, mgmt, name, qual = self.emp_tree.item(sel[0])["values"]
        self.emp_mgmt.set(mgmt)
        self.emp_name.set(name)
        self.emp_qual.set(qual)

    def _emp_form_values(self):
        name = self.emp_name.get().strip()
        qual = self.emp_qual.get().strip()
        if not name:
            messagebox.showwarning("入力不足", "氏名は必須です。")
            return None
        if not qual:
            messagebox.showwarning("入力不足", "資格を選択してください。")
            return None
        return self.emp_mgmt.get().strip(), name, qual

    def _emp_add(self):
        v = self._emp_form_values()
        if not v:
            return
        db.add_employee(*v)
        self._refresh_employees()
        self._refresh_paste_employees()
        self._emp_clear()
        self.notify("従業員を追加しました。")

    def _emp_update(self):
        sel = self.emp_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "更新する従業員を一覧から選択してください。")
            return
        v = self._emp_form_values()
        if not v:
            return
        eid = int(self.emp_tree.item(sel[0])["values"][0])
        db.update_employee(eid, *v)
        self._refresh_employees()
        self._refresh_paste_employees()
        self.notify("従業員情報を更新しました。")

    def _emp_delete(self):
        sel = self.emp_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "削除する従業員を選択してください。")
            return
        eid = int(self.emp_tree.item(sel[0])["values"][0])
        if messagebox.askyesno("確認", "選択した従業員を削除しますか？"):
            db.delete_employee(eid)
            self._refresh_employees()
            self._refresh_paste_employees()
            self._emp_clear()
            self.notify("削除しました。")

    def _emp_clear(self):
        self.emp_mgmt.set("")
        self.emp_name.set("")
        self.emp_qual.set("")

    # ============================================================ 給与計算
    def _build_calc(self):
        t = self.tab_calc
        theme.section(
            t, "給与計算",
            "対象月はヘッダーで選びます。取り込み → 単価の確認 → 計算 の順に進みます。"
        ).pack(anchor="w", fill="x")

        # 手順1・2は「入力条件」。計算が終わったら畳んで、結果を前に出す。
        setup = ttk.Frame(t, style="Page.TFrame")
        self.calc_setup = setup

        # --- 手順1: 取り込み ------------------------------------------
        c1, f1 = theme.card(setup)
        c1.pack(fill="x", pady=(12, 0))
        self._step_head(f1, "1", "実績を取り込む",
                        "「職員別スケジュール表(実績)」のPDFを選びます（1ファイル＝1人。"
                        "複数まとめて選べます）。対象月以外の日付の行は除いて計算します。")
        row = ttk.Frame(f1)
        row.pack(fill="x", pady=(10, 0))
        ttk.Button(row, text="実績PDFを選択  Ctrl+O", command=self.open_file,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(row, text="Excelを選択", command=self.open_excel).pack(
            side="left", padx=(8, 0))
        ttk.Button(row, text="表を貼り付けて追加", command=self.open_paste_dialog).pack(
            side="left", padx=8)
        self.file_label = ttk.Label(row, text="ファイル未選択", style="Faint.TLabel")
        self.file_label.pack(side="left", padx=12)

        # --- 手順2: 単価 ----------------------------------------------
        s = db.get_settings()
        c2, f2 = theme.card(setup)
        c2.pack(fill="x", pady=(12, 0))
        self._step_head(f2, "2", "この月の単価を確かめる",
                        "変更してもここでは保存されません。恒久的に変えるなら"
                        "「その他項目設定」で保存してください。")
        self.var_visit = tk.StringVar(value=s["rate_visit"])
        self.var_train = tk.StringVar(value=s["rate_training"])
        self.var_kotsu = tk.StringVar(value=s["kotsu_unit"])
        self.var_kaizen = tk.StringVar(value=s["rate_kaizen"])
        self.var_shikaku_default = tk.StringVar(value=s["default_shikaku"])

        grid = ttk.Frame(f2)
        grid.pack(fill="x", pady=(12, 0))
        self._rate_in(grid, "訪問時給", self.var_visit, 0, 0, "円/時")
        self._rate_in(grid, "研修時給（★同行）", self.var_train, 1, 0, "円/時")
        self._rate_in(grid, "交通費", self.var_kotsu, 2, 0, "円/件")
        self._rate_in(grid, "処遇改善手当", self.var_kaizen, 0, 1, "円/時")
        self._rate_in(grid, "資格手当の初期値", self.var_shikaku_default, 1, 1, "円/時")
        ttk.Label(f2, text=f"時間の数え方：{prof.P['time_rule']}"
                           "サービス内容ごとの時間は「対応表編集」で決めます。",
                  style="Faint.TLabel").pack(anchor="w", pady=(8, 0))

        # 畳んだときに出す1行サマリ（何の条件で計算したかは常に見えるようにする）
        c0, f0 = theme.card(t, padding=10)
        self.calc_setup_bar = c0
        self.calc_setup_text = ttk.Label(f0, text="", style="Muted.TLabel")
        self.calc_setup_text.pack(side="left")
        ttk.Button(f0, text="入力条件を開く", style="Small.TButton",
                   command=lambda: self._set_calc_setup(True)).pack(side="right")

        # --- 手順3: 計算 ----------------------------------------------
        bar = self._toolbar(t, pady=(14, 0))
        self.calc_bar = bar
        ttk.Button(bar, text="計算する  F5", command=self.calculate,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="全員分を記録保存  Ctrl+S",
                   command=self.save_all_records).pack(side="left", padx=8)
        self._set_calc_setup(True)

        # --- 計算結果（人物ごと）---------------------------------------
        holder = ttk.Frame(t, style="Page.TFrame")
        holder.pack(fill="both", expand=True, pady=(12, 0))
        self.calc_holder = holder

        self.calc_empty = tk.Frame(holder, bg=theme.SURFACE,
                                   highlightbackground=theme.LINE,
                                   highlightcolor=theme.LINE,
                                   highlightthickness=1, bd=0)
        inner = ttk.Frame(self.calc_empty, padding=40)
        inner.pack(fill="both", expand=True)
        ttk.Label(inner, text="まだ計算結果がありません",
                  style="Head.TLabel").pack()
        ttk.Label(inner,
                  text="上の「実績PDFを選択」で職員別スケジュール表(実績)を読み込むか、"
                       "「表を貼り付けて追加」で1人分を追加してください。\n"
                       "以前に取り込んだ分は「労働記録」タブから再読込できます。",
                  style="Muted.TLabel", justify="center").pack(pady=(8, 0))

        self.calc_nb = ttk.Notebook(holder, style="Person.TNotebook")
        self._sync_calc_empty()

    def _step_head(self, parent, no, title, desc=None):
        """カード見出し「① 何をするところか」。手順を画面に出して迷わせない。"""
        head = ttk.Frame(parent)
        head.pack(fill="x")
        badge = tk.Label(head, text=no, bg=theme.INK, fg=theme.ON_INK,
                         font=theme.FONT_SM_B, width=2, height=1)
        badge.pack(side="left", padx=(0, 9))
        ttk.Label(head, text=title, style="Head.TLabel").pack(side="left")
        if desc:
            ttk.Label(parent, text=desc, style="Muted.TLabel",
                      justify="left").pack(anchor="w", pady=(6, 0))

    def _calc_setup_summary(self):
        """畳んだ入力条件の1行表示。"""
        f = self.file_label.cget("text")
        mode = "規定時間" if self._use_kitei() else "実時間"
        return (f"{f}　／　訪問 {self.var_visit.get()}・研修 {self.var_train.get()} 円/時"
                f"（{mode}で計上）　交通費 {self.var_kotsu.get()} 円/件　"
                f"資格手当 {self.var_shikaku_default.get()} 円/時")

    def _set_calc_setup(self, open_):
        """入力条件（手順1・2）の開閉。計算後は畳んで結果を画面の前に出す。"""
        if open_:
            self.calc_setup_bar.pack_forget()
            self.calc_setup.pack(fill="x", before=self.calc_bar)
        else:
            self.calc_setup.pack_forget()
            self.calc_setup_text.config(text=self._calc_setup_summary())
            self.calc_setup_bar.pack(fill="x", pady=(12, 0), before=self.calc_bar)

    def _sync_calc_empty(self):
        """結果の有無で、空状態の表示と入力条件の開閉を切り替える。

        結果が出たら入力条件を畳む。そうしないと、計算後にいちばん見たい金額が
        画面の下に隠れてスクロールが要る。
        """
        if not hasattr(self, "calc_bar"):
            return
        if self.calc_nb.tabs():
            self.calc_empty.pack_forget()
            self.calc_nb.pack(fill="both", expand=True)
            self._set_calc_setup(False)
        else:
            self.calc_nb.pack_forget()
            self.calc_empty.pack(fill="both", expand=True)
            self._set_calc_setup(True)

    def _rate_in(self, parent, label, var, col, row=0, suffix=""):
        """単価の入力欄（ラベルは小さく上、数字は等幅で右揃え）。"""
        box = ttk.Frame(parent)
        box.grid(row=row, column=col, sticky="w", padx=(0, 26),
                 pady=(0, 12) if row == 0 else 0)
        theme.field_head(box, label).pack(anchor="w")
        line = ttk.Frame(box)
        line.pack(anchor="w", pady=(3, 0))
        e = ttk.Entry(line, textvariable=var, width=8, justify="right",
                      style="Num.TEntry")
        e.pack(side="left")
        if suffix:
            ttk.Label(line, text=suffix, style="Muted.TLabel").pack(
                side="left", padx=(5, 0))
        e.bind("<Return>", lambda _e: self.calculate())
        return e

    def open_file(self):
        """実績PDF（職員別スケジュール表）を読み込む。複数選択可・1ファイル＝1人。"""
        paths = filedialog.askopenfilenames(
            title="職員別スケジュール表(実績)のPDFを選択（複数可）",
            filetypes=[("PDF", "*.pdf"), ("すべて", "*.*")])
        if not paths:
            return
        sources, failed = [], []
        for path in paths:
            try:
                d = schedule_pdf.read(path)
            except Exception as ex:
                failed.append(f"{os.path.basename(path)}: {ex}")
                continue
            name = d["person"] or os.path.splitext(os.path.basename(path))[0]
            sources.append((name, d["rows"]))
        if failed:
            messagebox.showwarning("読めなかったファイル", "\n".join(failed))
        if not sources:
            return
        names = "、".join(n for n, _r in sources)
        label = (os.path.basename(paths[0]) if len(paths) == 1
                 else f"PDF {len(sources)}件")
        self._set_sources(sources, f"{label}（{names}）")

    def open_excel(self):
        """Excel を読み込む（1シート＝1人。列は 日付・曜日・開始・終了・分・利用者名・
        サービス内容・備考）。"""
        path = filedialog.askopenfilename(
            title="実績のExcelを選択",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("すべて", "*.*")])
        if not path:
            return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception as ex:
            messagebox.showerror("読込エラー", f"Excelの読込に失敗しました。\n{ex}")
            return
        sources = [(ws.title, calc.sheet_rows(ws)) for ws in wb.worksheets]
        self._set_sources(sources, os.path.basename(path))

    def _set_sources(self, sources, label):
        """読み込んだ実績を保持し、対象月を合わせて計算する。

        実績PDFは「10/1〜11/4」のように月をまたぐことがあるため、いちばん件数の多い
        月を対象月にする（ヘッダーで変えれば、その月の行だけで計算し直せる）。
        """
        self.sources = sources
        self.file_label.config(text=label, style="Strong.TLabel")
        all_rows = [r for _n, rows in sources for r in rows]
        ym = calc.main_month(all_rows)
        if ym:
            self.var_period.set(f"{ym[0]}年{ym[1]:02d}月")
        self.calculate()

    def _kotsu_first_included(self):
        """1日の最初の1件を交通費対象に含めるか（その他項目設定の値）。"""
        return db.get_settings().get("kotsu_first_included", "1") != "0"

    def _use_kitei(self):
        """計上する時間: True=対応表の規定時間 / False=実時間（その他項目設定の値）。"""
        return db.get_settings().get("use_kitei", "1") != "0"

    def _calc(self, rows, flags=None):
        """設定（交通費の条件・計上する時間・単位）を反映して1人分を集計する。"""
        return calc.calc_rows(rows, db.get_mapping(), self._kotsu_first_included(),
                              flags, use_kitei=self._use_kitei(), unit_min=prof.UNIT_MIN)

    def _details(self, rows, flags=None):
        return calc.visit_details(rows, db.get_mapping(), self._kotsu_first_included(),
                                  flags, use_kitei=self._use_kitei())

    def _rates(self):
        return {"training": float(parse_int(self.var_train.get())),
                "visit": float(parse_int(self.var_visit.get())),
                "kaizen": float(parse_int(self.var_kaizen.get()))}

    def _period_ym(self):
        m = re.match(r"^\s*(\d{4})\D+(\d{1,2})", self.var_period.get())
        return (int(m.group(1)), int(m.group(2))) if m else (None, None)

    def calculate(self):
        if not self.sources:
            self.notify("先に「実績PDFを選択」で実績を読み込んでください。", "warn")
            return
        for tab in self.calc_nb.tabs():
            self.calc_nb.forget(tab)
        self.person_tabs.clear()

        rates = self._rates()
        kotsu_unit = parse_int(self.var_kotsu.get())
        period = self.var_period.get()
        year, month = self._period_ym()

        errors, dropped_total, empty = [], 0, []
        for person, all_rows in self.sources:
            rows, dropped = calc.filter_month(all_rows, year, month)
            dropped_total += dropped
            if not rows:
                empty.append(person)
                continue
            # 対象月の生データを保存(氏名+対象月で上書き)。
            # 交通費あり/なしの手動設定は保持されるので、それを反映して計算する。
            db.save_work_data(person, period, rows)
            flags = db.get_kotsu_flags(person, period)
            res = self._calc(rows, flags)
            self._add_person_tab(person, res, rates, kotsu_unit, rows=rows)
            if res["errors"]:
                errors.append(f"[{person}] {', '.join(sorted(set(res['errors'])))}")
        self._refresh_workdata()

        self._sync_calc_empty()
        extra = ""
        if dropped_total:
            extra += f"　対象月以外の {dropped_total} 件は除きました。"
        if empty:
            extra += f"　{'、'.join(empty)} は {period} の実績がありません。"
        if errors:
            self.notify(
                "対応表にないサービスがあります（その分は計算に含めていません）。"
                "「対応表編集」タブで追加してください → " + " / ".join(errors) + extra,
                "alert")
        else:
            self.notify(
                f"{period} の計算が終わりました（{len(self.person_tabs)}名）。{extra}"
                "資格手当・その他を入力してから記録保存・PDF出力できます。",
                "warn" if empty else "ok")

    def _add_person_tab(self, person, res, rates, kotsu_unit, rows=None):
        frame = ttk.Frame(self.calc_nb, padding=(0, 14, 0, 0), style="Page.TFrame")
        self.calc_nb.add(frame, text=f"  {person}  ")
        st = {"person": person, "res": res, "rates": rates, "kotsu_unit": kotsu_unit,
              "rows": rows}
        self.person_tabs[person] = st

        # 従業員登録から資格を照合し、資格手当(円/時)の初期値を決める
        emp = db.get_employee_by_name(person)
        qual = (emp.get("qualification") if emp else None)
        settings = db.get_settings()
        settings["default_shikaku"] = str(parse_int(self.var_shikaku_default.get()))
        shikaku_rate = db.shikaku_rate_for(qual, settings)
        st["qual"] = qual

        # --- 上段: 稼働の要約（左） と 合計支給額（右・この画面の主役）-----
        top = ttk.Frame(frame, style="Page.TFrame")
        top.pack(fill="x")

        c_work, work = theme.card(top)
        c_work.pack(side="left", fill="both", expand=True)
        theme.field_head(work, "稼働").pack(anchor="w")
        stats = ttk.Frame(work)
        stats.pack(anchor="w", pady=(8, 0))

        def stat(col, label, num, unit, sub=None):
            b = ttk.Frame(stats)
            b.grid(row=0, column=col, sticky="nw", padx=(0, 34))
            theme.field_head(b, label).pack(anchor="w")
            line = ttk.Frame(b)
            line.pack(anchor="w", pady=(2, 0))
            ttk.Label(line, text=num, style="NumMid.TLabel").pack(side="left")
            ttk.Label(line, text=unit, style="Muted.TLabel").pack(
                side="left", anchor="s", padx=(3, 0))
            if sub:
                ttk.Label(b, text=sub, style="Faint.TLabel").pack(anchor="w")

        h, m = divmod(res["total_min"], 60)
        stat(0, "合計時間", f"{res['total_min']/60:.2f}", "時間", f"{h}時間{m}分")
        stat(1, "訪問件数", f"{res['visits']}", "件")
        col = 2
        if prof.UNIT_MIN:
            stat(col, "単位数", f"{res.get('units', 0)}", "単位",
                 f"{prof.UNIT_MIN}分＝1単位（同行を除く）")
            col += 1

        qual_text = f"{qual}　資格手当 {int(round(shikaku_rate)):,} 円/時" if qual \
            else "従業員未登録　既定の資格手当を適用"
        qbox = ttk.Frame(stats)
        qbox.grid(row=0, column=col, sticky="nw")
        theme.field_head(qbox, "資格").pack(anchor="w")
        ttk.Label(qbox, text=qual_text,
                  style="Strong.TLabel" if qual else "Warn.TLabel").pack(
            anchor="w", pady=(4, 0))

        # 交通費・按分の適用状況を明示（見て分かるように）
        notes = []
        slots = calc.slot_text(res, prof.UNIT_MIN)
        if slots:
            mode = "規定時間" if self._use_kitei() else "実時間"
            notes.append(f"時間の内訳（{mode}で計上）　{slots}")
        if res["training_min"]:
            notes.append(f"★同行 {res['training_min']} 分は研修時給で計算")
        if res.get("kotsu_count", res["visits"]) != res["visits"]:
            notes.append(
                f"交通費対象 {res.get('kotsu_count')} 件"
                f"（連続提供や初回除外により {res['visits']} 件から調整）")
        if res["errors"]:
            notes.append(f"対応表外 {len(res['errors'])} 件は計算に含めていません："
                         + "、".join(sorted(set(res["errors"]))))
        if notes:
            ttk.Separator(work, orient="horizontal").pack(fill="x", pady=(12, 8))
            for n in notes:
                ttk.Label(work, text=n, style="Info.TLabel",
                          font=theme.FONT_SM).pack(anchor="w")

        c_hero, hero = theme.card(top, padding=20)
        c_hero.pack(side="left", fill="y", padx=(12, 0))
        # 金額はこの画面の主役。縦にも中央へ置いて、視線が最初にここへ来るようにする。
        hbox = ttk.Frame(hero)
        hbox.pack(expand=True)
        theme.field_head(hbox, "合計支給額").pack(anchor="e")
        amount_line = ttk.Frame(hbox)
        amount_line.pack(anchor="e", pady=(6, 0))
        ttk.Label(amount_line, text="¥", style="Muted.TLabel").pack(
            side="left", anchor="s", pady=(0, 6), padx=(0, 5))
        st["lbl_total"] = ttk.Label(amount_line, text="0", style="NumHero.TLabel")
        st["lbl_total"].pack(side="left")
        ttk.Label(hbox, text=f"{person}　{self.var_period.get()}",
                  style="Faint.TLabel").pack(anchor="e", pady=(6, 0))

        # --- 下段: 支給の内訳 --------------------------------------------
        c_pay, pay_card = theme.card(frame)
        c_pay.pack(fill="both", expand=True, pady=(12, 0))
        theme.field_head(pay_card, "支給の内訳").pack(anchor="w", pady=(0, 10))
        pay = ttk.Frame(pay_card)
        pay.pack(fill="x")
        pay.columnconfigure(0, weight=1)
        pay.columnconfigure(1, minsize=120)
        pay.columnconfigure(2, minsize=150)
        for c, head in enumerate(["項目", "時間", "金額"]):
            ttk.Label(pay, text=head, style="FieldHead.TLabel",
                      anchor="w" if c == 0 else "e").grid(
                row=0, column=c, sticky="ew", pady=(0, 6))

        def disp(r, name, minutes, amount):
            ttk.Label(pay, text=name, anchor="w").grid(
                row=r, column=0, sticky="w", pady=3)
            tt = "" if minutes is None else f"{minutes//60}:{minutes%60:02d}"
            ttk.Label(pay, text=tt, style="Num.TLabel", anchor="e").grid(
                row=r, column=1, sticky="e", pady=3)
            lbl = ttk.Label(pay, text=f"{amount:,}", style="NumB.TLabel", anchor="e")
            lbl.grid(row=r, column=2, sticky="e", pady=3)
            return lbl

        amt_v = calc.yen(res["visit_min"], rates["visit"])
        amt_t = calc.yen(res["training_min"], rates["training"])
        amt_k = calc.calc_kaizen(res, rates["kaizen"])
        kotsu_count = res.get("kotsu_count", res["visits"])
        kotsu = kotsu_count * kotsu_unit
        st.update(amt_v=amt_v, amt_t=amt_t, amt_k=amt_k, kotsu=kotsu)

        disp(1, f"訪問時給（@{int(rates['visit']):,}）", res["visit_min"], amt_v)
        disp(2, f"研修時給・★同行（@{int(rates['training']):,}）", res["training_min"], amt_t)
        disp(3, "処遇改善手当", None, amt_k)
        disp(4, f"交通費 ({kotsu_count}件×{kotsu_unit}円)", None, kotsu)

        # 編集可能項目 (資格手当は保有資格に応じた円/時をプリセット・上書き可)
        st["var_shikaku"] = tk.StringVar(value=str(int(round(shikaku_rate))))
        st["var_other1"] = tk.StringVar(value="0")
        st["var_other2"] = tk.StringVar(value="0")
        self._shikaku_row(pay, 6, person)
        self._edit_row(pay, 7, "その他", st["var_other1"], person)
        self._edit_row(pay, 8, "その他2", st["var_other2"], person)

        ttk.Label(pay, text="※ 時間は「時:分」表記。金額の合計は右上に出ます。",
                  style="Faint.TLabel").grid(row=9, column=0, columnspan=3,
                                             sticky="w", pady=(10, 0))

        btns = self._toolbar(frame, pady=(12, 0))
        ttk.Button(btns, text="この人を記録保存",
                   command=lambda p=person: self.save_record(p)).pack(side="left")
        ttk.Button(btns, text="給与明細PDFを出力",
                   command=lambda p=person: self.export_pdf(p)).pack(side="left", padx=8)
        ttk.Button(btns, text="このタブを閉じる", style="Small.TButton",
                   command=lambda p=person: self._remove_person_tab(p)).pack(side="right")

        self._update_total(person)
        self._sync_calc_empty()

    def _edit_row(self, pay, r, name, var, person):
        ttk.Label(pay, text=name, anchor="w").grid(row=r, column=0, sticky="w", pady=3)
        ttk.Label(pay, text="手入力", style="Faint.TLabel", anchor="e").grid(
            row=r, column=1, sticky="e", pady=3)
        e = ttk.Entry(pay, textvariable=var, width=12, justify="right",
                      style="Num.TEntry")
        e.grid(row=r, column=2, sticky="e", pady=3)
        var.trace_add("write", lambda *_a, p=person: self._update_total(p))

    def _shikaku_row(self, pay, r, person):
        """資格手当行: 円/時を入力→稼働時間に応じて金額を自動計算・表示。"""
        st = self.person_tabs[person]
        ttk.Label(pay, text="資格手当", anchor="w").grid(
            row=r, column=0, sticky="w", pady=3)
        rate_frame = ttk.Frame(pay)
        rate_frame.grid(row=r, column=1, sticky="e", pady=3)
        ttk.Entry(rate_frame, textvariable=st["var_shikaku"], width=5,
                  justify="right", style="Num.TEntry").pack(side="left")
        ttk.Label(rate_frame, text="円/時", style="Muted.TLabel").pack(
            side="left", padx=(4, 0))
        st["lbl_shikaku"] = ttk.Label(pay, text="0", style="NumB.TLabel", anchor="e")
        st["lbl_shikaku"].grid(row=r, column=2, sticky="e", pady=3)
        st["var_shikaku"].trace_add("write", lambda *_a, p=person: self._update_total(p))

    def _update_total(self, person):
        st = self.person_tabs[person]
        shikaku = calc.calc_shikaku(st["res"], parse_int(st["var_shikaku"].get()))
        st["amt_shikaku"] = shikaku
        if "lbl_shikaku" in st:
            st["lbl_shikaku"].config(text=f"{shikaku:,}")
        other1 = parse_int(st["var_other1"].get())
        other2 = parse_int(st["var_other2"].get())
        total = (st["amt_v"] + st["amt_t"] + st["amt_k"]
                 + st["kotsu"] + shikaku + other1 + other2)
        st["lbl_total"].config(text=f"{total:,}")

    def _record_from_tab(self, person):
        st = self.person_tabs[person]
        rec = calc.build_record(
            person, self.var_period.get(), st["res"], st["rates"], st["kotsu_unit"],
            shikaku_rate=parse_int(st["var_shikaku"].get()),
            other1=parse_int(st["var_other1"].get()),
            other2=parse_int(st["var_other2"].get()))
        return rec

    def save_record(self, person):
        rec = self._record_from_tab(person)
        db.save_record(rec)
        self.notify(f"{person} の記録を保存しました。")
        self._refresh_records()

    def save_all_records(self):
        if not self.person_tabs:
            return
        for person in self.person_tabs:
            db.save_record(self._record_from_tab(person))
        self.notify(f"全員（{len(self.person_tabs)}名）の記録を保存しました。")
        self._refresh_records()

    def export_pdf(self, person):
        rec = self._record_from_tab(person)
        rec["created_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        default = f"給与明細_{person}_{self.var_period.get()}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(
            title="PDF保存先", defaultextension=".pdf", initialfile=default,
            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        try:
            pdf.export_record_pdf(rec, path)
        except Exception as ex:
            messagebox.showerror("PDF出力エラー", str(ex))
            return
        if messagebox.askyesno("完了", f"PDFを出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    # ------------------------------------------------------------ 貼り付け入力
    def open_paste_dialog(self):
        dlg = tk.Toplevel(self)
        dlg.title("表を貼り付けて追加")
        dlg.geometry("720x600")
        dlg.transient(self)
        dlg.configure(bg=theme.CREAM)

        head = tk.Frame(dlg, bg=theme.INK, height=52)
        head.pack(fill="x")
        head.pack_propagate(False)
        tk.Frame(head, bg=theme.GOLD, width=4).pack(side="left", fill="y")
        tk.Label(head, text="表を貼り付けて追加", bg=theme.INK, fg=theme.ON_INK,
                 font=theme.FONT_H).pack(side="left", padx=14)
        tk.Label(head, text=f"対象月 {self.var_period.get()}", bg=theme.INK,
                 fg=theme.ON_INK_SUB, font=theme.FONT_SM).pack(side="right", padx=16)

        body = ttk.Frame(dlg, style="Page.TFrame", padding=(16, 14))
        body.pack(fill="both", expand=True)

        theme.field_head(body, "だれの実績か").pack(anchor="w")
        row = ttk.Frame(body, style="Page.TFrame")
        row.pack(fill="x", pady=(4, 0))
        self._paste_name = tk.StringVar()
        self._paste_combo = ttk.Combobox(row, textvariable=self._paste_name,
                                         width=28)
        self._paste_combo.pack(side="left")
        ttk.Label(row, text="登録済みから選ぶか、氏名を直接入力できます",
                  style="PageDesc.TLabel").pack(side="left", padx=10)
        self._refresh_paste_employees()

        theme.field_head(body, "勤務実績の表").pack(anchor="w", pady=(16, 0))
        ttk.Label(body, text="表計算ソフトで「日付」から「備考」までの列をコピーし、"
                            "下の欄に貼り付けてください（Ctrl+V）。\n"
                            "列の並び：日付・サービス提供時間(08:40～09:00)・分・利用者名・"
                            "サービス内容・備考\n"
                            "　　または 日付・曜日・開始・終了・分・利用者名・サービス内容・備考。"
                            "見出し行や空行は自動で読み飛ばします。",
                  style="PageDesc.TLabel", justify="left").pack(anchor="w",
                                                                pady=(4, 6))
        txt = tk.Text(dlg, height=18, width=80, bg=theme.SURFACE, fg=theme.BODY,
                      relief="flat", borderwidth=0, highlightthickness=1,
                      highlightbackground=theme.LINE, highlightcolor=theme.GOLD,
                      insertbackground=theme.INK, padx=10, pady=8, font=FONT)
        txt.pack(in_=body, fill="both", expand=True)
        txt.focus_set()

        bar = ttk.Frame(body, style="Page.TFrame")
        bar.pack(fill="x", pady=(14, 0))

        def do_add():
            name = self._paste_name.get().strip()
            if not name:
                messagebox.showwarning("入力不足", "氏名を入力または選択してください。", parent=dlg)
                return
            rows = calc.parse_pasted(txt.get("1.0", "end"))
            if not rows:
                messagebox.showwarning(
                    "データなし", "貼り付けデータから有効な行が読み取れませんでした。", parent=dlg)
                return
            rates = self._rates()
            kotsu_unit = parse_int(self.var_kotsu.get())
            # 貼り付けた生データもそのまま保存（既存の交通費手動設定は保持）
            db.save_work_data(name, self.var_period.get(), rows)
            flags = db.get_kotsu_flags(name, self.var_period.get())
            res = self._calc(rows, flags)
            self._remove_person_tab(name)
            self._add_person_tab(name, res, rates, kotsu_unit, rows=rows)
            self._refresh_workdata()
            self.calc_nb.select(self.calc_nb.tabs()[-1])
            msg = f"{name} を追加しました（{res['visits']}件）。"
            kind = "ok"
            if res["errors"]:
                msg += " 対応表にないサービス: " + ", ".join(sorted(set(res["errors"])))
                kind = "alert"
            self.notify(msg, kind)
            dlg.destroy()

        ttk.Button(bar, text="この内容で追加", command=do_add,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="閉じる", command=dlg.destroy,
                   style="Ghost.TButton").pack(side="left", padx=8)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

    def _refresh_paste_employees(self):
        combo = getattr(self, "_paste_combo", None)
        if combo is None:
            return
        try:
            if not combo.winfo_exists():
                return
            combo["values"] = [e["name"] for e in db.list_employees()]
        except tk.TclError:
            pass

    def _remove_person_tab(self, name):
        if name not in self.person_tabs:
            return
        for tab in self.calc_nb.tabs():
            if self.calc_nb.tab(tab, "text").strip() == name:
                self.calc_nb.forget(tab)
                break
        self.person_tabs.pop(name, None)
        self._sync_calc_empty()

    # ============================================================ 労働記録
    def _build_workdata(self):
        t = self.tab_workdata
        theme.section(
            t, "労働記録",
            "PDF・Excel読込・貼り付けの都度、氏名＋対象月ごとに自動保存されます。"
            "アプリを閉じても残り、勤務実績・個人出勤簿として出力できます。").pack(
            anchor="w", fill="x")

        cols = ("id", "person", "period", "lines", "updated_at")
        heads = {"id": "ID", "person": "氏名", "period": "対象月",
                 "lines": "明細行数", "updated_at": "保存日時"}
        widths = {"id": 56, "person": 150, "period": 120, "lines": 90,
                  "updated_at": 180}
        box, tv = theme.scrolled_tree(t, cols, height=7)
        theme.setup_tree(tv, numeric=("id", "lines"),
                         center=("period", "updated_at"),
                         widths=widths, heads=heads, stretch="person")
        box.pack(fill="x", pady=(12, 0))
        tv.bind("<<TreeviewSelect>>", self._wd_on_select)
        self.wd_tree = tv
        self._sortable(tv, numeric_cols=("id", "lines"))

        bar = self._toolbar(t, pady=(12, 0))
        ttk.Button(bar, text="勤務実績を出力", command=self._wd_export_record,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="個人出勤簿を出力",
                   command=self._wd_export_attendance).pack(side="left", padx=8)
        ttk.Button(bar, text="給与計算へ再読込",
                   command=self._wd_reload_selected).pack(side="left")
        ttk.Button(bar, text="すべて再読込", style="Small.TButton",
                   command=self._wd_reload_all).pack(side="left", padx=8)
        ttk.Button(bar, text="選択を削除", command=self._wd_delete,
                   style="Danger.TButton").pack(side="right")
        ttk.Button(bar, text="再表示", command=self._refresh_workdata,
                   style="Small.TButton").pack(side="right", padx=8)

        # --- 選択した人の訪問明細（1件ずつ交通費のあり／なしを切替）---------
        c_det, det = theme.card(t)
        c_det.pack(fill="both", expand=True, pady=(14, 0))
        theme.field_head(det, "訪問明細（交通費の あり／なし をここで直せます）").pack(
            anchor="w")
        self.wd_detail_info = ttk.Label(
            det, text="上の一覧から氏名を選ぶと、その人の訪問明細を表示します。",
            style="Muted.TLabel")
        self.wd_detail_info.pack(anchor="w", pady=(6, 10))

        dcols = ("no", "date", "dow", "user", "service", "minutes", "kind", "kotsu")
        dheads = {"no": "No.", "date": "日付", "dow": "曜日", "user": "利用者",
                  "service": "サービス内容", "minutes": "実時間", "kind": "計上",
                  "kotsu": "交通費"}
        dwidths = {"no": 48, "date": 100, "dow": 50, "user": 150, "service": 200,
                   "minutes": 70, "kind": 100, "kotsu": 90}
        dbox, dtv = theme.scrolled_tree(det, dcols, height=10)
        theme.setup_tree(dtv, numeric=("no", "minutes"),
                         center=("date", "dow", "kind", "kotsu"),
                         widths=dwidths, heads=dheads, stretch="service")
        # 「なし」は異常ではないので赤にしない。手動で直した行だけを目立たせる。
        dtv.tag_configure("on", foreground=theme.OK)
        dtv.tag_configure("off", foreground=theme.MUTED)
        dtv.tag_configure("manual", background=theme.GOLD_PALE)
        dbox.pack(fill="both", expand=True)
        dtv.bind("<Double-1>", lambda _e: self._wd_toggle_kotsu())
        self.wd_detail = dtv

        dbar = ttk.Frame(det)
        dbar.pack(fill="x", pady=(12, 0))
        ttk.Button(dbar, text="交通費 あり／なし を切替",
                   command=self._wd_toggle_kotsu).pack(side="left")
        ttk.Button(dbar, text="自動判定に戻す", style="Small.TButton",
                   command=self._wd_reset_kotsu).pack(side="left", padx=8)
        ttk.Label(dbar, text="行をダブルクリックでも切り替わります。"
                            "変更は自動保存され、給与計算とPDFに反映されます。",
                  style="Faint.TLabel").pack(side="left", padx=10)
        self._wd_cur = None
        self._refresh_workdata()

    def _refresh_workdata(self):
        if not hasattr(self, "wd_tree"):
            return
        self.wd_tree.delete(*self.wd_tree.get_children())
        for d in db.list_work_data():
            self.wd_tree.insert("", "end", values=(
                d["id"], d["person"], d["period"] or "", f"{len(d['rows'])}行",
                d["updated_at"]))
        theme.stripe(self.wd_tree)
        self._wd_clear_detail()

    # ------------------------------------------------------------ 明細・交通費切替
    @staticmethod
    def _detail_kind(d):
        """計上の欄：「30分」「実時間 31分」「同行 30分」「対応表外」。"""
        if d.get("error"):
            return "対応表外"
        if d.get("training"):
            return f"同行 {d['training']}分"
        if d.get("kind") == "実時間":
            return f"実時間 {d['pay']}分"
        return d.get("kind") or "-"

    def _wd_clear_detail(self):
        if not hasattr(self, "wd_detail"):
            return
        self.wd_detail.delete(*self.wd_detail.get_children())
        self._wd_cur = None
        self.wd_detail_info.config(
            text="上の一覧から氏名を選ぶと、その人の訪問明細を表示します。",
            style="Muted.TLabel")

    def _wd_on_select(self, _e=None):
        sel = self.wd_tree.selection()
        if not sel:
            return
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        d = db.get_work_data(wid)
        if not d:
            return
        rows = [tuple(r) for r in d["rows"]]
        flags = list(d.get("kotsu_flags") or [])
        self._wd_cur = {"id": wid, "person": d["person"], "period": d.get("period"),
                        "rows": rows, "flags": flags}
        self._wd_render_detail()

    def _wd_render_detail(self):
        cur = self._wd_cur
        if not cur:
            return
        dtv = self.wd_detail
        dtv.delete(*dtv.get_children())
        details = self._details(cur["rows"], cur["flags"])
        charge = 0
        for d in details:
            i = d["visit_index"]
            on = d["kotsu"]
            if on:
                charge += 1
            mark = "あり" if on else "なし"
            tags = ["on" if on else "off"]
            if d["kotsu"] != d["kotsu_auto"]:
                mark += "（手動）"    # 自動判定を上書きした行
                tags.append("manual")
            dd = calc.parse_date(d.get("date"))
            dtv.insert("", "end", iid=str(i), tags=tuple(tags), values=(
                i + 1, f"{dd.month}/{dd.day}" if dd else d.get("date", ""),
                d.get("dow", "") or "",
                d.get("user", "") or "", d.get("service", "") or "",
                d.get("minutes", 0), self._detail_kind(d), mark))
        self.wd_detail_info.config(
            text=f"{cur['person']}　{cur['period'] or ''}　／　"
                 f"訪問 {len(details)} 件のうち交通費の対象は {charge} 件"
                 f"（背景色つきの行は自動判定を手動で変更したもの）",
            style="Muted.TLabel")

    def _wd_selected_visit_index(self):
        if not self._wd_cur:
            messagebox.showinfo("選択なし", "上の一覧から氏名（労働記録）を選択してください。")
            return None
        sel = self.wd_detail.selection()
        if not sel:
            messagebox.showinfo("選択なし", "交通費を切り替える明細行を選択してください。")
            return None
        return int(sel[0])

    def _wd_apply_flag(self, idx, value):
        """訪問 idx の交通費フラグを value(True/False/None)にして保存・再描画。"""
        cur = self._wd_cur
        flags = cur["flags"]
        if len(flags) <= idx:
            flags.extend([None] * (idx + 1 - len(flags)))
        flags[idx] = value
        db.save_kotsu_flags(cur["person"], cur["period"], flags)
        self._wd_render_detail()
        self.wd_detail.selection_set(str(idx))
        self.wd_detail.see(str(idx))

    def _wd_toggle_kotsu(self):
        idx = self._wd_selected_visit_index()
        if idx is None:
            return
        details = self._details(self._wd_cur["rows"], self._wd_cur["flags"])
        current = details[idx]["kotsu"] if idx < len(details) else False
        self._wd_apply_flag(idx, not current)
        self.notify(f"No.{idx + 1} の交通費を「{'あり' if not current else 'なし'}」に変更しました。")

    def _wd_reset_kotsu(self):
        idx = self._wd_selected_visit_index()
        if idx is None:
            return
        self._wd_apply_flag(idx, None)
        self.notify(f"No.{idx + 1} の交通費を自動判定に戻しました。")

    def _wd_load_into_calc(self, d):
        """保存済み取込データ1件を給与計算タブへ再読込する。"""
        rows = [tuple(r) for r in d["rows"]]
        rates = self._rates()
        kotsu_unit = parse_int(self.var_kotsu.get())
        res = self._calc(rows, d.get("kotsu_flags"))
        self._remove_person_tab(d["person"])
        self._add_person_tab(d["person"], res, rates, kotsu_unit, rows=rows)
        return res

    def _wd_reload_selected(self):
        sel = self.wd_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "再読込するデータを選択してください。")
            return
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        d = db.get_work_data(wid)
        if not d:
            return
        if d.get("period"):
            self.var_period.set(d["period"])
        self._wd_load_into_calc(d)
        self._select_tab(self.tab_calc)
        self.calc_nb.select(self.calc_nb.tabs()[-1])
        self.notify(f"{d['person']} の取込データを再読込しました。")

    def _wd_reload_all(self):
        data = db.list_work_data()
        if not data:
            messagebox.showinfo("データなし", "保存された取込データがありません。")
            return
        for tab in self.calc_nb.tabs():
            self.calc_nb.forget(tab)
        self.person_tabs.clear()
        for d in data:
            self._wd_load_into_calc(d)
        self._select_tab(self.tab_calc)
        if self.calc_nb.tabs():
            self.calc_nb.select(self.calc_nb.tabs()[0])
        self.notify(f"保存された取込データ {len(data)} 件を再読込しました。")

    def _wd_delete(self):
        sel = self.wd_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "削除するデータを選択してください。")
            return
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        if messagebox.askyesno("確認", "選択した労働記録を削除しますか？"):
            db.delete_work_data(wid)
            self._refresh_workdata()
            self.notify("削除しました。")

    def _wd_selected(self):
        sel = self.wd_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "出力するデータを一覧から選択してください。")
            return None
        wid = int(self.wd_tree.item(sel[0])["values"][0])
        return db.get_work_data(wid)

    def _wd_export_record(self):
        """選択した労働記録を「勤務実績」書式の PDF で出力（元デザイン準拠）。"""
        d = self._wd_selected()
        if not d:
            return
        default = f"勤務実績_{d['person']}_{d.get('period') or ''}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(
            title="労働記録(勤務実績)の保存先", defaultextension=".pdf",
            initialfile=default, filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        try:
            pdf.export_work_record_pdf(d["person"], d.get("period"), d["rows"],
                                       db.get_mapping(), path,
                                       use_kitei=self._use_kitei())
        except Exception as ex:
            messagebox.showerror("出力エラー", str(ex))
            return
        self.notify(f"労働記録(勤務実績)を出力しました: {path}")
        if messagebox.askyesno("完了", f"労働記録を出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    def _wd_export_attendance(self):
        """選択した労働記録から「個人出勤簿」書式の PDF をダウンロード（元デザイン準拠）。"""
        d = self._wd_selected()
        if not d:
            return
        default = f"個人出勤簿_{d['person']}_{d.get('period') or ''}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(
            title="個人出勤簿の保存先", defaultextension=".pdf",
            initialfile=default, filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        s = db.get_settings()
        try:
            pdf.export_attendance_pdf(
                d["person"], d.get("period"), d["rows"], path,
                start=s.get("attend_start", "8:30"),
                end=s.get("attend_end", "17:30"),
                shotei=s.get("attend_shotei", "8:00"),
                post=s.get("attend_post", prof.P["attend_post"]),
                employ=s.get("attend_employ", "パートスタッフ"))
        except Exception as ex:
            messagebox.showerror("出力エラー", str(ex))
            return
        self.notify(f"個人出勤簿をダウンロードしました: {path}")
        if messagebox.askyesno("完了", f"個人出勤簿を出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    # ============================================================ 計算記録
    def _build_records(self):
        t = self.tab_records
        theme.section(
            t, "計算記録",
            "「記録保存」した計算結果です。見出しをクリックすると並び替えできます。").pack(
            anchor="w", fill="x")

        cols = ("id", "person", "period", "visits", "total_min", "total_amount",
                "created_at")
        heads = {"id": "ID", "person": "氏名", "period": "対象月", "visits": "訪問",
                 "total_min": "勤務時間", "total_amount": "支給額",
                 "created_at": "保存日時"}
        widths = {"id": 56, "person": 140, "period": 110, "visits": 80,
                  "total_min": 100, "total_amount": 130, "created_at": 170}
        box, tv = theme.scrolled_tree(t, cols, height=14)
        theme.setup_tree(tv, numeric=("id", "visits", "total_min", "total_amount"),
                         center=("period", "created_at"),
                         widths=widths, heads=heads, stretch="person")
        box.pack(fill="both", expand=True, pady=(12, 0))
        self.rec_tree = tv
        self._sortable(tv, numeric_cols=("id", "visits", "total_min",
                                         "total_amount"))

        bar = self._toolbar(t, pady=(12, 0))
        ttk.Button(bar, text="給与明細PDFを出力", command=self._rec_pdf,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="再読み込み", command=self._refresh_records,
                   style="Small.TButton").pack(side="left", padx=8)
        ttk.Button(bar, text="選択を削除", command=self._rec_delete,
                   style="Danger.TButton").pack(side="right")
        self._refresh_records()

    def _refresh_records(self):
        if not hasattr(self, "rec_tree"):
            return
        self.rec_tree.delete(*self.rec_tree.get_children())
        for r in db.list_records():
            self.rec_tree.insert("", "end", values=(
                r["id"], r["person"], r["period"], f"{r['visits']}件",
                f"{r['total_min']/60:.2f}h", f"{r['total_amount']:,}円",
                r["created_at"]))
        theme.stripe(self.rec_tree)

    def _selected_rec_id(self):
        sel = self.rec_tree.selection()
        if not sel:
            messagebox.showinfo("選択なし", "記録を選択してください。")
            return None
        return int(self.rec_tree.item(sel[0])["values"][0])

    def _rec_pdf(self):
        rid = self._selected_rec_id()
        if rid is None:
            return
        rec = db.get_record(rid)
        default = f"給与明細_{rec['person']}_{rec.get('period') or ''}.pdf".replace("/", "-")
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=default,
                                            filetypes=[("PDF", "*.pdf")])
        if not path:
            return
        pdf.export_record_pdf(rec, path)
        if messagebox.askyesno("完了", f"PDFを出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    def _rec_delete(self):
        rid = self._selected_rec_id()
        if rid is None:
            return
        if messagebox.askyesno("確認", "選択した記録を削除しますか？"):
            db.delete_record(rid)
            self._refresh_records()

    # ============================================================ 全員総計
    def _build_totals(self):
        t = self.tab_totals
        theme.section(t, "全員総計",
                      "この画面の主役は右の合計額です。").pack(anchor="w", fill="x")

        c_top, top = theme.card(t)
        c_top.pack(fill="x", pady=(12, 0))
        left = ttk.Frame(top)
        left.pack(side="left", anchor="w")
        theme.field_head(left, "集計対象").pack(anchor="w")
        self.total_mode = tk.StringVar(value="current")
        modes = ttk.Frame(left)
        modes.pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(modes, text="今回の計算結果", variable=self.total_mode,
                        value="current", command=self._refresh_totals).pack(
            side="left", padx=(0, 16))
        ttk.Radiobutton(modes, text="保存済み記録すべて", variable=self.total_mode,
                        value="records", command=self._refresh_totals).pack(
            side="left")

        right = ttk.Frame(top)
        right.pack(side="right", anchor="e")
        theme.field_head(right, "総支給額").pack(anchor="e")
        line = ttk.Frame(right)
        line.pack(anchor="e", pady=(4, 0))
        ttk.Label(line, text="¥", style="Muted.TLabel").pack(
            side="left", anchor="s", pady=(0, 5), padx=(0, 4))
        self.tot_amount = ttk.Label(line, text="0", style="NumHero.TLabel")
        self.tot_amount.pack(side="left")
        self.tot_count = ttk.Label(right, text="", style="Faint.TLabel")
        self.tot_count.pack(anchor="e")

        bar = self._toolbar(t, pady=(12, 0))
        ttk.Button(bar, text="総計をPDF出力", command=self._totals_pdf,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="表を更新", command=self._refresh_totals,
                   style="Small.TButton").pack(side="left", padx=8)

        cols = ("person", "visits", "total_min", "amt_visit", "amt_training",
                "kotsu", "extra", "total_amount")
        heads = {"person": "氏名", "visits": "訪問", "total_min": "合計時間",
                 "amt_visit": "訪問給", "amt_training": "研修(同行)",
                 "kotsu": "交通費", "extra": "手当等", "total_amount": "支給額"}
        widths = {"person": 130, "visits": 70, "total_min": 90,
                  "amt_visit": 100, "amt_training": 100,
                  "kotsu": 90, "extra": 90, "total_amount": 120}
        box, tv = theme.scrolled_tree(t, cols, height=14)
        theme.setup_tree(tv, numeric=[c for c in cols if c != "person"],
                         widths=widths, heads=heads, stretch="person")
        box.pack(fill="both", expand=True, pady=(12, 0))
        self.tot_tree = tv

    def _collect_total_records(self):
        if self.total_mode.get() == "records":
            return db.list_records()
        return [self._record_from_tab(p) for p in self.person_tabs]

    def _refresh_totals(self):
        recs = self._collect_total_records()
        self.tot_tree.delete(*self.tot_tree.get_children())
        tot = dict(visits=0, total_min=0, amt_visit=0, amt_training=0,
                   kotsu=0, total_amount=0)
        for r in recs:
            extra = r["shikaku"] + r["other1"] + r["other2"] + (r["amt_kaizen"] or 0)
            self.tot_tree.insert("", "end", values=(
                r["person"], f"{r['visits']}件", f"{r['total_min']/60:.2f}h",
                f"{r['amt_visit']:,}", f"{r['amt_training']:,}",
                f"{r['kotsu']:,}", f"{extra:,}", f"{r['total_amount']:,}"))
            for k in tot:
                tot[k] += r[k]
        theme.stripe(self.tot_tree)
        self.tot_tree.insert("", "end", tags=("total",), values=(
            "総計", f"{tot['visits']}件", f"{tot['total_min']/60:.2f}h",
            f"{tot['amt_visit']:,}", f"{tot['amt_training']:,}",
            f"{tot['kotsu']:,}", "", f"{tot['total_amount']:,}"))
        self.tot_amount.config(text=f"{tot['total_amount']:,}")
        self.tot_count.config(text=f"{len(recs)}名分の合計")

    def _totals_pdf(self):
        recs = self._collect_total_records()
        if not recs:
            messagebox.showinfo("データなし", "集計対象がありません。")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF", "*.pdf")],
            initialfile=f"全員総計_{prof.JOB}_{self.var_period.get()}.pdf")
        if not path:
            return
        pdf.export_summary_pdf(recs, path)
        if messagebox.askyesno("完了", f"PDFを出力しました。\n{path}\n\n開きますか？"):
            os.startfile(path)

    # ============================================================ 対応表編集
    def _build_mapping(self):
        t = self.tab_mapping
        theme.section(
            t, "対応表",
            "サービス内容ごとに、給与で数える時間（分）を決めます。行をダブルクリックで"
            "編集し、「編集内容を保存」で確定します。").pack(anchor="w", fill="x")

        c_note, note = theme.card(t, padding=12)
        c_note.pack(fill="x", pady=(12, 0))
        theme.field_head(note, "計算のきまり").pack(anchor="w", pady=(0, 6))
        ttk.Label(note, text=f"{prof.P['time_rule']}"
                             "例：訪看Ⅰ２（実績29分）→ 30分、訪看Ⅰ５(×２) → 40分（2単位）。",
                  style="Muted.TLabel").pack(anchor="w")
        ttk.Label(note, text="規定時間を 0 にしたサービス（医療保険の基本療養費・定期巡回など）は"
                             "実時間で数えます。",
                  style="Muted.TLabel").pack(anchor="w", pady=(3, 0))
        ttk.Label(note, text="サービス内容に「★」が付く訪問は同行として、実時間を研修時給で"
                             "計算します（この一覧は使いません）。「★」や空白は照合のときに"
                             "無視します。",
                  style="Muted.TLabel").pack(anchor="w", pady=(3, 0))
        ttk.Label(note, text="ここに無いサービスは計算に含められません。"
                             "計算時にエラーが出たら「対応表外を追加」で取り込んでください。",
                  style="Alert.TLabel", font=theme.FONT_SM).pack(anchor="w", pady=(3, 0))

        cols = ("service", "minutes", "rule")
        heads = {"service": "サービス内容", "minutes": "規定時間(分)", "rule": "数え方"}
        widths = {"service": 300, "minutes": 120, "rule": 220}
        box, tv = theme.scrolled_tree(t, cols, height=14)
        theme.setup_tree(tv, numeric=("minutes",), center=("rule",), widths=widths,
                         heads=heads, stretch="service")
        box.pack(fill="both", expand=True, pady=(12, 0))
        tv.bind("<Double-1>", self._map_edit_cell)
        self.map_tree = tv
        self._sortable(tv, numeric_cols=("minutes",))

        bar = self._toolbar(t, pady=(12, 0))
        ttk.Button(bar, text="編集内容を保存", command=self._map_save,
                   style="Primary.TButton").pack(side="left")
        ttk.Button(bar, text="行を追加", command=self._map_add).pack(
            side="left", padx=8)
        ttk.Button(bar, text="対応表外を追加", command=self._map_add_errors).pack(
            side="left")
        ttk.Button(bar, text="既定に戻す", command=self._map_reset,
                   style="Small.TButton").pack(side="right", padx=8)
        ttk.Button(bar, text="選択行を削除", command=self._map_del,
                   style="Danger.TButton").pack(side="right")
        self._refresh_mapping()

    @staticmethod
    def _rule_text(minutes):
        m = parse_int(minutes)
        if m <= 0:
            return "実時間で数える"
        if prof.UNIT_MIN:
            return f"{m}分で数える（{m / prof.UNIT_MIN:g}単位）"
        return f"{m}分で数える"

    def _refresh_mapping(self):
        self.map_tree.delete(*self.map_tree.get_children())
        for s, m in db.get_mapping_rows():
            self.map_tree.insert("", "end", values=(s, m, self._rule_text(m)))
        theme.stripe(self.map_tree)

    def _map_edit_cell(self, event):
        tv = self.map_tree
        row = tv.identify_row(event.y)
        col = tv.identify_column(event.x)
        if not row:
            return
        col_idx = int(col[1:]) - 1
        if col_idx >= 2:                     # 数え方は自動表示(編集不可)
            return
        x, y, w, hgt = tv.bbox(row, col)
        cur = tv.item(row)["values"][col_idx]
        ent = ttk.Entry(tv)
        ent.place(x=x, y=y, width=w, height=hgt)
        ent.insert(0, str(cur))
        ent.focus_set()
        ent.select_range(0, "end")

        def commit(_e=None):
            vals = list(tv.item(row)["values"])
            v = ent.get().strip()
            if col_idx == 1:
                v = parse_int(v)
                vals[2] = self._rule_text(v)
            vals[col_idx] = v
            tv.item(row, values=vals)
            ent.destroy()
        ent.bind("<Return>", commit)
        ent.bind("<FocusOut>", commit)
        ent.bind("<Escape>", lambda _e: ent.destroy())

    def _map_add(self, service="新規サービス", minutes=0):
        iid = self.map_tree.insert("", "end",
                                   values=(service, minutes, self._rule_text(minutes)))
        theme.stripe(self.map_tree)
        self.map_tree.selection_set(iid)
        self.map_tree.see(iid)
        return iid

    def _map_add_errors(self):
        """計算中の人物で「対応表外」になったサービス内容を、規定0(実時間)で行に足す。"""
        have = {calc.normalize_service(self.map_tree.item(i)["values"][0])
                for i in self.map_tree.get_children()}
        missing = []
        for st in self.person_tabs.values():
            for s in st["res"].get("errors", []):
                if s and s not in have and s not in missing:
                    missing.append(s)
        if not missing:
            self.notify("対応表外のサービスはありません（給与計算タブの計算結果を見ています）。",
                        "info")
            return
        for s in missing:
            self._map_add(s, 0)
        self.notify(f"{len(missing)} 件を「実時間」で追加しました。規定時間を決めたら"
                    "数字を直し、「編集内容を保存」を押してください。", "info")

    def _map_del(self):
        for s in self.map_tree.selection():
            self.map_tree.delete(s)

    def _map_save(self):
        mapping = {}
        for iid in self.map_tree.get_children():
            vals = self.map_tree.item(iid)["values"]
            s = str(vals[0]).strip()
            if s:
                mapping[s] = parse_int(vals[1])
        db.save_mapping(mapping)
        self._refresh_mapping()
        self.notify("対応表を保存しました。次回の計算から反映されます。")
        if self.sources or self.person_tabs:
            self.notify("対応表を保存しました。「計算する」を押すと結果に反映されます。")

    def _map_reset(self):
        if messagebox.askyesno("確認", "対応表を既定値に戻しますか？（現在の編集は失われます）"):
            db.save_mapping(db.DEFAULT_MAPPING)
            self._refresh_mapping()
            self.notify("既定値に戻しました。")

    # ============================================================ その他項目設定
    def _build_settings(self):
        t = self.tab_settings
        theme.section(
            t, "その他項目設定",
            "ここで保存した値が、給与計算画面の初期値になります。").pack(
            anchor="w", fill="x")
        s = db.get_settings()
        self.set_vars = {}

        def money_fields(parent, items, cols=4):
            """単価の入力欄を等間隔に並べる（ラベルは小さく上、数字は等幅）。"""
            grid = ttk.Frame(parent)
            grid.pack(fill="x", pady=(10, 0))
            for i, (key, label, suffix) in enumerate(items):
                box = ttk.Frame(grid)
                box.grid(row=i // cols, column=i % cols, sticky="w",
                         padx=(0, 26), pady=(0, 12))
                theme.field_head(box, label).pack(anchor="w")
                line = ttk.Frame(box)
                line.pack(anchor="w", pady=(3, 0))
                v = tk.StringVar(value=s[key])
                self.set_vars[key] = v
                ttk.Entry(line, textvariable=v, width=8, justify="right",
                          style="Num.TEntry").pack(side="left")
                ttk.Label(line, text=suffix, style="Muted.TLabel").pack(
                    side="left", padx=(5, 0))

        # --- 単価 -------------------------------------------------------
        c1, f1 = theme.card(t)
        c1.pack(fill="x", pady=(12, 0))
        theme.field_head(f1, "時給・交通費").pack(anchor="w")
        money_fields(f1, [("rate_visit", "訪問時給", "円/時"),
                          ("rate_training", "研修時給（★同行）", "円/時"),
                          ("kotsu_unit", "交通費", "円/件"),
                          ("rate_kaizen", "処遇改善手当", "円/時")])
        ttk.Label(f1, text="時給は給与計算画面でも対象月ごとに変更できます"
                           "（そちらの変更は保存されません）。",
                  style="Faint.TLabel").pack(anchor="w")

        # --- 資格手当 ---------------------------------------------------
        c2, f2 = theme.card(t)
        c2.pack(fill="x", pady=(12, 0))
        theme.field_head(f2, "資格手当").pack(anchor="w")
        money_fields(f2, [("default_shikaku", "初期値・未登録者", "円/時")]
                     + [(db.shikaku_key(i), q, "円/時")
                        for i, q in enumerate(db.QUALIFICATIONS)])
        ttk.Label(f2, text="従業員に登録された資格に応じて、計算時に自動で適用されます。"
                           "処遇改善手当・資格手当は訪問時間（★同行を除く）に対してかかります。",
                  style="Faint.TLabel").pack(anchor="w")

        # --- 計上する時間 -------------------------------------------------
        c6, f6 = theme.card(t)
        c6.pack(fill="x", pady=(12, 0))
        theme.field_head(f6, "計上する時間").pack(anchor="w")
        self.var_use_kitei = tk.StringVar(
            value=("0" if s.get("use_kitei", "1") == "0" else "1"))
        urow = ttk.Frame(f6)
        urow.pack(anchor="w", pady=(10, 0))
        ttk.Radiobutton(urow, text="対応表の規定時間（訪看Ⅰ２ の 29分 → 30分）",
                        variable=self.var_use_kitei, value="1").pack(
            side="left", padx=(0, 16))
        ttk.Radiobutton(urow, text="実績の実時間（29分 → 29分）",
                        variable=self.var_use_kitei, value="0").pack(side="left")
        ttk.Label(f6, text=f"{prof.P['time_rule']}"
                           "規定時間が 0 のサービスは、どちらを選んでも実時間です。",
                  style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

        # --- 交通費の条件 ------------------------------------------------
        c3, f3 = theme.card(t)
        c3.pack(fill="x", pady=(12, 0))
        theme.field_head(f3, "交通費の条件").pack(anchor="w")
        self.var_kotsu_first = tk.StringVar(
            value=("0" if s.get("kotsu_first_included", "1") == "0" else "1"))
        krow = ttk.Frame(f3)
        krow.pack(anchor="w", pady=(10, 0))
        ttk.Label(krow, text="1日の最初の1件目を交通費計算に").pack(side="left")
        ttk.Radiobutton(krow, text="含める", variable=self.var_kotsu_first,
                        value="1").pack(side="left", padx=(16, 12))
        ttk.Radiobutton(krow, text="含めない", variable=self.var_kotsu_first,
                        value="0").pack(side="left")
        ttk.Label(f3, text="「含めない」なら、1日にN件あるとき最初の1件を除く N-1 件が対象。"
                          "また、前のサービスの終了時刻と次の開始時刻が同じ行は"
                          "「移動なし」として全利用者共通で対象外です。",
                  style="Muted.TLabel").pack(anchor="w", pady=(8, 0))

        # --- 個人出勤簿 --------------------------------------------------
        c4, f4 = theme.card(t)
        c4.pack(fill="x", pady=(12, 0))
        theme.field_head(f4, "個人出勤簿の既定値").pack(anchor="w")
        self.attend_vars = {}
        agrid = ttk.Frame(f4)
        agrid.pack(fill="x", pady=(10, 0))
        aitems = [("attend_start", "始業時間", "例 8:30"),
                  ("attend_end", "終業時間", "例 17:30"),
                  ("attend_shotei", "所定内勤務時間", "例 8:00"),
                  ("attend_post", "所属", ""),
                  ("attend_employ", "雇用形態", "")]
        for i, (key, label, ph) in enumerate(aitems):
            box = ttk.Frame(agrid)
            box.grid(row=i // 3, column=i % 3, sticky="w", padx=(0, 26),
                     pady=(0, 12))
            theme.field_head(box, label).pack(anchor="w")
            v = tk.StringVar(value=s.get(key, ""))
            self.attend_vars[key] = v
            ttk.Entry(box, textvariable=v, width=16).pack(anchor="w", pady=(3, 0))
            if ph:
                ttk.Label(box, text=ph, style="Faint.TLabel").pack(anchor="w")
        ttk.Label(f4, text="訪問時間は実績から日別に集計し（実時間）、"
                           "訪問以外時間＝所定内−訪問時間で算出します。",
                  style="Faint.TLabel").pack(anchor="w")

        bar = self._toolbar(t, pady=(14, 0))
        ttk.Button(bar, text="設定を保存", command=self._settings_save,
                   style="Primary.TButton").pack(side="left")

        # --- バージョン --------------------------------------------------
        c5, f5 = theme.card(t, padding=12)
        c5.pack(fill="x", pady=(18, 0))
        ttk.Label(f5, text=f"{APP_TITLE}  v{__version__}",
                  style="Muted.TLabel").pack(side="left")
        ttk.Button(f5, text="更新を確認", style="Small.TButton",
                   command=lambda: updater.check_for_updates(
                       self, silent=False)).pack(side="right")

    def _settings_save(self):
        d = {k: str(parse_int(v.get())) for k, v in self.set_vars.items()}
        # 個人出勤簿の既定値は文字列のまま保存
        for k, v in self.attend_vars.items():
            d[k] = v.get().strip()
        # 交通費の条件（含める/含めない）
        d["kotsu_first_included"] = self.var_kotsu_first.get()
        d["use_kitei"] = self.var_use_kitei.get()
        db.save_settings(d)
        # 計算画面へ反映
        self.var_visit.set(d["rate_visit"])
        self.var_train.set(d["rate_training"])
        self.var_kotsu.set(d["kotsu_unit"])
        self.var_kaizen.set(d["rate_kaizen"])
        self.var_shikaku_default.set(d["default_shikaku"])
        if self.sources:
            self.calculate()
            self.notify("設定を保存し、読み込み中の実績で計算し直しました。")
        else:
            self.notify("設定を保存しました。")

    # ============================================================ 共通
    def _on_tab_change(self, _e):
        cur = self.main_nb.tab(self.main_nb.select(), "text")
        if cur == "全員総計":
            self._refresh_totals()
        elif cur == "計算記録":
            self._refresh_records()
        elif cur == "労働記録":
            self._refresh_workdata()


if __name__ == "__main__":
    App().mainloop()
