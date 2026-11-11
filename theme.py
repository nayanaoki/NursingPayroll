# -*- coding: utf-8 -*-
"""給与自動計算システムのデザイントークンと ttk スタイル（標準ttkのみ・依存追加なし）。

考え方（製品タイプ = 帳簿・数字）:
    白いSaaS管理画面ではなく「事務所の帳場」。地はクリーム、面は白、線は茶みの
    あるグレー、差し色はゴールド、主ボタンは墨。金額がこのアプリの主役なので、
    階層は装飾ではなく「数字の大小」で作る。
    色は bar-mellow（同じ帳簿・数字型）と同一トークンを使い、道具として揃える。

使い方:
    theme.apply(root) を App 生成直後に1回だけ呼ぶ。
    画面側では生の色を書かず、ここのトークンと Style 名だけを使う。
"""
import tkinter as tk
from tkinter import ttk
from tkinter import font as tkfont

# --- 地・面・線 -------------------------------------------------------
INK = "#17130F"          # 最も暗い面（ヘッダー・主ボタン）
INK_SOFT = "#2A231B"     # ホバー時の墨
GOLD = "#B8905A"         # 差し色（見出しの罫・フォーカス・選択）
GOLD_SOFT = "#D8BD91"
GOLD_PALE = "#F3E7D4"    # 選択行
GOLD_DEEP = "#8A6B41"
CREAM = "#FAF7F2"        # 地（ページ・ガター）
SURFACE = "#FFFFFF"      # 面（カード・入力・表）
SURFACE_ALT = "#FCFAF6"  # 表のゼブラ
SUNKEN = "#F3EDE4"       # 一段沈んだ面（表の見出し・補助エリア）
LINE = "#E8E0D4"
LINE_STRONG = "#D5C8B6"

# --- 文字 -------------------------------------------------------------
BODY = "#2E2721"         # 本文
MUTED = "#796F64"        # 補助
FAINT = "#A49A8E"        # ごく薄い（プレースホルダ等）
ON_INK = "#F7F3EC"       # 墨の上の文字
ON_INK_SUB = "#B8AC9C"

# --- 意味色（4つだけ。装飾には使わない）-------------------------------
OK = "#1B7A52"
OK_BG = "#E7F4ED"
WARN = "#96660F"
WARN_BG = "#FDF3E2"
ALERT = "#B3352C"
ALERT_BG = "#FCECEB"
INFO = "#2A5F8F"
INFO_BG = "#EBF2F9"

SEMANTIC = {"ok": (OK, OK_BG), "warn": (WARN, WARN_BG),
            "alert": (ALERT, ALERT_BG), "info": (INFO, INFO_BG)}

# --- 旧名（既存コードが参照していたので残す）---------------------------
BG = CREAM
TEXT = BODY
ACCENT = GOLD
BORDER = LINE
HEADER_BG = INK
HEADER_FG = ON_INK
HEADER_SUB = ON_INK_SUB
SUCCESS = OK
DANGER = ALERT

# --- 文字組み ---------------------------------------------------------
FAMILY = "Meiryo UI"
FONT = (FAMILY, 10)
FONT_B = (FAMILY, 10, "bold")
FONT_SM = (FAMILY, 9)
FONT_SM_B = (FAMILY, 9, "bold")
FONT_H = (FAMILY, 12, "bold")       # セクション見出し
FONT_TITLE = (FAMILY, 15, "bold")   # アプリ名

# 金額は桁がそろわないと比較できないので等幅（tabular figures）を使う。
NUM = ("Consolas", 11)
NUM_B = ("Consolas", 11, "bold")
NUM_MID = ("Consolas", 15, "bold")
NUM_HERO = ("Consolas", 30, "bold")

PAD = 14        # カード内の余白
GAP = 10        # 要素どうしの間隔


def apply(root):
    """root(Tk) に配色とスタイルを適用して Style を返す。"""
    style = ttk.Style(root)
    try:
        style.theme_use("clam")   # 色を素直に反映できるテーマ
    except tk.TclError:
        pass
    root.configure(bg=CREAM)

    # 既定フォントは名前付きフォント経由で差し替える。
    # option_add("*Font", ...) は ttk ウィジェットの -font を直接埋めてしまい、
    # スタイル側のフォント指定（見出し・金額）を全部潰すので使わない。
    for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont",
                 "TkHeadingFont", "TkTooltipFont", "TkIconFont"):
        try:
            tkfont.nametofont(name, root=root).configure(family=FAMILY, size=10)
        except tk.TclError:
            pass

    style.configure(".", background=SURFACE, foreground=BODY, font=FONT,
                    focuscolor=GOLD)

    # ---- 面 ----------------------------------------------------------
    # 既定（TFrame / TLabel）は「カードの中」。ページ直置きは Page.* を使う。
    style.configure("TFrame", background=SURFACE)
    style.configure("Page.TFrame", background=CREAM)
    style.configure("Sunken.TFrame", background=SUNKEN)
    style.configure("Ink.TFrame", background=INK)
    style.configure("Rule.TFrame", background=LINE)
    style.configure("GoldRule.TFrame", background=GOLD)

    style.configure("TLabel", background=SURFACE, foreground=BODY)
    style.configure("Muted.TLabel", background=SURFACE, foreground=MUTED,
                    font=FONT_SM)
    style.configure("Faint.TLabel", background=SURFACE, foreground=FAINT,
                    font=FONT_SM)
    style.configure("Strong.TLabel", background=SURFACE, foreground=BODY,
                    font=FONT_B)
    style.configure("Head.TLabel", background=SURFACE, foreground=INK,
                    font=FONT_H)
    style.configure("FieldHead.TLabel", background=SURFACE, foreground=MUTED,
                    font=FONT_SM_B)
    # ページ（クリーム地）の上に直接置く文字
    style.configure("PageTitle.TLabel", background=CREAM, foreground=INK,
                    font=FONT_H)
    style.configure("PageDesc.TLabel", background=CREAM, foreground=MUTED,
                    font=FONT_SM)
    style.configure("Page.TLabel", background=CREAM, foreground=BODY)
    for kind, (fg, _bg) in SEMANTIC.items():
        style.configure(f"{kind.title()}.TLabel", background=SURFACE, foreground=fg)
        style.configure(f"Page{kind.title()}.TLabel", background=CREAM,
                        foreground=fg, font=FONT_SM)

    # 金額
    style.configure("Num.TLabel", background=SURFACE, foreground=BODY, font=NUM)
    style.configure("NumB.TLabel", background=SURFACE, foreground=BODY, font=NUM_B)
    style.configure("NumMid.TLabel", background=SURFACE, foreground=INK, font=NUM_MID)
    style.configure("NumHero.TLabel", background=SURFACE, foreground=INK,
                    font=NUM_HERO)

    # ---- ボタン（1画面に Primary は1つ）--------------------------------
    def _btn(name, bg, fg, border, hover, active, pad=(14, 7), font=FONT):
        # clam は relief="solid" のときだけ bordercolor で1pxの枠を描く。
        # 枠が無いと、クリーム地の上で白いボタンが「ただの文字」に見える。
        style.configure(name, background=bg, foreground=fg, bordercolor=border,
                        lightcolor=bg, darkcolor=bg, focuscolor=border,
                        relief="solid", borderwidth=1, padding=pad, font=font,
                        anchor="center")
        style.map(name,
                  background=[("pressed", active), ("active", hover),
                              ("disabled", SUNKEN)],
                  foreground=[("disabled", FAINT)],
                  bordercolor=[("active", border)])

    _btn("TButton", SURFACE, BODY, LINE_STRONG, SUNKEN, LINE)
    _btn("Primary.TButton", INK, ON_INK, INK, INK_SOFT, INK_SOFT, font=FONT_B)
    _btn("Danger.TButton", SURFACE, ALERT, LINE_STRONG, ALERT_BG, ALERT_BG)
    _btn("Ghost.TButton", CREAM, MUTED, CREAM, SUNKEN, LINE, pad=(10, 6))
    _btn("Small.TButton", SURFACE, BODY, LINE_STRONG, SUNKEN, LINE,
         pad=(9, 4), font=FONT_SM)
    # 旧名（Accent = 主ボタン）
    _btn("Accent.TButton", INK, ON_INK, INK, INK_SOFT, INK_SOFT, font=FONT_B)

    # ---- タブ ---------------------------------------------------------
    style.configure("TNotebook", background=CREAM, borderwidth=0,
                    tabmargins=(10, 8, 10, 0))
    style.configure("TNotebook.Tab", background=CREAM, foreground=MUTED,
                    padding=(16, 8), borderwidth=0, font=FONT)
    style.map("TNotebook.Tab",
              background=[("selected", SURFACE), ("active", SUNKEN)],
              foreground=[("selected", INK), ("active", BODY)],
              font=[("selected", FONT_B)])
    # 人物タブ（給与計算の中の入れ子ノートブック）は一段小さく
    style.configure("Person.TNotebook", background=SURFACE, borderwidth=0,
                    tabmargins=(0, 6, 0, 0))
    style.configure("Person.TNotebook.Tab", background=SUNKEN, foreground=MUTED,
                    padding=(14, 6), borderwidth=0, font=FONT_SM)
    style.map("Person.TNotebook.Tab",
              background=[("selected", INK), ("active", LINE)],
              foreground=[("selected", ON_INK), ("active", BODY)],
              font=[("selected", FONT_SM_B)])

    # ---- 枠つきグループ ------------------------------------------------
    style.configure("TLabelframe", background=SURFACE, bordercolor=LINE,
                    lightcolor=SURFACE, darkcolor=SURFACE,
                    relief="solid", borderwidth=1)
    style.configure("TLabelframe.Label", background=SURFACE, foreground=MUTED,
                    font=FONT_SM_B)

    # ---- 入力 ---------------------------------------------------------
    style.configure("TEntry", fieldbackground=SURFACE, foreground=BODY,
                    bordercolor=LINE_STRONG, lightcolor=LINE, darkcolor=LINE,
                    insertcolor=INK, padding=5, relief="flat")
    style.map("TEntry",
              bordercolor=[("focus", GOLD)], lightcolor=[("focus", GOLD)],
              darkcolor=[("focus", GOLD)],
              fieldbackground=[("disabled", SUNKEN)])
    style.configure("Num.TEntry", fieldbackground=SURFACE, foreground=BODY,
                    bordercolor=LINE_STRONG, lightcolor=LINE, darkcolor=LINE,
                    insertcolor=INK, padding=5, relief="flat", font=NUM)
    style.map("Num.TEntry",
              bordercolor=[("focus", GOLD)], lightcolor=[("focus", GOLD)],
              darkcolor=[("focus", GOLD)])

    style.configure("TCombobox", fieldbackground=SURFACE, background=SURFACE,
                    foreground=BODY, bordercolor=LINE_STRONG,
                    lightcolor=LINE, darkcolor=LINE,
                    arrowcolor=MUTED, arrowsize=14, padding=4, relief="flat")
    style.map("TCombobox",
              bordercolor=[("focus", GOLD)], lightcolor=[("focus", GOLD)],
              darkcolor=[("focus", GOLD)],
              arrowcolor=[("active", INK)],
              fieldbackground=[("readonly", SURFACE), ("disabled", SUNKEN)])
    # ドロップダウンのリスト側（ttk では option 経由でしか色を変えられない）
    root.option_add("*TCombobox*Listbox.background", SURFACE)
    root.option_add("*TCombobox*Listbox.foreground", BODY)
    root.option_add("*TCombobox*Listbox.selectBackground", GOLD_PALE)
    root.option_add("*TCombobox*Listbox.selectForeground", INK)
    root.option_add("*TCombobox*Listbox.font", FONT)

    # ---- 一覧表（このアプリで最も長く見る面）-----------------------------
    style.configure("Treeview", background=SURFACE, fieldbackground=SURFACE,
                    foreground=BODY, rowheight=29, borderwidth=0, relief="flat")
    style.configure("Treeview.Heading", background=SUNKEN, foreground=MUTED,
                    relief="flat", borderwidth=0, font=FONT_SM_B,
                    padding=(6, 8))
    style.map("Treeview.Heading",
              background=[("active", LINE)], foreground=[("active", INK)])
    style.map("Treeview",
              background=[("selected", GOLD_PALE)],
              foreground=[("selected", INK)])
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])

    # ---- その他 --------------------------------------------------------
    style.configure("TRadiobutton", background=SURFACE, foreground=BODY,
                    indicatorcolor=SURFACE, bordercolor=LINE_STRONG)
    style.map("TRadiobutton",
              foreground=[("active", INK)],
              indicatorcolor=[("selected", GOLD)],
              background=[("active", SURFACE)])
    style.configure("Page.TRadiobutton", background=CREAM, foreground=BODY,
                    indicatorcolor=SURFACE, bordercolor=LINE_STRONG)
    style.map("Page.TRadiobutton",
              foreground=[("active", INK)],
              indicatorcolor=[("selected", GOLD)],
              background=[("active", CREAM)])

    style.configure("TSeparator", background=LINE)
    for orient in ("Vertical", "Horizontal"):
        style.configure(f"{orient}.TScrollbar", background=LINE_STRONG,
                        troughcolor=CREAM, bordercolor=CREAM, arrowcolor=MUTED,
                        lightcolor=LINE_STRONG, darkcolor=LINE_STRONG,
                        relief="flat", borderwidth=0, arrowsize=13)
        style.map(f"{orient}.TScrollbar",
                  background=[("active", GOLD_SOFT), ("pressed", GOLD)])
    return style


# ======================================================== 共通パーツ
def card(parent, padding=PAD):
    """白い面（カード）を作り、中身を置くフレームを返す。

    枠線つきの器は tk.Frame の highlight で描く（clam の TFrame より確実）。
    """
    outer = tk.Frame(parent, bg=SURFACE, highlightbackground=LINE,
                     highlightcolor=LINE, highlightthickness=1, bd=0)
    inner = ttk.Frame(outer, padding=padding)
    inner.pack(fill="both", expand=True)
    inner.card = outer          # pack/grid は呼び出し側が outer に対して行う
    return outer, inner


def section(parent, title, desc=None, note=None, note_kind="info"):
    """ページ見出し（ゴールドの短い罫 + 見出し + 補足）を作る。"""
    box = ttk.Frame(parent, style="Page.TFrame")
    head = ttk.Frame(box, style="Page.TFrame")
    head.pack(fill="x")
    tk.Frame(head, bg=GOLD, width=3, height=17).pack(side="left", padx=(0, 8))
    ttk.Label(head, text=title, style="PageTitle.TLabel").pack(side="left")
    if desc:
        ttk.Label(box, text=desc, style="PageDesc.TLabel",
                  justify="left").pack(anchor="w", pady=(5, 0), padx=(11, 0))
    if note:
        ttk.Label(box, text=note, style=f"Page{note_kind.title()}.TLabel",
                  justify="left").pack(anchor="w", pady=(3, 0), padx=(11, 0))
    return box


def field_head(parent, text):
    """カード内の小見出し（表の見出しと同じ質感の小文字ラベル）。"""
    return ttk.Label(parent, text=text, style="FieldHead.TLabel")


def setup_tree(tree, numeric=(), center=(), widths=None, heads=None,
               stretch=None):
    """Treeview の見出し・揃え・ゼブラ用タグをまとめて設定する。

    余った横幅は stretch で指定した1列だけが吸収する。全列を伸ばすと
    値と値の間が間延びして、行として読めなくなるため。
    """
    for col in tree["columns"]:
        anchor = "e" if col in numeric else ("center" if col in center else "w")
        if heads:
            tree.heading(col, text=heads.get(col, col), anchor=anchor)
        tree.column(col, anchor=anchor,
                    width=(widths or {}).get(col, 100),
                    minwidth=40, stretch=(col == stretch))
    tree.tag_configure("odd", background=SURFACE_ALT)
    tree.tag_configure("even", background=SURFACE)
    tree.tag_configure("total", background=SUNKEN, font=FONT_B, foreground=INK)
    tree.tag_configure("ok", foreground=OK)
    tree.tag_configure("alert", foreground=ALERT)
    tree.tag_configure("muted", foreground=MUTED)
    return tree


def stripe(tree):
    """一覧の再描画後に呼ぶ。既存タグを保ったままゼブラを付け直す。"""
    for i, iid in enumerate(tree.get_children()):
        tags = [t for t in tree.item(iid, "tags") if t not in ("odd", "even")]
        tags.append("odd" if i % 2 else "even")
        tree.item(iid, tags=tuple(tags))


def scrolled_tree(parent, columns, height=12):
    """カードに収めた Treeview（縦スクロールバー付き）を作る。

    戻り値は (カード外枠, Treeview)。外枠を pack/grid する。
    """
    outer = tk.Frame(parent, bg=SURFACE, highlightbackground=LINE,
                     highlightcolor=LINE, highlightthickness=1, bd=0)
    wrap = ttk.Frame(outer, padding=1)
    wrap.pack(fill="both", expand=True)
    tree = ttk.Treeview(wrap, columns=columns, show="headings", height=height)
    vsb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)

    def _auto(lo, hi):
        """行が収まっているときはスクロールバーを出さない。"""
        if float(lo) <= 0.0 and float(hi) >= 1.0:
            vsb.pack_forget()
        elif not vsb.winfo_ismapped():
            vsb.pack(side="right", fill="y", before=tree)
        vsb.set(lo, hi)

    tree.configure(yscrollcommand=_auto)
    tree.pack(side="left", fill="both", expand=True)
    return outer, tree


class Banner(ttk.Frame):
    """ヘッダー直下に固定される通知帯。

    画面の一番下の薄い文字ではなく、常に見える位置で操作結果を返す。
    ok / info は数秒で自動的に消え、warn / alert は消さない（対処が要るため）。
    """

    def __init__(self, master):
        super().__init__(master, style="Page.TFrame")
        self._after = None
        self._box = tk.Frame(self, bg=OK_BG, bd=0)
        self._rule = tk.Frame(self._box, bg=OK, width=3)
        self._rule.pack(side="left", fill="y")
        self._icon = tk.Label(self._box, bg=OK_BG, fg=OK, font=FONT_B,
                              text="", padx=10)
        self._icon.pack(side="left")
        self._text = tk.Label(self._box, bg=OK_BG, fg=OK, font=FONT,
                              text="", justify="left", anchor="w", wraplength=1100)
        self._text.pack(side="left", fill="x", expand=True, pady=8, padx=(0, 8))
        self._close = tk.Label(self._box, bg=OK_BG, fg=OK, font=FONT_B,
                               text="×", padx=12, cursor="hand2")
        self._close.pack(side="right")
        self._close.bind("<Button-1>", lambda _e: self.hide())

    MARKS = {"ok": "○", "info": "i", "warn": "!", "alert": "×"}

    def show(self, text, kind="ok", timeout=6000):
        fg, bg = SEMANTIC.get(kind, SEMANTIC["info"])
        self._box.configure(bg=bg)
        self._rule.configure(bg=fg)
        for w in (self._icon, self._text, self._close):
            w.configure(bg=bg, fg=fg)
        self._icon.configure(text=self.MARKS.get(kind, "i"))
        self._text.configure(text=text)
        self._box.pack(fill="x")
        self.pack_configure(pady=(0, 0))
        if self._after:
            self.after_cancel(self._after)
            self._after = None
        if kind in ("ok", "info") and timeout:
            self._after = self.after(timeout, self.hide)

    def hide(self):
        if self._after:
            self.after_cancel(self._after)
            self._after = None
        self._box.pack_forget()
