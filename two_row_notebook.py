import tkinter as tk
from tkinter import ttk


class TwoRowNotebook(ttk.Frame):
    """2段タブヘッダーを持つカスタムNotebookウィジェット"""

    # ボタン共通設定
    FONT_NORMAL = ('', 10)
    FONT_BOLD = ('', 10, 'bold')
    BTN_PADX = 14
    BTN_PADY = 5

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)

        # タブヘッダー領域
        self._header = ttk.Frame(self)
        self._header.pack(fill=tk.X, padx=4)

        # 1段目タブバー
        self._row1 = ttk.Frame(self._header)
        self._row1.pack(fill=tk.X, pady=(4, 0))

        # 2段目タブバー（機器操作ラベル付き）
        self._row2_frame = ttk.Frame(self._header)
        self._row2_frame.pack(fill=tk.X, pady=(2, 0))
        ttk.Label(self._row2_frame, text="機器操作 :", font=('', 9)).pack(
            side=tk.LEFT, padx=(2, 4))
        self._row2 = ttk.Frame(self._row2_frame)
        self._row2.pack(side=tk.LEFT, fill=tk.X)

        # セパレータ
        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(4, 0))

        # コンテンツ領域 — 各タブフレームの親として使用
        self.content = ttk.Frame(self)
        self.content.pack(fill=tk.BOTH, expand=True)

        # 内部管理
        self._tabs = []      # [{frame, text, row, button}, ...]
        self._current = None  # 選択中インデックス

        # タブ色定義
        self._tab_colors = {}
        self._selected_colors = {}

    def create_group(self, row=1, label=""):
        """行内にグループ枠（LabelFrame）を作成して返す"""
        target = self._row1 if row == 1 else self._row2
        group = tk.LabelFrame(target, text=label, font=('', 8),
                              padx=2, pady=2, bd=1, relief=tk.GROOVE)
        group.pack(side=tk.LEFT, padx=(6, 2), pady=0)
        return group

    def add(self, frame, text="", row=1, group=None):
        """タブを追加。row=1: 1段目、row=2: 2段目。group: グループ枠"""
        idx = len(self._tabs)
        if group is not None:
            target = group
        else:
            target = self._row1 if row == 1 else self._row2

        btn = tk.Button(target, text=text, relief=tk.FLAT, bd=1,
                        padx=self.BTN_PADX, pady=self.BTN_PADY,
                        cursor='hand2', font=self.FONT_NORMAL,
                        command=lambda i=idx: self._select_tab(i))
        btn.pack(side=tk.LEFT, padx=1)

        self._tabs.append({
            'frame': frame,
            'text': text,
            'row': row,
            'button': btn,
        })

        # 最初のタブを自動選択
        if self._current is None:
            self._select_tab(0)

    def _select_tab(self, index):
        """タブ選択の内部処理"""
        if index < 0 or index >= len(self._tabs):
            return

        # 現在のタブを非表示
        if self._current is not None:
            self._tabs[self._current]['frame'].pack_forget()

        # 新しいタブを表示
        self._tabs[index]['frame'].pack(fill=tk.BOTH, expand=True)
        self._current = index
        self._update_all_buttons()
        self.event_generate('<<NotebookTabChanged>>')

    def select(self, tab_or_index=None):
        """タブ選択（引数あり）/ 現在のタブインデックス取得（引数なし）"""
        if tab_or_index is None:
            return self._current if self._current is not None else 0

        if isinstance(tab_or_index, int):
            self._select_tab(tab_or_index)
        else:
            for i, tab in enumerate(self._tabs):
                if tab['frame'] == tab_or_index:
                    self._select_tab(i)
                    return

    def index(self, what):
        """ttk.Notebook互換: タブインデックスを返す"""
        if isinstance(what, int):
            return what
        return self._current if self._current is not None else 0

    def set_tab_colors(self, colors, selected_colors):
        """タブごとの色を設定: {index: '#色コード'}"""
        self._tab_colors = colors
        self._selected_colors = selected_colors
        self._update_all_buttons()

    def _update_all_buttons(self):
        """全タブボタンの見た目を更新"""
        for i, tab in enumerate(self._tabs):
            btn = tab['button']
            if i == self._current:
                bg = self._selected_colors.get(i, '#c8d8e8')
                btn.config(relief=tk.GROOVE, bg=bg, fg='#000000',
                           font=self.FONT_BOLD)
            else:
                bg = self._tab_colors.get(i, '#f0f0f0')
                btn.config(relief=tk.FLAT, bg=bg, fg='#444444',
                           font=self.FONT_NORMAL)
