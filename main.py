import matplotlib
matplotlib.use('TkAgg')

import tkinter as tk
from tkinter import ttk
from tabs.communication_tab import CommunicationTab
from tabs.test_tab import TestTab
from tabs.graph_tab import GraphTab
from tabs.dac_tab import DACTab
from tabs.datagen_tab import DataGenTab
from tabs.file_tab import FileTab
from tabs.scanner_tab import ScannerTab
from tabs.linearity_tab import LinearityTab
from tabs.dc_char_tab import DCCharTab
from gpib_controller import GPIBController
from tabs.dmm3458a_tab import DMM3458ATab
from serial_manager import SerialManager
from about_dialog import show_about_dialog
from two_row_notebook import TwoRowNotebook
from version import __version__

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title(f"DEF Command Set App v{__version__}")
        self.geometry("900x780")
        
        # ウィンドウを閉じる際の処理
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # GPIB通信コントローラーのインスタンス化(2台分)
        self.gpib_3458a = GPIBController()  # HP 3458A (DMM) 用
        self.gpib_3499b = GPIBController()  # HP 3499B (Switch) 用
        
        # シリアル通信マネージャー(DEFシリアル通信用: 38400bps)
        self.serial_manager = SerialManager()

        # DataGen専用SerialManager(115200bps)
        self.datagen_manager = SerialManager(baudrate=115200)
        self.datagen_manager2 = SerialManager(baudrate=115200)
        
        # メニューバーの作成
        self.create_menu()
        
        # タブコントロールの作成（2段構成）
        self.notebook = TwoRowNotebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # コンテンツ領域を各タブの親にする
        content = self.notebook.content

        # 各タブの作成(2台のコントローラーとシリアルマネージャーを渡す)
        self.comm_tab = CommunicationTab(content, self.gpib_3458a, self.gpib_3499b, self.serial_manager,
                                          datagen_manager=self.datagen_manager, datagen_manager2=self.datagen_manager2)
        self.test_tab = TestTab(content, self.serial_manager)
        self.graph_tab = GraphTab(content, self.gpib_3458a)
        self.dac_tab = DACTab(content, self.gpib_3499b, self.serial_manager)
        self.datagen_tab = DataGenTab(content, self.datagen_manager, self.datagen_manager2)
        self.file_tab = FileTab(content)
        self.dmm3458a_tab = DMM3458ATab(content, self.gpib_3458a)
        self.scanner_tab = ScannerTab(content, self.gpib_3499b)
        self.linearity_tab = LinearityTab(content, self.gpib_3458a, self.gpib_3499b,
                                           self.datagen_manager, self.test_tab)
        self.dc_char_tab = DCCharTab(content, self.gpib_3458a, self.gpib_3499b,
                                      self.datagen_manager, self.serial_manager, self.test_tab)

        # TestタブにDACタブのDEF選択状態を共有
        self.test_tab.set_def_vars(self.dac_tab.def_vars)

        # タブのスタイル設定（色分け・パディング）
        self._setup_tab_styles()

        # 1段目: メイン機能タブ
        self.notebook.add(self.comm_tab, text="  通信設定  ", row=1)
        self.notebook.add(self.file_tab, text="  ファイル保存  ", row=1)
        # Pattern Test + グラフ描画をグループ枠で囲む
        test_group = self.notebook.create_group(row=1, label="温特&パターン")
        self.notebook.add(self.test_tab, text="  Pattern Test  ", row=1, group=test_group)
        self.notebook.add(self.graph_tab, text="  グラフ描画  ", row=1, group=test_group)
        self.notebook.add(self.linearity_tab, text="  Linearity  ", row=1)
        self.notebook.add(self.dc_char_tab, text="  DC特性  ", row=1)

        # 2段目: 機器操作タブ
        self.notebook.add(self.dac_tab, text="  DEF操作  ", row=2)
        self.notebook.add(self.datagen_tab, text="  DataGen  ", row=2)
        self.notebook.add(self.dmm3458a_tab, text="  DMM3458A  ", row=2)
        self.notebook.add(self.scanner_tab, text="  スキャナー  ", row=2)

        # ステータスバーの作成
        self.create_statusbar()
    
    def _setup_tab_styles(self):
        """タブのスタイルを設定（色分け・パディング）"""
        style = ttk.Style()

        # カスタマイズ可能なテーマに変更（Windowsデフォルトは色変更不可）
        style.theme_use('clam')

        # チェックボックスのレ点画像を作成
        self._create_checkbox_images(style)

        # 背景色を薄いグレーに、ボタンを白系に設定
        bg_color = '#f0f0f0'
        style.configure('TFrame', background=bg_color, borderwidth=0, relief='flat')
        style.configure('TLabel', background=bg_color)
        style.configure('TLabelframe', background=bg_color)
        style.configure('TLabelframe.Label', background=bg_color)
        style.configure('TCheckbutton', background=bg_color)
        style.configure('TRadiobutton', background=bg_color)

        # ボタンを濃い灰色で目立たせる（浮いた立体感）
        style.configure('TButton',
                        background='#c0c0c0',
                        padding=[4, 2],
                        relief='raised',
                        borderwidth=2,
                        lightcolor='#e8e8e8',
                        darkcolor='#808080')
        style.map('TButton',
                  background=[('active', '#b0b0b0'), ('pressed', '#a0a0a0')],
                  relief=[('pressed', 'sunken')],
                  lightcolor=[('pressed', '#808080')],
                  darkcolor=[('pressed', '#e8e8e8')])

        # タブ変更時に色を適用
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)

        # 初期状態でも色を適用するため、少し遅延して実行
        self.after(100, self._apply_tab_colors)

    def _create_checkbox_images(self, style):
        """チェックボックス用のレ点画像を作成"""
        # チェックなし（空の四角）
        self._checkbox_unchecked = tk.PhotoImage(width=16, height=16)
        self._checkbox_unchecked.put(('#999999',), to=(0, 0, 16, 1))    # 上辺
        self._checkbox_unchecked.put(('#999999',), to=(0, 15, 16, 16))  # 下辺
        self._checkbox_unchecked.put(('#999999',), to=(0, 0, 1, 16))    # 左辺
        self._checkbox_unchecked.put(('#999999',), to=(15, 0, 16, 16))  # 右辺
        # 内部を白で塗りつぶし
        for y in range(1, 15):
            self._checkbox_unchecked.put(('#ffffff',), to=(1, y, 15, y+1))

        # チェックあり（レ点）
        self._checkbox_checked = tk.PhotoImage(width=16, height=16)
        self._checkbox_checked.put(('#999999',), to=(0, 0, 16, 1))
        self._checkbox_checked.put(('#999999',), to=(0, 15, 16, 16))
        self._checkbox_checked.put(('#999999',), to=(0, 0, 1, 16))
        self._checkbox_checked.put(('#999999',), to=(15, 0, 16, 16))
        for y in range(1, 15):
            self._checkbox_checked.put(('#ffffff',), to=(1, y, 15, y+1))
        # レ点を描画（緑色のチェックマーク）
        checkmark = [
            (3, 8), (4, 9), (5, 10), (6, 11),  # 左下から
            (7, 10), (8, 9), (9, 8), (10, 7), (11, 6), (12, 5), (13, 4)  # 右上へ
        ]
        for x, y in checkmark:
            self._checkbox_checked.put(('#22aa22',), to=(x, y, x+2, y+2))

        # スタイルに適用
        style.element_create('custom.checkbox.indicator', 'image', self._checkbox_unchecked,
                             ('selected', self._checkbox_checked))
        style.layout('TCheckbutton', [
            ('Checkbutton.padding', {'children': [
                ('custom.checkbox.indicator', {'side': 'left', 'sticky': ''}),
                ('Checkbutton.label', {'side': 'left', 'sticky': 'nswe'})
            ], 'sticky': 'nswe'})
        ])

    def _on_tab_changed(self, event=None):
        """タブ変更時の処理"""
        self._apply_tab_colors()

    def _apply_tab_colors(self):
        """各タブに色を適用（2段タブ対応）"""
        # 新しいタブ順序に合わせた色定義
        # 1段目: 通信設定(0), ファイル保存(1), Pattern Test(2), グラフ描画(3), Linearity(4), DC特性(5)
        # 2段目: DEF操作(6), DataGen(7), DMM3458A(8), スキャナー(9)
        tab_colors = {
            0: '#d4e6f1',  # 通信設定 - 青系
            1: '#d4e6f1',  # ファイル保存 - 青系
            2: '#d5f5e3',  # Pattern Test - 緑系
            3: '#d5f5e3',  # グラフ描画 - 緑系
            4: '#d5f5e3',  # Linearity - 緑系
            5: '#d5f5e3',  # DC特性 - 緑系
            6: '#e8daef',  # DEF操作 - 紫系
            7: '#e8daef',  # DataGen - 紫系
            8: '#e8daef',  # DMM3458A - 紫系
            9: '#e8daef',  # スキャナー - 紫系
        }

        selected_colors = {
            0: '#a9cce3',  # 通信設定
            1: '#a9cce3',  # ファイル保存
            2: '#abebc6',  # Pattern Test
            3: '#abebc6',  # グラフ描画
            4: '#abebc6',  # Linearity
            5: '#abebc6',  # DC特性
            6: '#d2b4de',  # DEF操作
            7: '#d2b4de',  # DataGen
            8: '#d2b4de',  # DMM3458A
            9: '#d2b4de',  # スキャナー
        }

        self.notebook.set_tab_colors(tab_colors, selected_colors)

    def create_menu(self):
        """メニューバーを作成"""
        menubar = tk.Menu(self)
        self.config(menu=menubar)
        
        # ファイルメニュー
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        file_menu.add_command(label="終了", command=self.on_closing)
        
        # 接続メニュー
        connect_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="接続", menu=connect_menu)
        connect_menu.add_command(label="機器検索", command=self.search_devices)
        connect_menu.add_separator()
        connect_menu.add_command(label="全機器接続解除", command=self.disconnect_device)
        
        # ヘルプメニュー
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="ヘルプ", menu=help_menu)
        help_menu.add_command(label="バージョン情報", command=self.show_about)
    
    def create_statusbar(self):
        """ステータスバーを作成"""
        self.statusbar = ttk.Frame(self, relief=tk.SUNKEN)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.status_label = ttk.Label(self.statusbar, text="準備完了", anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, padx=10, pady=2)
        
        self.connection_label = ttk.Label(self.statusbar, text="未接続", anchor=tk.E)
        self.connection_label.pack(side=tk.RIGHT, padx=10, pady=2)
    
    def update_status(self, message):
        """ステータスバーのメッセージを更新"""
        self.status_label.config(text=message)
    
    def update_connection_status(self):
        """接続状態を更新"""
        status_text = []
        if self.gpib_3458a.connected:
            status_text.append("3458A接続中")
        if self.gpib_3499b.connected:
            status_text.append("3499B接続中")
        if self.serial_manager.is_connected():
            status_text.append("DEF接続中")
        
        if status_text:
            self.connection_label.config(text=" / ".join(status_text), foreground="green")
        else:
            self.connection_label.config(text="未接続", foreground="red")
    
    def search_devices(self):
        """機器検索(メニューから)"""
        self.notebook.select(0)  # 通信設定タブに切り替え
        self.comm_tab.search_resources()
    
    def disconnect_device(self):
        """接続解除(メニューから) - 全機器を切断"""
        disconnected = []
        
        if self.gpib_3458a.connected:
            success, message = self.gpib_3458a.disconnect(go_to_local=True)
            if success:
                disconnected.append("3458A")
            self.update_status(f"3458A: {message}")
        
        if self.gpib_3499b.connected:
            success, message = self.gpib_3499b.disconnect(go_to_local=True)
            if success:
                disconnected.append("3499B")
            self.update_status(f"3499B: {message}")
        
        if self.serial_manager.is_connected():
            try:
                self.serial_manager.disconnect()
                disconnected.append("DEF")
            except Exception as e:
                self.update_status(f"DEFシリアル切断エラー: {e}")

        if self.datagen_manager.is_connected():
            try:
                self.datagen_manager.disconnect()
                disconnected.append("DG1")
            except Exception as e:
                self.update_status(f"DG1切断エラー: {e}")

        if self.datagen_manager2.is_connected():
            try:
                self.datagen_manager2.disconnect()
                disconnected.append("DG2")
            except Exception as e:
                self.update_status(f"DG2切断エラー: {e}")

        if disconnected:
            self.update_status(f"{', '.join(disconnected)} を切断しました")
        else:
            self.update_status("切断する機器がありません")
        
        self.update_connection_status()
    
    def show_about(self):
        """バージョン情報を表示"""
        show_about_dialog(self)
    
    def on_closing(self):
        """アプリケーション終了時の処理"""
        # 両機器をローカルモードに復帰させて切断
        if self.gpib_3458a.connected:
            self.gpib_3458a.disconnect(go_to_local=True)
        if self.gpib_3499b.connected:
            self.gpib_3499b.disconnect(go_to_local=True)
        # シリアル通信を切断
        if self.serial_manager.is_connected():
            self.serial_manager.disconnect()
        # DataGen通信を切断
        if self.datagen_manager.is_connected():
            self.datagen_manager.disconnect()
        if self.datagen_manager2.is_connected():
            self.datagen_manager2.disconnect()
        # Matplotlibのグラフウィンドウを全て閉じる
        import matplotlib.pyplot as plt
        plt.close('all')
        self.destroy()

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()