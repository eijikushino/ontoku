import tkinter as tk
from tkinter import ttk
from tabs.communication_tab import CommunicationTab
from tabs.test_tab import TestTab
from tabs.graph_tab import GraphTab
from tabs.dac_tab import DACTab
from tabs.file_tab import FileTab
from tabs.scanner_tab import ScannerTab
from gpib_controller import GPIBController
from tabs.dmm3458a_tab import DMM3458ATab
from serial_manager import SerialManager

class MainApplication(tk.Tk):
    def __init__(self):
        super().__init__()
        
        self.title("DEF Command Set App")
        self.geometry("900x700")
        
        # ウィンドウを閉じる際の処理
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        # GPIB通信コントローラーのインスタンス化(2台分)
        self.gpib_3458a = GPIBController()  # HP 3458A (DMM) 用
        self.gpib_3499b = GPIBController()  # HP 3499B (Switch) 用
        
        # シリアル通信マネージャー(通信1用)
        self.serial_manager = SerialManager()
        
        # メニューバーの作成
        self.create_menu()
        
        # タブコントロールの作成
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 各タブの作成(2台のコントローラーとシリアルマネージャーを渡す)
        self.comm_tab = CommunicationTab(self.notebook, self.gpib_3458a, self.gpib_3499b, self.serial_manager)
        self.test_tab = TestTab(self.notebook, self.serial_manager)  # serial_managerを渡す
        self.graph_tab = GraphTab(self.notebook, self.gpib_3458a)
        self.dac_tab = DACTab(self.notebook, self.gpib_3499b, self.serial_manager)
        self.file_tab = FileTab(self.notebook)  # 設定管理用タブ（GPIB不要）
        self.dmm3458a_tab = DMM3458ATab(self.notebook, self.gpib_3458a)
        self.scanner_tab = ScannerTab(self.notebook, self.gpib_3499b)
        
        # TestタブにDACタブのDEF選択状態を共有
        self.test_tab.set_def_vars(self.dac_tab.def_vars)
        
        # タブの追加
        self.notebook.add(self.comm_tab, text="通信設定")
        self.notebook.add(self.test_tab, text="Pattern Test")
        self.notebook.add(self.graph_tab, text="グラフ描画")
        self.notebook.add(self.dac_tab, text="DAC操作")
        self.notebook.add(self.file_tab, text="ファイル保存")
        self.notebook.add(self.dmm3458a_tab, text="DMM3458A")
        self.notebook.add(self.scanner_tab, text="スキャナー")
        
        # ステータスバーの作成
        self.create_statusbar()
    
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
            status_text.append("通信1接続中")
        
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
                disconnected.append("通信1")
            except Exception as e:
                self.update_status(f"通信1切断エラー: {e}")
        
        if disconnected:
            self.update_status(f"{', '.join(disconnected)} を切断しました")
        else:
            self.update_status("切断する機器がありません")
        
        self.update_connection_status()
    
    def show_about(self):
        """バージョン情報を表示"""
        about_window = tk.Toplevel(self)
        about_window.title("バージョン情報")
        about_window.geometry("300x150")
        about_window.resizable(False, False)
        
        ttk.Label(about_window, text="DEF Command Set App", 
                  font=("", 12, "bold")).pack(pady=5)
        ttk.Label(about_window, text="Version 1.0").pack(pady=3)
        ttk.Label(about_window, text="・パターン試験").pack(pady=1)
        ttk.Label(about_window, text="・温特グラフ表示").pack(pady=1)
        
        ttk.Button(about_window, text="OK", 
                   command=about_window.destroy).pack(pady=10)
    
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
        self.destroy()

if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()