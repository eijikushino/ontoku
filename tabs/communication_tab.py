import tkinter as tk
from tkinter import ttk
import serial.tools.list_ports
from utils import LoggerWidget, validate_integer
import json
import os

class CommunicationTab(ttk.Frame):
    def __init__(self, parent, gpib_3458a, gpib_3499b, serial_manager):
        super().__init__(parent)
        self.gpib_3458a = gpib_3458a
        self.gpib_3499b = gpib_3499b
        self.serial_mgr = serial_manager  # 通信1用SerialManager
        
        # 汎用設定ファイルのパス
        self.config_file = "app_settings.json"
        
        self.create_widgets()
        
        # 起動時に設定を読み込み
        self.load_config()
    
    def create_widgets(self):
        # リソース検索フレーム
        resource_frame = ttk.LabelFrame(self, text="通信機器検索", padding=10)
        resource_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(resource_frame, text="機器を検索", 
                   command=self.search_resources).pack(side=tk.LEFT, padx=5)
        
        self.resource_list = tk.Listbox(resource_frame, height=4, width=25)
        self.resource_list.pack(side=tk.LEFT, padx=5)
        
        # スクロールバー
        scrollbar = ttk.Scrollbar(resource_frame, orient=tk.VERTICAL, 
                                  command=self.resource_list.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.resource_list.config(yscrollcommand=scrollbar.set)
        
        # === 3458A 接続フレーム ===
        frame_3458a = ttk.LabelFrame(self, text="HP 3458A (DMM) 接続制御", padding=10)
        frame_3458a.pack(fill=tk.X, padx=10, pady=5)
        
        # リソース名入力
        ttk.Label(frame_3458a, text="リソース名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.resource_3458a_entry = ttk.Entry(frame_3458a, width=30)
        self.resource_3458a_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        self.resource_3458a_entry.insert(0, "GPIB0::22::INSTR")  # デフォルト値
        
        # リストから設定ボタン
        ttk.Button(frame_3458a, text="リストから設定", 
                   command=lambda: self.set_resource_from_list(self.resource_3458a_entry)).grid(
                   row=0, column=2, padx=5, pady=2)
        
        # 接続/切断ボタン + ステータス
        btn_frame_3458a = ttk.Frame(frame_3458a)
        btn_frame_3458a.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.W)
        
        self.btn_connect_3458a = ttk.Button(btn_frame_3458a, text="接続", 
                                             command=self.connect_3458a, width=12)
        self.btn_connect_3458a.pack(side=tk.LEFT, padx=5)
        
        self.btn_disconnect_3458a = ttk.Button(btn_frame_3458a, text="切断", 
                                                command=self.disconnect_3458a, 
                                                width=12, state=tk.DISABLED)
        self.btn_disconnect_3458a.pack(side=tk.LEFT, padx=5)
        
        # 接続状態表示（切断ボタンの横）
        self.status_3458a = ttk.Label(btn_frame_3458a, text="未接続", foreground="red")
        self.status_3458a.pack(side=tk.LEFT, padx=10)
        
        # === 3499B 接続フレーム ===
        frame_3499b = ttk.LabelFrame(self, text="HP 3499B (Switch) 接続制御", padding=10)
        frame_3499b.pack(fill=tk.X, padx=10, pady=5)
        
        # リソース名入力
        ttk.Label(frame_3499b, text="リソース名:").grid(row=0, column=0, sticky=tk.W, padx=5, pady=2)
        self.resource_3499b_entry = ttk.Entry(frame_3499b, width=30)
        self.resource_3499b_entry.grid(row=0, column=1, sticky=tk.W, padx=5, pady=2)
        self.resource_3499b_entry.insert(0, "GPIB0::09::INSTR")  # デフォルト値
        
        # リストから設定ボタン
        ttk.Button(frame_3499b, text="リストから設定", 
                   command=lambda: self.set_resource_from_list(self.resource_3499b_entry)).grid(
                   row=0, column=2, padx=5, pady=2)
        
        # 接続/切断ボタン + ステータス
        btn_frame_3499b = ttk.Frame(frame_3499b)
        btn_frame_3499b.grid(row=1, column=0, columnspan=3, pady=5, sticky=tk.W)
        
        self.btn_connect_3499b = ttk.Button(btn_frame_3499b, text="接続", 
                                             command=self.connect_3499b, width=12)
        self.btn_connect_3499b.pack(side=tk.LEFT, padx=5)
        
        self.btn_disconnect_3499b = ttk.Button(btn_frame_3499b, text="切断", 
                                                command=self.disconnect_3499b, 
                                                width=12, state=tk.DISABLED)
        self.btn_disconnect_3499b.pack(side=tk.LEFT, padx=5)
        
        # 接続状態表示（切断ボタンの横）
        self.status_3499b = ttk.Label(btn_frame_3499b, text="未接続", foreground="red")
        self.status_3499b.pack(side=tk.LEFT, padx=10)
        
        # ===== シリアル通信1フレーム =====
        serial_frame = ttk.LabelFrame(self, text="シリアル通信1", padding=10)
        serial_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 色ライン + UI
        content = tk.Frame(serial_frame, bd=0, highlightthickness=0)
        content.pack(fill="x")
        
        # 色ライン（通信1の色: #81D4FA）
        color_line = tk.Frame(content, bg="#81D4FA", width=8, height=28, bd=0, highlightthickness=0)
        color_line.pack(side="left", fill="y", padx=(0, 10))
        color_line.pack_propagate(False)
        
        # UI要素
        ui_frame = ttk.Frame(content)
        ui_frame.pack(side="left", fill="x", expand=True)
        
        # ポート選択
        port_frame = ttk.Frame(ui_frame)
        port_frame.pack(side="left", padx=(0, 10))
        ttk.Label(port_frame, text="ポート:").pack(side="left", padx=(0, 5))
        self.port_var = tk.StringVar()
        self.port_combo = ttk.Combobox(port_frame, textvariable=self.port_var, 
                                        state="readonly", width=15)
        self.port_combo.pack(side="left", padx=(0, 5))
        ttk.Button(port_frame, text="再スキャン", 
                   command=self.rescan_ports).pack(side="left")
        
        # 接続/切断ボタン
        button_frame = ttk.Frame(ui_frame)
        button_frame.pack(side="left", padx=(0, 10))
        ttk.Button(button_frame, text="接続", 
                   command=self.connect_serial).pack(side="left", padx=(0, 5))
        ttk.Button(button_frame, text="切断", 
                   command=self.disconnect_serial).pack(side="left")
        
        # ステータス
        status_frame = ttk.Frame(ui_frame)
        status_frame.pack(side="left")
        ttk.Label(status_frame, text="ステータス:").pack(side="left", padx=(0, 5))
        self.serial_status_label = ttk.Label(status_frame, text="未接続", 
                                              foreground="gray", width=20, anchor="w")
        self.serial_status_label.pack(side="left")
        
        # ログウィジェット（拡大）
        log_frame = ttk.LabelFrame(self, text="通信ログ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # LoggerWidgetは内部でpack()を実行するため、明示的なpack()は不要
        self.logger = LoggerWidget(log_frame, height=20)
        
        # 初期化
        self.rescan_ports()
    
    def save_config(self):
        """接続設定を保存(既存の設定を保持しながら更新)"""
        # 既存の設定ファイルを読み込む
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            except:
                config = {}
        else:
            config = {}
        
        # GPIB設定セクションを更新
        config["gpib"] = {
            "3458a_resource": self.resource_3458a_entry.get().strip(),
            "3499b_resource": self.resource_3499b_entry.get().strip()
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            self.logger.log("接続設定を保存しました", "SUCCESS")
        except Exception as e:
            self.logger.log(f"設定保存失敗: {str(e)}", "ERROR")
    
    def load_config(self):
        """接続設定を読み込み"""
        if not os.path.exists(self.config_file):
            self.logger.log("設定ファイルが見つかりません(初回起動)", "INFO")
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # GPIB設定セクションから読み込み
            if "gpib" in config:
                gpib_config = config["gpib"]
                
                if "3458a_resource" in gpib_config and gpib_config["3458a_resource"]:
                    self.resource_3458a_entry.delete(0, tk.END)
                    self.resource_3458a_entry.insert(0, gpib_config["3458a_resource"])
                
                if "3499b_resource" in gpib_config and gpib_config["3499b_resource"]:
                    self.resource_3499b_entry.delete(0, tk.END)
                    self.resource_3499b_entry.insert(0, gpib_config["3499b_resource"])
                
                self.logger.log("前回の接続設定を読み込みました", "SUCCESS")
        except Exception as e:
            self.logger.log(f"設定読み込み失敗: {str(e)}", "ERROR")
    
    def search_resources(self):
        """GPIB機器を検索"""
        # VISA初期化チェック(未初期化なら自動初期化)
        if self.gpib_3458a.rm is None:
            self.logger.log("VISA未初期化 - 自動初期化を実行します", "INFO")
            success, message = self.gpib_3458a.initialize()
            if not success:
                self.logger.log(f"VISA初期化失敗: {message}", "ERROR")
                return
            # 3499B側にも同じResourceManagerを設定
            self.gpib_3499b.rm = self.gpib_3458a.rm
            self.logger.log("VISA初期化成功", "SUCCESS")
        
        success, result = self.gpib_3458a.list_resources()
        
        if success:
            self.resource_list.delete(0, tk.END)
            if isinstance(result, list) and len(result) > 0:
                for resource in result:
                    self.resource_list.insert(tk.END, resource)
                self.logger.log(f"{len(result)}個の機器を検出", "SUCCESS")
            else:
                self.logger.log("機器が見つかりませんでした", "WARNING")
        else:
            self.logger.log(result, "ERROR")
    
    def set_resource_from_list(self, target_entry):
        """リストボックスで選択されたリソースをエントリーに設定"""
        selection = self.resource_list.curselection()
        if selection:
            selected_resource = self.resource_list.get(selection[0])
            target_entry.delete(0, tk.END)
            target_entry.insert(0, selected_resource)
            self.logger.log(f"リソース選択: {selected_resource}", "INFO")
    
    def connect_3458a(self):
        """3458Aに接続"""
        resource = self.resource_3458a_entry.get().strip()
        if not resource:
            self.logger.log("リソース名を入力してください", "ERROR")
            return
        
        # VISA初期化チェック(未初期化なら自動初期化)
        if self.gpib_3458a.rm is None:
            self.logger.log("VISA未初期化 - 自動初期化を実行します", "INFO")
            success, message = self.gpib_3458a.initialize()
            if not success:
                self.logger.log(f"VISA初期化失敗: {message}", "ERROR")
                return
            # 3499B側にも同じResourceManagerを設定
            self.gpib_3499b.rm = self.gpib_3458a.rm
            self.logger.log("VISA初期化成功", "SUCCESS")
        
        self.logger.log(f"3458A 接続中: {resource}", "INFO")
        # テストモード"none"で接続、device_type="3458A"を指定
        success, message = self.gpib_3458a.connect(
            resource, 
            timeout=5000, 
            test_mode="none",
            device_type="3458A"
        )
        level = "SUCCESS" if success else "ERROR"
        self.logger.log(f"3458A: {message}", level)
        
        # ボタン状態更新
        if success:
            self.btn_connect_3458a.config(state=tk.DISABLED)
            self.btn_disconnect_3458a.config(state=tk.NORMAL)
            self.status_3458a.config(text=f"接続中: {resource}", foreground="green")
            # 接続成功時に設定を保存
            self.save_config()
        else:
            self.status_3458a.config(text="接続失敗", foreground="red")
    
    def disconnect_3458a(self):
        """3458Aから切断(ローカルモードに復帰)"""
        success, message = self.gpib_3458a.disconnect(go_to_local=True)
        level = "SUCCESS" if success else "ERROR"
        self.logger.log(f"3458A: {message}", level)
        
        # ボタン状態更新
        self.btn_connect_3458a.config(state=tk.NORMAL)
        self.btn_disconnect_3458a.config(state=tk.DISABLED)
        self.status_3458a.config(text="未接続", foreground="red")
    
    def connect_3499b(self):
        """3499Bに接続"""
        resource = self.resource_3499b_entry.get().strip()
        if not resource:
            self.logger.log("リソース名を入力してください", "ERROR")
            return
        
        # VISA初期化チェック(未初期化なら自動初期化)
        if self.gpib_3499b.rm is None:
            self.logger.log("VISA未初期化 - 自動初期化を実行します", "INFO")
            # どちらか一方で初期化
            if self.gpib_3458a.rm is None:
                success, message = self.gpib_3458a.initialize()
                if not success:
                    self.logger.log(f"VISA初期化失敗: {message}", "ERROR")
                    return
                self.gpib_3499b.rm = self.gpib_3458a.rm
            else:
                # 3458A側が既に初期化済みならそれを使用
                self.gpib_3499b.rm = self.gpib_3458a.rm
            self.logger.log("VISA初期化成功", "SUCCESS")
        
        self.logger.log(f"3499B 接続中: {resource}", "INFO")
        # テストモード"none"で接続、device_type="3499B"を指定
        success, message = self.gpib_3499b.connect(
            resource, 
            timeout=5000, 
            test_mode="none",
            device_type="3499B"
        )
        level = "SUCCESS" if success else "ERROR"
        self.logger.log(f"3499B: {message}", level)
        
        # ボタン状態更新
        if success:
            self.btn_connect_3499b.config(state=tk.DISABLED)
            self.btn_disconnect_3499b.config(state=tk.NORMAL)
            self.status_3499b.config(text=f"接続中: {resource}", foreground="green")
            # 接続成功時に設定を保存
            self.save_config()
        else:
            self.status_3499b.config(text="接続失敗", foreground="red")
    
    def disconnect_3499b(self):
        """3499Bから切断(ローカルモードに復帰)"""
        success, message = self.gpib_3499b.disconnect(go_to_local=True)
        level = "SUCCESS" if success else "ERROR"
        self.logger.log(f"3499B: {message}", level)
        
        # ボタン状態更新
        self.btn_connect_3499b.config(state=tk.NORMAL)
        self.btn_disconnect_3499b.config(state=tk.DISABLED)
        self.status_3499b.config(text="未接続", foreground="red")
    
    # ========== シリアル通信関連メソッド ==========
    def rescan_ports(self):
        """シリアルポートを再スキャン"""
        ports = [p.device for p in serial.tools.list_ports.comports()]
        self.port_combo["values"] = ports
        if ports:
            current = self.port_var.get()
            if current not in ports:
                self.port_var.set(ports[0])
        else:
            self.port_var.set("")
        self.logger.log(f"ポートスキャン完了: {len(ports)}個のポートが見つかりました", "INFO")
    
    def connect_serial(self):
        """シリアルポートに接続"""
        selected_port = self.port_var.get()
        if not selected_port:
            self.serial_status_label.config(text="ポート未選択", foreground="red")
            self.logger.log("ポートを選択してください", "ERROR")
            return
        
        # 既存接続があれば切断
        if self.serial_mgr.is_connected():
            try:
                self.serial_mgr.disconnect()
            except Exception:
                pass
        
        result = self.serial_mgr.connect(selected_port)
        if result:
            self.serial_status_label.config(text=f"{selected_port} に接続中", foreground="green")
            self.logger.log(f"シリアル通信1: {selected_port} に接続しました", "SUCCESS")
        else:
            self.serial_status_label.config(text=f"{selected_port} 接続失敗", foreground="red")
            self.logger.log(f"シリアル通信1: {selected_port} への接続に失敗しました", "ERROR")
    
    def disconnect_serial(self):
        """シリアルポートを切断"""
        try:
            self.serial_mgr.disconnect()
            self.serial_status_label.config(text="未接続", foreground="gray")
            self.logger.log("シリアル通信1: 切断しました", "SUCCESS")
        except Exception as e:
            self.logger.log(f"切断エラー: {e}", "ERROR")