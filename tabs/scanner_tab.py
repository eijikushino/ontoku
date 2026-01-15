import tkinter as tk
from tkinter import ttk
from utils import LoggerWidget
import json
import time
import os

class ScannerTab(ttk.Frame):
    SETTINGS_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "app_settings.json")

    def __init__(self, parent, gpib_controller):
        super().__init__(parent)
        self.gpib = gpib_controller
        self.channel_vars = {}  # チャンネルの状態を保持 (True=CLOSED, False=OPEN)

        # 設定を読み込み
        self.relay_delay = self.load_relay_delay()

        self.create_widgets()
    
    def create_widgets(self):
        # スロット選択フレーム
        slot_frame = ttk.LabelFrame(self, text="スロット設定", padding=10)
        slot_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Label(slot_frame, text="スロット番号:").pack(side=tk.LEFT, padx=5)
        self.slot_var = tk.StringVar(value="1")
        slot_combo = ttk.Combobox(slot_frame, textvariable=self.slot_var, 
                                   values=["1", "2", "3", "4", "5", "6", "7", "8"], 
                                   width=5, state="readonly")
        slot_combo.pack(side=tk.LEFT, padx=5)
        slot_combo.bind("<<ComboboxSelected>>", self.on_slot_changed)
        
        ttk.Label(slot_frame, text="モジュール: N2270A (10ch)",
                  foreground="blue").pack(side=tk.LEFT, padx=10)

        # リレー切替待ち時間設定フレーム
        delay_frame = ttk.LabelFrame(self, text="リレー切替設定", padding=10)
        delay_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Label(delay_frame, text="OPEN→CLOSE 待ち時間:").pack(side=tk.LEFT, padx=5)
        self.delay_var = tk.DoubleVar(value=self.relay_delay)
        self.delay_spinbox = ttk.Spinbox(
            delay_frame,
            from_=0.5,
            to=10.0,
            increment=0.1,
            textvariable=self.delay_var,
            width=6,
            format="%.1f",
            command=self.on_delay_changed
        )
        self.delay_spinbox.pack(side=tk.LEFT, padx=5)
        self.delay_spinbox.bind("<Return>", lambda e: self.on_delay_changed())
        self.delay_spinbox.bind("<FocusOut>", lambda e: self.on_delay_changed())
        ttk.Label(delay_frame, text="秒").pack(side=tk.LEFT)

        # チャンネル選択フレーム (レ点=CLOSE, 空=OPEN)
        channel_frame = ttk.LabelFrame(self, text="チャンネル選択 (CH00-CH09) ※レ点=CLOSE", padding=10)
        channel_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 10チャンネル x 5列
        channels_per_row = 5
        total_channels = 10
        self.channel_vars = {}

        for i in range(0, total_channels):
            row = i // channels_per_row
            col = i % channels_per_row

            var = tk.BooleanVar(value=False)  # False=OPEN, True=CLOSED
            self.channel_vars[i] = var

            cb = ttk.Checkbutton(
                channel_frame,
                text=f"CH{i:02d}",
                variable=var,
                command=lambda ch=i: self.toggle_channel(ch)
            )
            cb.grid(row=row, column=col, padx=15, pady=5, sticky=tk.W)
        
        # エラー確認フレーム
        error_frame = ttk.LabelFrame(self, text="エラー確認", padding=10)
        error_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(error_frame, text="エラー詳細確認", 
                   command=self.check_error).pack(side=tk.LEFT, padx=5)
        ttk.Button(error_frame, text="エラークリア", 
                   command=self.clear_error).pack(side=tk.LEFT, padx=5)
        ttk.Button(error_frame, text="ステータス確認", 
                   command=self.check_status).pack(side=tk.LEFT, padx=5)
        
        # エラー表示ラベル
        self.error_label = ttk.Label(error_frame, text="エラー: 未確認", 
                                      foreground="gray", font=("", 10, "bold"))
        self.error_label.pack(side=tk.LEFT, padx=20)
        
        # ログ表示フレーム
        log_frame = ttk.LabelFrame(self, text="ログ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.logger = LoggerWidget(log_frame, height=10)
        
        # 使用方法の表示
        self.logger.log("N2270A リレーモジュール制御", "INFO")
        self.logger.log("チャンネル: CH00-09 → @X00-@X09 (Xはスロット番号)", "INFO")
    
    def on_slot_changed(self, event=None):
        """スロット変更時の処理"""
        slot = self.slot_var.get()
        self.logger.log(f"スロット {slot} を選択しました", "INFO")
    
    def get_channel_address(self, channel):
        """チャンネルアドレスを生成
        CH00-09 → @100-@109 (スロット1の場合)
        """
        slot = self.slot_var.get()
        return f"@{slot}{channel:02d}"
    
    def toggle_channel(self, channel):
        """チャンネルのCLOSE/OPENをトグル（レ点=CLOSE, 空=OPEN）単一選択のみ"""
        is_checked = self.channel_vars[channel].get()
        address = self.get_channel_address(channel)

        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            self.channel_vars[channel].set(not is_checked)
            return

        if is_checked:
            # レ点が入った → 他のチャンネルをOPENしてからCLOSE
            # 他のチェックを外す（単一選択）
            opened_any = False
            for ch, var in self.channel_vars.items():
                if ch != channel and var.get():
                    other_addr = self.get_channel_address(ch)
                    self.gpib.write(f"OPEN ({other_addr})")
                    var.set(False)
                    self.logger.log(f"CH{ch:02d} ({other_addr}) をOPENしました", "INFO")
                    opened_any = True

            # OPEN後の待ち時間
            if opened_any:
                self.logger.log(f"リレー切替待機中... ({self.relay_delay:.1f}秒)", "INFO")
                self.update_idletasks()
                time.sleep(self.relay_delay)

            # 選択したチャンネルをCLOSE
            command = f"CLOSE ({address})"
            success, message = self.gpib.write(command)
            if success:
                self.logger.log(f"CH{channel:02d} ({address}) をCLOSEしました", "SUCCESS")
            else:
                self.logger.log(f"CH{channel:02d} のCLOSEに失敗: {message}", "ERROR")
                self.channel_vars[channel].set(False)
        else:
            # レ点が外れた → OPEN
            command = f"OPEN ({address})"
            success, message = self.gpib.write(command)
            if success:
                self.logger.log(f"CH{channel:02d} ({address}) をOPENしました", "SUCCESS")
            else:
                self.logger.log(f"CH{channel:02d} のOPENに失敗: {message}", "ERROR")
                self.channel_vars[channel].set(True)

    def check_error(self):
        """エラー詳細を確認"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        self.logger.log("エラー確認中...", "INFO")
        
        # SYST:ERR?コマンドでエラーを取得
        success, response = self.gpib.query("SYST:ERR?")
        
        if success:
            error_text = response.strip()
            self.logger.log(f"エラー詳細: {error_text}", "INFO")
            
            # エラーコードを解析
            if error_text.startswith("0,") or error_text.startswith("+0,") or "No error" in error_text:
                self.error_label.config(text="エラー: なし", foreground="green")
                self.logger.log("エラーはありません", "SUCCESS")
            else:
                self.error_label.config(text=f"エラー: あり", foreground="red")
                self.logger.log(f"エラー検出: {error_text}", "ERROR")
                
                # 追加のエラー情報を取得
                self.get_additional_error_info()
        else:
            self.logger.log("エラー確認コマンド失敗", "ERROR")
            self.error_label.config(text="エラー: 確認失敗", foreground="red")
    
    def get_additional_error_info(self):
        """追加のエラー情報を取得（エラーキューを空にする）"""
        # すべてのエラーを読み出す（エラーキューが空になるまで）
        error_count = 1
        while error_count < 10:  # 最大10個まで
            success, response = self.gpib.query("SYST:ERR?")
            if success:
                error_text = response.strip()
                if error_text.startswith("0,") or error_text.startswith("+0,"):
                    break  # エラーキューが空になった
                self.logger.log(f"追加エラー[{error_count}]: {error_text}", "ERROR")
                error_count += 1
            else:
                break
    
    def clear_error(self):
        """エラーをクリア"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        # *CLSコマンドでエラーをクリア
        success, message = self.gpib.write("*CLS")
        
        if success:
            self.logger.log("エラーをクリアしました", "SUCCESS")
            self.error_label.config(text="エラー: クリア済み", foreground="blue")
            
            # エラーキューを空にする
            for _ in range(10):
                success, response = self.gpib.query("SYST:ERR?")
                if not success:
                    break
                if response.startswith("0,") or response.startswith("+0,"):
                    break
        else:
            self.logger.log(f"エラークリア失敗: {message}", "ERROR")
    
    def check_status(self):
        """機器のステータスを確認"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        self.logger.log("ステータス確認中...", "INFO")
        
        # Standard Event Status Registerを確認
        success, response = self.gpib.query("*ESR?")
        if success:
            esr = response.strip()
            self.logger.log(f"Event Status Register: {esr}", "INFO")
            try:
                self.decode_esr(int(esr) if esr.isdigit() else 0)
            except ValueError:
                self.logger.log(f"ESR値の解析失敗: {esr}", "WARNING")
        else:
            self.logger.log("ESR確認失敗", "ERROR")
        
        # Status Byte Registerを確認
        success, response = self.gpib.query("*STB?")
        if success:
            stb = response.strip()
            self.logger.log(f"Status Byte: {stb}", "INFO")
        else:
            self.logger.log("STB確認失敗", "ERROR")
    
    def decode_esr(self, esr_value):
        """Event Status Registerをデコード"""
        status_bits = {
            0: "Operation Complete",
            2: "Query Error",
            3: "Device Error",
            4: "Execution Error",
            5: "Command Error",
            7: "Power On"
        }
        
        if esr_value == 0:
            self.logger.log("  ステータス: 正常", "SUCCESS")
        else:
            for bit, description in status_bits.items():
                if esr_value & (1 << bit):
                    self.logger.log(f"  ビット{bit}: {description}", "WARNING")

    # ---------- 設定読み込み・保存 ----------
    def load_relay_delay(self):
        """設定ファイルからリレー待ち時間を読み込み"""
        try:
            with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                settings = json.load(f)
                return settings.get("scanner", {}).get("relay_switch_delay", 0.5)
        except (FileNotFoundError, json.JSONDecodeError):
            return 0.5

    def save_relay_delay(self, delay_value):
        """リレー待ち時間を設定ファイルに保存"""
        try:
            # 既存の設定を読み込み
            try:
                with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                    settings = json.load(f)
            except (FileNotFoundError, json.JSONDecodeError):
                settings = {}

            # scanner設定を更新
            if "scanner" not in settings:
                settings["scanner"] = {}
            settings["scanner"]["relay_switch_delay"] = delay_value

            # 保存
            with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
        except Exception as e:
            self.logger.log(f"設定保存エラー: {e}", "ERROR")

    def on_delay_changed(self):
        """待ち時間変更時の処理"""
        try:
            value = self.delay_var.get()
            # 最小値チェック
            if value < 0.5:
                value = 0.5
                self.delay_var.set(value)
            self.relay_delay = value
            self.save_relay_delay(value)
            self.logger.log(f"リレー切替待ち時間を {value:.1f}秒 に設定しました", "INFO")
        except tk.TclError:
            # 無効な値が入力された場合
            self.delay_var.set(self.relay_delay)