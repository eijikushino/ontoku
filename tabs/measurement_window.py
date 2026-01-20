import tkinter as tk
from tkinter import ttk, messagebox
import time
import json
import os
import sys
import threading
import queue

# プロジェクトルートをパスに追加
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if project_root not in sys.path:
    sys.path.insert(0, project_root)

# CSV保存用ロガーのインポート
from utils.csv_logger import MeasurementCSVLogger


class MeasurementWindow(tk.Toplevel):
    """Pattern Test用の計測ウィンドウ（シンプル版 + CSV保存機能）"""
    
    def __init__(self, parent, gpib_3458a, gpib_3499b, test_tab):
        super().__init__(parent)
        
        self.title("パターン計測")
        self.geometry("700x600")
        
        self.parent_app = parent
        self.gpib_dmm = gpib_3458a
        self.gpib_scanner = gpib_3499b
        self.test_tab = test_tab
        
        self.is_measuring = False
        self.measurement_count = 0
        
        # ★★★ 前回のスキャナ切替時間を読み込み ★★★
        saved_switch_delay = self._load_switch_delay()
        self.switch_delay_sec = tk.DoubleVar(value=saved_switch_delay)

        self.scanner_slot = "1"
        self.last_closed_channel = None
        self.update_timer_id = None
        
        # ★★★ DMM設定情報を保持 ★★★
        self.dmm_mode = "---"
        self.dmm_range = "---"
        self.dmm_nplc = "---"
        
        # ★★★ CSV保存機能 ★★★
        self.csv_logger = None
        self.is_csv_logging = False

        # ★★★ パターン実行同期オプション ★★★
        self.sync_with_pattern_var = tk.BooleanVar(value=self._load_sync_option())
        self.last_pattern_running_state = False  # パターン実行状態の前回値

        # ★★★ パターン切替前の待機状態管理 ★★★
        self.is_waiting_for_pattern_change = False  # パターン切替待機中フラグ
        self.waiting_pattern_index = -1  # 待機開始時のパターンインデックス
        self.waiting_selected_defs = None  # 待機中の選択DEF
        self.waiting_def_index = 0  # 待機中のDEFインデックス
        self.waiting_pole = "Pos"  # 待機中の極性

        # ★★★ スレッド化用: 結果キューとロック（スキャナーとDMMで分離）★★★
        self.scanner_queue = queue.Queue()
        self.dmm_queue = queue.Queue()
        self.measurement_lock = threading.Lock()

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.create_widgets()

        # パターン情報の定期更新を開始
        self.start_pattern_info_update()

        # 測定間隔目安の初期表示と定期更新
        self._update_estimate()
        self._start_estimate_update()
        
    def create_widgets(self):
        """ウィジェットを作成"""
        
        # 設定エリア
        config_frame = ttk.LabelFrame(self, text="計測設定", padding=10)
        config_frame.pack(fill=tk.X, padx=10, pady=5)
        
        delay_frame = ttk.Frame(config_frame)
        delay_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(delay_frame, text="スキャナー切替時間:").pack(side=tk.LEFT, padx=5)
        # 最小0.5秒: リレー切替の物理的な時間を確保（OPEN後とCLOSE後に半分ずつ使用）
        switch_spinbox = ttk.Spinbox(delay_frame, from_=0.5, to=10.0, increment=0.1,
                                     textvariable=self.switch_delay_sec, width=6, format="%.1f",
                                     command=self._on_delay_changed)
        switch_spinbox.pack(side=tk.LEFT, padx=2)
        switch_spinbox.bind('<FocusOut>', lambda e: self._on_delay_changed())
        ttk.Label(delay_frame, text="sec").pack(side=tk.LEFT)

        # 測定間隔目安表示
        estimate_frame = ttk.Frame(config_frame)
        estimate_frame.pack(fill=tk.X, pady=5)
        ttk.Label(estimate_frame, text="測定間隔目安:").pack(side=tk.LEFT, padx=5)
        self.estimate_label = ttk.Label(estimate_frame, text="---", foreground="green", font=("", 9, "bold"))
        self.estimate_label.pack(side=tk.LEFT, padx=5)
        
        # ★★★ DMM設定表示エリアを追加 ★★★
        dmm_settings_frame = ttk.Frame(config_frame)
        dmm_settings_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(dmm_settings_frame, text="DMM設定:").pack(side=tk.LEFT, padx=5)
        ttk.Label(dmm_settings_frame, text="Mode:").pack(side=tk.LEFT, padx=(15, 2))
        self.dmm_mode_label = ttk.Label(dmm_settings_frame, text="---", 
                                         foreground="blue", font=("", 9, "bold"))
        self.dmm_mode_label.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(dmm_settings_frame, text="Range:").pack(side=tk.LEFT, padx=(10, 2))
        self.dmm_range_label = ttk.Label(dmm_settings_frame, text="---", 
                                          foreground="blue", font=("", 9, "bold"))
        self.dmm_range_label.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(dmm_settings_frame, text="NPLC:").pack(side=tk.LEFT, padx=(10, 2))
        self.dmm_nplc_label = ttk.Label(dmm_settings_frame, text="---", 
                                         foreground="blue", font=("", 9, "bold"))
        self.dmm_nplc_label.pack(side=tk.LEFT, padx=2)
        
        # 実行制御
        control_frame = ttk.LabelFrame(self, text="実行制御", padding=10)
        control_frame.pack(fill=tk.X, padx=10, pady=5)

        # ボタン行
        button_row = ttk.Frame(control_frame)
        button_row.pack(fill=tk.X)

        self.start_button = ttk.Button(button_row, text="計測開始", command=self.start_measurement, width=10)
        self.start_button.pack(side=tk.LEFT, padx=2)

        self.stop_button = ttk.Button(button_row, text="計測停止", command=self.stop_measurement,
                                      state=tk.DISABLED, width=10)
        self.stop_button.pack(side=tk.LEFT, padx=2)

        # ★★★ CSV保存コントロール（ボタン+チェックボックスを縦に配置）★★★
        csv_control_frame = ttk.Frame(button_row)
        csv_control_frame.pack(side=tk.LEFT, padx=(15, 2))

        # CSV保存ボタン行
        csv_button_row = ttk.Frame(csv_control_frame)
        csv_button_row.pack()

        self.csv_start_button = ttk.Button(csv_button_row, text="保存開始",
                                            command=self.start_csv_logging, state=tk.DISABLED, width=10)
        self.csv_start_button.pack(side=tk.LEFT, padx=(0, 2))

        self.csv_stop_button = ttk.Button(csv_button_row, text="保存終了",
                                           command=self.stop_csv_logging, state=tk.DISABLED, width=10)
        self.csv_stop_button.pack(side=tk.LEFT)

        # パターン実行同期オプション（保存ボタンの下）
        self.sync_checkbox = ttk.Checkbutton(
            csv_control_frame,
            text="パターン実行と保存を同期",
            variable=self.sync_with_pattern_var,
            command=self._on_sync_option_changed
        )
        self.sync_checkbox.pack(anchor=tk.W)

        # ステータス表示
        self.count_label = ttk.Label(button_row, text="計測回数: 0", font=("", 10, "bold"))
        self.count_label.pack(side=tk.LEFT, padx=15)

        # 保存中ステータス表示
        self.csv_status_label = ttk.Label(button_row, text="", font=("", 10, "bold"))
        self.csv_status_label.pack(side=tk.LEFT, padx=5)

        # 状態表示
        status_frame = ttk.LabelFrame(self, text="現在の状態", padding=10)
        status_frame.pack(fill=tk.X, padx=10, pady=5)
        
        info_frame = ttk.Frame(status_frame)
        info_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(info_frame, text="DEF:", font=("", 11)).pack(side=tk.LEFT, padx=5)
        self.def_label = ttk.Label(info_frame, text="---", font=("", 14, "bold"), foreground="blue")
        self.def_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(info_frame, text="Pole:", font=("", 11)).pack(side=tk.LEFT, padx=15)
        self.pole_label = ttk.Label(info_frame, text="---", font=("", 14, "bold"), foreground="green")
        self.pole_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(info_frame, text="CH:", font=("", 11)).pack(side=tk.LEFT, padx=15)
        self.channel_label = ttk.Label(info_frame, text="---", font=("", 12, "bold"), foreground="purple")
        self.channel_label.pack(side=tk.LEFT, padx=5)
        
        data_frame = ttk.Frame(status_frame)
        data_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(data_frame, text="Pattern No:").pack(side=tk.LEFT, padx=5)
        self.pattern_no_label = ttk.Label(data_frame, text="---", foreground="red", font=("", 10, "bold"))
        self.pattern_no_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(data_frame, text="DataSet:").pack(side=tk.LEFT, padx=15)
        self.dataset_label = ttk.Label(data_frame, text="---", foreground="darkblue")
        self.dataset_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(data_frame, text="Code:").pack(side=tk.LEFT, padx=15)
        self.code_label = ttk.Label(data_frame, text="---", foreground="darkgreen")
        self.code_label.pack(side=tk.LEFT, padx=5)
        
        ttk.Separator(status_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)
        
        measurement_frame = ttk.Frame(status_frame)
        measurement_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(measurement_frame, text="計測値:", font=("", 11)).pack(side=tk.LEFT, padx=5)
        self.measurement_label = ttk.Label(measurement_frame, text="--- V", font=("", 16, "bold"), foreground="red")
        self.measurement_label.pack(side=tk.LEFT, padx=5)
        
        # ログ
        log_frame = ttk.LabelFrame(self, text="ログ", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # 詳細モードチェックボックス
        log_option_frame = ttk.Frame(log_frame)
        log_option_frame.pack(fill=tk.X, pady=(0, 5))
        self.detail_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(log_option_frame, text="詳細モード (OPEN/CLOSE表示)",
                        variable=self.detail_mode_var).pack(side=tk.LEFT)

        log_scroll = ttk.Scrollbar(log_frame)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=15, width=80, yscrollcommand=log_scroll.set,
                                state=tk.DISABLED, font=("Courier", 9))
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.config(command=self.log_text.yview)
        
        self.log_text.tag_config("INFO", foreground="black")
        self.log_text.tag_config("SUCCESS", foreground="green")
        self.log_text.tag_config("WARNING", foreground="orange")
        self.log_text.tag_config("ERROR", foreground="red")
        
        self.log("計測ウィンドウを初期化", "INFO")
        
    def log(self, message, level="INFO"):
        """ログ出力"""
        ms = int(time.time() * 1000) % 1000
        timestamp = time.strftime("%H:%M:%S") + f".{ms:03d}"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.update_idletasks()
    
    def get_selected_defs(self):
        """選択DEF取得"""
        selected_defs = []
        
        if hasattr(self.test_tab, 'def_check_vars'):
            for i, check_var in enumerate(self.test_tab.def_check_vars):
                if check_var.get():
                    pos_ch = self.test_tab.scanner_channels_pos[i].get() if i < len(self.test_tab.scanner_channels_pos) else "ー"
                    neg_ch = self.test_tab.scanner_channels_neg[i].get() if i < len(self.test_tab.scanner_channels_neg) else "ー"
                    
                    if (pos_ch == "ー" or pos_ch == "") and (neg_ch == "ー" or neg_ch == ""):
                        self.log(f"DEF{i}: 未設定のため除外", "WARNING")
                        continue
                    
                    selected_defs.append({'index': i, 'name': f"DEF{i}", 'pos_channel': pos_ch, 'neg_channel': neg_ch})
        
        return selected_defs
    
    def start_measurement(self):
        """計測開始"""
        self.log("=== 計測開始 ===", "INFO")
        
        if not self.gpib_dmm.connected or not self.gpib_scanner.connected:
            messagebox.showerror("エラー", "機器が接続されていません")
            self.log("計測開始失敗: 機器未接続", "ERROR")
            return
        
        selected_defs = self.get_selected_defs()
        if not selected_defs:
            messagebox.showwarning("警告", "DEFが選択されていません")
            return
        
        # ★★★ DMM設定を取得 ★★★
        self.log("DMM設定を取得中...", "INFO")
        if not self.get_dmm_settings():
            if not messagebox.askyesno("警告", "DMM設定の取得に失敗しました。計測を続行しますか?"):
                return
        
        # 【単一選択の初期化】計測開始時に使用する全チャンネルをOPEN
        # これにより計測開始時点で複数チャンネルがCLOSEされている状態を防ぐ
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000

        try:
            for def_entry in selected_defs:
                pos_ch = def_entry['pos_channel']
                neg_ch = def_entry['neg_channel']

                if pos_ch != "ー" and pos_ch != "":
                    pos_num = pos_ch.replace("CH", "")
                    pos_addr = f"@{self.scanner_slot}{pos_num}"
                    self.gpib_scanner.write(f"OPEN ({pos_addr})")
                    time.sleep(0.05)

                if neg_ch != "ー" and neg_ch != "":
                    neg_num = neg_ch.replace("CH", "")
                    neg_addr = f"@{self.scanner_slot}{neg_num}"
                    self.gpib_scanner.write(f"OPEN ({neg_addr})")
                    time.sleep(0.05)
        finally:
            self.gpib_scanner.instrument.timeout = orig_timeout

        # 【単一選択】CLOSEチャンネル追跡をリセット（全てOPEN状態から開始）
        self.last_closed_channel = None
        
        self.is_measuring = True
        self.measurement_count = 0
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)

        # ★★★ CSV保存ボタンを有効化 ★★★
        self.csv_start_button.config(state=tk.NORMAL)

        # ★★★ パターン実行状態の追跡をリセット（現在の状態で初期化）★★★
        self.last_pattern_running_state = (
            hasattr(self.test_tab, 'is_running') and self.test_tab.is_running
        )

        self.log(f"選択DEF数: {len(selected_defs)}, 切替時間: {self.switch_delay_sec.get()}sec", "INFO")
        
        # 少し待機してから計測開始
        self.after(100, lambda: self.do_one_measurement(selected_defs, 0, "Pos"))
    
    def stop_measurement(self):
        """計測停止"""
        self.log("=== 計測停止 ===", "WARNING")
        self.is_measuring = False
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)

        # ★★★ パターン切替待機中なら解除 ★★★
        if self.is_waiting_for_pattern_change:
            self._end_waiting_for_pattern_change()

        # ★★★ CSV保存中なら自動停止 ★★★
        if self.is_csv_logging:
            self.stop_csv_logging(show_dialog=False)
        else:
            # CSV保存していない場合もDEF選択を有効化
            self._lock_def_checkboxes(False)

        # ★★★ CSV保存ボタンを無効化 ★★★
        self.csv_start_button.config(state=tk.DISABLED)

        # ★★★ 停止時に全チャンネルをOPEN ★★★
        self.open_all_used_channels()
    
    def do_one_measurement(self, selected_defs, def_index, pole):
        """1回の計測実行"""

        if not self.is_measuring:
            return

        if def_index >= len(selected_defs):
            self.log(f"=== 1周完了 === (CSV保存中: {self.is_csv_logging})", "INFO")

            # ★★★ パターン切替前の待機チェック（CSV保存中のみ） ★★★
            if self.is_csv_logging:
                should_wait = self._should_wait_for_pattern_change()
                if should_wait:
                    self._start_waiting_for_pattern_change(selected_defs)
                    return

            self.do_one_measurement(selected_defs, 0, "Pos")
            return
        
        def_info = selected_defs[def_index]
        channel = def_info['pos_channel'] if pole == "Pos" else def_info['neg_channel']
        
        if channel == "ー" or channel == "":
            self.log(f"{def_info['name']} {pole} - CH未設定、スキップ", "WARNING")
            next_idx, next_pole = self.get_next(def_index, pole, len(selected_defs))
            self.after(100, lambda: self.do_one_measurement(selected_defs, next_idx, next_pole))
            return
        
        self.update_display(def_info, pole, channel)
        self.execute_measurement(selected_defs, def_index, pole, def_info, channel)
    
    def update_display(self, def_info, pole, channel):
        """表示更新"""
        self.def_label.config(text=def_info['name'])
        self.pole_label.config(text=pole)
        
        ch_number = channel.replace("CH", "")
        self.channel_label.config(text=f"@{self.scanner_slot}{ch_number}")
        
        # ★★★ パターン情報は定期更新タイマーに任せるため、ここでは更新しない ★★★
    
    def execute_measurement(self, selected_defs, def_index, pole, def_info, channel):
        """計測実行（非同期版）- スキャナー切替を別スレッドで実行"""
        import time as time_module

        ch_number = channel.replace("CH", "")
        channel_addr = f"@{self.scanner_slot}{ch_number}"

        # スキャナー切替開始時刻を記録
        self._scanner_start_time = time_module.time()

        # 別スレッドでスキャナー切り替えを実行
        thread = threading.Thread(
            target=self._scanner_switch_worker,
            args=(channel_addr, def_info, pole),
            daemon=True
        )
        thread.start()

        # 結果のポーリング開始
        self.after(10, lambda: self._check_scanner_result(selected_defs, def_index, pole, def_info))

    def _scanner_switch_worker(self, channel_addr, def_info, pole):
        """スキャナー切り替えのワーカースレッド

        【単一選択の保証】
        - 常に1つのチャンネルのみCLOSE状態を維持
        - 新しいチャンネルをCLOSEする前に、前回のチャンネルを必ずOPEN
        - OPENが失敗した場合はCLOSEを実行しない（複数CLOSE防止）
        """
        result = {
            'success': False,
            'channel_addr': channel_addr,
            'error': None,
            'detail': []  # 詳細ログ用
        }

        try:
            with self.measurement_lock:
                orig_timeout = self.gpib_scanner.instrument.timeout
                self.gpib_scanner.instrument.timeout = 5000
                switch_delay = self.switch_delay_sec.get()

                try:
                    # 【単一選択】前回CLOSEしたチャンネルをOPEN（複数CLOSE防止）
                    if self.last_closed_channel is not None:
                        open_success, _ = self.gpib_scanner.write(f"OPEN ({self.last_closed_channel})")
                        wait_time = switch_delay / 2
                        result['detail'].append(f"OPEN ({self.last_closed_channel}) → 待機 {wait_time:.2f}秒")
                        time.sleep(wait_time)

                        # OPENが失敗した場合、CLOSEを実行しない（複数チャンネルCLOSE防止）
                        if not open_success:
                            result['error'] = 'open_previous_failed'
                            self.scanner_queue.put(result)
                            return

                    # 【単一選択】新しいチャンネルをCLOSE（常に1つだけ）
                    success, _ = self.gpib_scanner.write(f"CLOSE ({channel_addr})")
                    result['detail'].append(f"CLOSE ({channel_addr})")
                    if success:
                        result['success'] = True
                    else:
                        result['error'] = 'close_failed'
                finally:
                    self.gpib_scanner.instrument.timeout = orig_timeout
        except Exception as e:
            result['error'] = str(e)

        self.scanner_queue.put(result)

    def _check_scanner_result(self, selected_defs, def_index, pole, def_info):
        """スキャナー切り替え結果をチェック

        【単一選択の管理】
        - 成功時: last_closed_channelを更新（次回OPEN対象として記憶）
        - 失敗時: 計測停止（不整合な状態を防ぐ）
        """
        import time as time_module

        if not self.is_measuring:
            return

        try:
            result = self.scanner_queue.get_nowait()

            # スキャナー切替時間を計算
            scanner_elapsed = 0
            if hasattr(self, '_scanner_start_time'):
                scanner_elapsed = time_module.time() - self._scanner_start_time

            if result['success']:
                # 【単一選択】今回CLOSEしたチャンネルを記憶（次回の切替時にOPENする対象）
                self.last_closed_channel = result['channel_addr']

                # 詳細モードの場合、OPEN/CLOSEの詳細をログ出力
                if self.detail_mode_var.get() and result.get('detail'):
                    for detail in result['detail']:
                        self.log(f"  {detail}", "INFO")

                self.log(f"スキャナー切替OK: {def_info['name']} {pole}", "SUCCESS")

                # 詳細モードの場合、スキャナー切替時間を表示
                if self.detail_mode_var.get():
                    self.log(f"  スキャナー切替時間: {scanner_elapsed:.3f}秒", "INFO")

                # スキャナー切替時間の1/2: CLOSE後、DMM計測開始までの待ち時間
                delay_ms = int(self.switch_delay_sec.get() / 2 * 1000)
                if self.detail_mode_var.get():
                    self.log(f"  CLOSE後待機 {delay_ms/1000:.2f}秒 → DMM計測", "INFO")
                self.after(delay_ms, lambda: self.do_dmm_measurement(selected_defs, def_index, pole, def_info))
            else:
                self.log(f"CLOSE失敗: {result.get('error', 'unknown')}", "ERROR")
                self.stop_measurement()

        except queue.Empty:
            # まだ結果がない → 再チェック（UIは応答可能）
            self.after(20, lambda: self._check_scanner_result(selected_defs, def_index, pole, def_info))
    
    def do_dmm_measurement(self, selected_defs, def_index, pole, def_info):
        """DMM計測（非同期版）- 別スレッドで実行してUIをブロックしない"""
        import time as time_module

        # 計測開始時刻を記録
        self._dmm_start_time = time_module.time()

        # 別スレッドで測定を実行
        thread = threading.Thread(
            target=self._dmm_measure_worker,
            args=(selected_defs, def_index, pole, def_info),
            daemon=True
        )
        thread.start()

        # 結果のポーリング開始
        self.after(10, lambda: self._check_dmm_result(selected_defs, def_index, pole, def_info))

    def _dmm_measure_worker(self, selected_defs, def_index, pole, def_info):
        """DMM測定のワーカースレッド"""
        result = {
            'success': False,
            'response': None,
            'error': None
        }

        try:
            with self.measurement_lock:
                orig_timeout = self.gpib_dmm.instrument.timeout
                self.gpib_dmm.instrument.timeout = 3000

                try:
                    # query()を使用してwrite+readを1回の通信にまとめる（高速化）
                    success, response = self.gpib_dmm.query("TRIG SGL")
                    if success:
                        result['success'] = True
                        result['response'] = response
                    else:
                        result['error'] = 'query_failed'
                finally:
                    self.gpib_dmm.instrument.timeout = orig_timeout
        except Exception as e:
            result['error'] = str(e)

        self.dmm_queue.put(result)

    def _check_dmm_result(self, selected_defs, def_index, pole, def_info):
        """DMM測定結果をチェック・UI更新"""
        import time as time_module

        if not self.is_measuring:
            return

        try:
            result = self.dmm_queue.get_nowait()

            # DMM計測時間を計算
            dmm_elapsed = 0
            if hasattr(self, '_dmm_start_time'):
                dmm_elapsed = time_module.time() - self._dmm_start_time

            if result['success']:
                value = result['response'].strip()
                self.measurement_label.config(text=f"{value} V")
                self.measurement_count += 1
                self.count_label.config(text=f"計測回数: {self.measurement_count}")
                self.log(f"計測OK ({self.measurement_count}): {def_info['name']} {pole} = {value} V", "SUCCESS")

                # 詳細モードの場合、DMM計測時間を表示
                if self.detail_mode_var.get():
                    self.log(f"  DMM計測時間: {dmm_elapsed:.3f}秒", "INFO")

                # ★★★ CSV保存中なら測定値を記録（全データ保存） ★★★
                if self.is_csv_logging and self.csv_logger:
                    is_cycle_start = (def_index == 0 and pole == "Pos")

                    # 現在のパターン情報を取得
                    current_pattern = self.get_current_pattern_info()
                    dataset = current_pattern['dataset']
                    code = current_pattern['code']

                    # 全データを記録（スキップなし）
                    pole_upper = "POS" if pole == "Pos" else "NEG"
                    self.csv_logger.record_measurement(
                        def_info['index'],
                        pole_upper,
                        value,
                        is_cycle_start=is_cycle_start,
                        dataset=dataset,
                        code=code
                    )
            else:
                error_msg = result.get('error', 'unknown')
                if error_msg == 'query_failed':
                    self.log("TRIG SGL失敗", "ERROR")
                else:
                    self.log(f"計測失敗: {error_msg}", "ERROR")

            # 次の測定へ
            next_idx, next_pole = self.get_next(def_index, pole, len(selected_defs))
            self.do_one_measurement(selected_defs, next_idx, next_pole)

        except queue.Empty:
            # まだ結果がない → 再チェック（UIは応答可能）
            self.after(20, lambda: self._check_dmm_result(selected_defs, def_index, pole, def_info))
    
    def get_next(self, idx, pole, total):
        """次の位置"""
        return (idx, "Neg") if pole == "Pos" else (idx + 1, "Pos")
    
    # ★★★ CSV保存機能 ★★★
    def start_csv_logging(self):
        """CSV保存開始"""
        if self.is_csv_logging:
            messagebox.showwarning("警告", "既にCSV保存中です")
            return
        
        # 設定ファイルから保存先とシリアルNo.を取得
        save_dir, filename, serial_numbers = self._load_save_config()
        
        if not serial_numbers:
            messagebox.showwarning("警告", "シリアルNo.が設定されていません\nファイル保存タブで設定してください")
            return
        
        # CSVロガーを初期化
        self.csv_logger = MeasurementCSVLogger(save_dir, filename, serial_numbers)
        
        # ログ記録開始
        success, message = self.csv_logger.start_logging()
        
        if success:
            self.is_csv_logging = True
            self.csv_start_button.config(state=tk.DISABLED)
            self.csv_stop_button.config(state=tk.NORMAL)
            self.csv_status_label.config(text="■ CSV保存中", foreground="red")

            # ★★★ DEF選択のチェックボックスを無効化 ★★★
            self._lock_def_checkboxes(True)

            self.log(message, "SUCCESS")
        else:
            self.log(message, "ERROR")
            messagebox.showerror("エラー", message)
    
    def stop_csv_logging(self, show_dialog=True):
        """CSV保存終了

        Args:
            show_dialog: True=完了ダイアログを表示、False=ログのみ（自動停止時用）
        """
        if not self.is_csv_logging:
            if show_dialog:
                messagebox.showwarning("警告", "CSV保存が開始されていません")
            return

        # ログ記録停止してCSVに保存
        success, message = self.csv_logger.stop_logging()

        self.is_csv_logging = False
        self.csv_start_button.config(state=tk.NORMAL if self.is_measuring else tk.DISABLED)
        self.csv_stop_button.config(state=tk.DISABLED)
        self.csv_status_label.config(text="", foreground="gray")

        # ★★★ DEF選択のチェックボックスを有効化 ★★★
        self._lock_def_checkboxes(False)

        if success:
            self.log(message, "SUCCESS")
            if show_dialog:
                messagebox.showinfo("保存完了", message)
        else:
            self.log(message, "ERROR")
            if show_dialog:
                messagebox.showerror("エラー", message)
    
    def _lock_def_checkboxes(self, lock):
        """
        DEF選択チェックボックスのロック/アンロック
        
        Args:
            lock: True=ロック（無効化）、False=アンロック（有効化）
        """
        if not hasattr(self.test_tab, 'def_checkboxes'):
            # チェックボックスウィジェットが見つからない場合は何もしない
            return
        
        try:
            state = tk.DISABLED if lock else tk.NORMAL
            for checkbox in self.test_tab.def_checkboxes:
                checkbox.config(state=state)
            
            if lock:
                self.log("DEF選択をロックしました（保存中は変更不可）", "INFO")
            else:
                self.log("DEF選択のロックを解除しました", "INFO")
                
        except Exception as e:
            self.log(f"DEFチェックボックス操作エラー: {e}", "WARNING")
    
    def _load_save_config(self):
        """設定ファイルから保存先とシリアルNo.を取得"""
        config_file = "app_settings.json"
        save_dir = "measurement_data"
        filename = "measurement.csv"  # デフォルトは拡張子付き
        serial_numbers = {}
        
        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # 保存先ディレクトリ
                save_dir = config.get("save_config", {}).get("save_dir", save_dir)
                
                # CSVファイル名（通信1の設定）
                comm1_config = config.get("comm_profiles", {}).get("1", {})
                file_name_with_ext = comm1_config.get("save_config", {}).get("file_name", "measurement.csv")
                # 拡張子が付いていない場合は追加
                if not file_name_with_ext.lower().endswith('.csv'):
                    filename = f"{file_name_with_ext}.csv"
                else:
                    filename = file_name_with_ext
                
                # シリアルナンバー（通信1の設定、チェックされたDEFのみ）
                serial_config = comm1_config.get("serial_numbers", {})
                
                # チェックされたDEFのみを取得
                if hasattr(self.test_tab, 'def_check_vars'):
                    for i, check_var in enumerate(self.test_tab.def_check_vars):
                        if check_var.get():
                            sn_key = f"DEF{i}_sn"
                            serial = serial_config.get(sn_key, f"DEF{i}")
                            serial_numbers[i] = serial
        
        except Exception as e:
            self.log(f"設定ファイル読み込みエラー: {e}", "ERROR")

        # ★★★ 同名ファイルが存在する場合は別名にする ★★★
        filename = self._get_unique_filename(save_dir, filename)

        return save_dir, filename, serial_numbers

    def _get_unique_filename(self, save_dir, filename):
        """同名ファイルが存在する場合、連番を付けた別名を返す"""
        filepath = os.path.join(save_dir, filename)

        if not os.path.exists(filepath):
            return filename

        # ファイル名と拡張子を分離
        base_name, ext = os.path.splitext(filename)

        # 連番を付けて重複しないファイル名を探す
        counter = 1
        while True:
            new_filename = f"{base_name}_{counter}{ext}"
            new_filepath = os.path.join(save_dir, new_filename)
            if not os.path.exists(new_filepath):
                self.log(f"同名ファイルが存在するため別名で保存: {new_filename}", "WARNING")
                return new_filename
            counter += 1
            if counter > 9999:  # 無限ループ防止
                break

        return filename
    
    def on_closing(self):
        """終了処理"""
        # ★★★ CSV保存中なら停止 ★★★
        if self.is_csv_logging:
            self.stop_csv_logging(show_dialog=False)
        
        # ★★★ test_tab側のコールバックに処理を委譲 ★★★
        # test_tab.on_measurement_window_close()が呼ばれる
        pass
            
    def open_all_used_channels(self):
        """使用したチャンネルを全てOPEN

        【単一選択のクリーンアップ】
        計測停止時に呼ばれ、使用した全チャンネルをOPENする
        これにより次回計測開始時にクリーンな状態から開始できる
        """
        try:
            selected_defs = self.get_selected_defs()
            if not selected_defs:
                return

            orig_timeout = self.gpib_scanner.instrument.timeout
            self.gpib_scanner.instrument.timeout = 5000

            try:
                for def_entry in selected_defs:
                    pos_ch = def_entry['pos_channel']
                    neg_ch = def_entry['neg_channel']

                    if pos_ch != "ー" and pos_ch != "":
                        pos_num = pos_ch.replace("CH", "")
                        pos_addr = f"@{self.scanner_slot}{pos_num}"
                        self.gpib_scanner.write(f"OPEN ({pos_addr})")
                        time.sleep(0.05)

                    if neg_ch != "ー" and neg_ch != "":
                        neg_num = neg_ch.replace("CH", "")
                        neg_addr = f"@{self.scanner_slot}{neg_num}"
                        self.gpib_scanner.write(f"OPEN ({neg_addr})")
                        time.sleep(0.05)
            finally:
                self.gpib_scanner.instrument.timeout = orig_timeout
                
        except Exception as e:
            self.log(f"チャンネルOPEN時エラー: {e}", "ERROR")
            
    def get_current_pattern_info(self):
        """現在実行中のパターン情報を取得"""
        pattern_no = "---"
        dataset = "---"
        code = "---"
        pole = "---"
        
        # current_pattern_index が -1（未実行）の場合は「---」を返す
        if not hasattr(self.test_tab, 'current_pattern_index') or self.test_tab.current_pattern_index < 0:
            return {'pattern_no': pattern_no, 'dataset': dataset, 'code': code, 'pole': pole}
        
        # test_tabがパターンテスト実行中の場合
        if hasattr(self.test_tab, 'is_running') and self.test_tab.is_running:
            idx = self.test_tab.current_pattern_index
            if 0 <= idx < len(self.test_tab.patterns):
                p = self.test_tab.patterns[idx]
                if p['enabled'].get():
                    pattern_no = f"No.{idx + 1}"
                    dataset = p['dataset'].get()
                    code = p['code'].get()
                    pole = p['pole'].get()
                    
                    # ★★★ Neg選択時はコード値を反転して表示 ★★★
                    if pole == 'Neg':
                        # プリセット値の定義
                        presets = {
                            '+Full': {"P": "FFFFF", "L": "FFFF"},
                            'Center': {"P": "80000", "L": "8000"},
                            '-Full': {"P": "00000", "L": "0000"}
                        }
                        
                        dac_type = "P" if dataset == 'Position' else "L"
                        hex_value = None
                        
                        if code == 'Manual':
                            # Manualの場合は入力値を使用
                            manual_hex = p['manual_value'].get()
                            if manual_hex:
                                hex_value = manual_hex.upper()
                        elif code in presets:
                            # プリセット値を使用
                            hex_value = presets[code][dac_type]
                        
                        # HEX値を反転
                        if hex_value:
                            try:
                                hex_int = int(hex_value, 16)
                                if dataset == 'Position':
                                    hex_int = 0xFFFFF - hex_int  # 20ビット反転
                                    inverted_hex = f"{hex_int:05X}"
                                else:  # LBC
                                    hex_int = 0xFFFF - hex_int  # 16ビット反転
                                    inverted_hex = f"{hex_int:04X}"
                                
                                # codeを反転値で更新
                                if code == 'Manual':
                                    code = f"Manual ({inverted_hex})"
                                else:
                                    code = f"{code} ({inverted_hex})"
                            except ValueError:
                                # HEX変換エラーの場合は、もとのcodeを使用
                                if code == 'Manual':
                                    manual_hex = p['manual_value'].get()
                                    if manual_hex:
                                        code = f"Manual ({manual_hex})"
                    else:
                        # ★★★ Pos選択時も全ての場合で出力コードを表示 ★★★
                        # プリセット値の定義
                        presets = {
                            '+Full': {"P": "FFFFF", "L": "FFFF"},
                            'Center': {"P": "80000", "L": "8000"},
                            '-Full': {"P": "00000", "L": "0000"}
                        }
                        
                        dac_type = "P" if dataset == 'Position' else "L"
                        
                        if code == 'Manual':
                            manual_hex = p['manual_value'].get()
                            if manual_hex:
                                code = f"Manual ({manual_hex})"
                        elif code in presets:
                            # プリセット値の場合もHEX値を併記
                            hex_value = presets[code][dac_type]
                            code = f"{code} ({hex_value})"
        
        return {'pattern_no': pattern_no, 'dataset': dataset, 'code': code, 'pole': pole}
        
    def start_pattern_info_update(self):
        """パターン情報の定期更新を開始"""
        self.update_pattern_info_display()
    
    def update_pattern_info_display(self):
        """パターン情報表示を更新（定期実行）"""
        # パターン情報を取得して表示を更新
        pattern_info = self.get_current_pattern_info()
        self.pattern_no_label.config(text=pattern_info['pattern_no'])
        self.dataset_label.config(text=pattern_info['dataset'])
        self.code_label.config(text=pattern_info['code'])

        # ★★★ パターン実行同期オプションが有効な場合 ★★★
        if self.sync_with_pattern_var.get():
            current_pattern_running = (
                hasattr(self.test_tab, 'is_running') and self.test_tab.is_running
            )

            # パターン実行開始を検知
            if current_pattern_running and not self.last_pattern_running_state:
                # 計測していない場合は、まず計測を開始
                if not self.is_measuring:
                    self.log("パターン実行開始を検知 → 計測自動開始", "INFO")
                    self.start_measurement()
                    # start_measurementが成功したかチェック
                    if self.is_measuring and not self.is_csv_logging:
                        self.log("計測開始成功 → 保存自動開始", "INFO")
                        self.start_csv_logging()
                elif not self.is_csv_logging:
                    # 計測中だが保存していない場合は保存開始
                    self.log("パターン実行開始を検知 → 保存自動開始", "INFO")
                    self.start_csv_logging()

            # パターン実行終了を検知 → 保存終了、計測終了
            elif not current_pattern_running and self.last_pattern_running_state:
                if self.is_csv_logging:
                    self.log("パターン実行終了を検知 → 保存自動終了", "INFO")
                    self.stop_csv_logging(show_dialog=False)
                if self.is_measuring:
                    self.log("パターン実行終了を検知 → 計測自動終了", "INFO")
                    self.stop_measurement()

            self.last_pattern_running_state = current_pattern_running

        # 200ms後に再度実行（定期更新）
        self.update_timer_id = self.after(200, self.update_pattern_info_display)
        
    def get_dmm_settings(self):
        """DMMの設定を取得（dmm3458a_tabの機能を利用）"""
        try:
            # 親ウィンドウのdmm3458a_tabを取得
            if hasattr(self.parent_app, 'dmm3458a_tab'):
                dmm_tab = self.parent_app.dmm3458a_tab
                
                # dmm3458a_tabの設定取得メソッドを呼び出し
                settings = dmm_tab.get_current_settings()
                
                if settings['success']:
                    self.dmm_mode = settings['mode']
                    self.dmm_range = settings['range']
                    self.dmm_nplc = settings['nplc']
                    
                    # 表示を更新
                    self.dmm_mode_label.config(text=self.dmm_mode)
                    self.dmm_range_label.config(text=self.dmm_range)
                    self.dmm_nplc_label.config(text=self.dmm_nplc)

                    self.log(f"DMM設定: Mode={self.dmm_mode}, Range={self.dmm_range}, NPLC={self.dmm_nplc}", "INFO")
                    self._update_estimate()
                    return True
                else:
                    self.dmm_mode = settings['mode']
                    self.dmm_range = settings['range']
                    self.dmm_nplc = settings['nplc']
                    
                    self.dmm_mode_label.config(text=self.dmm_mode)
                    self.dmm_range_label.config(text=self.dmm_range)
                    self.dmm_nplc_label.config(text=self.dmm_nplc)
                    
                    self.log("DMM設定取得に失敗しました", "ERROR")
                    return False
            else:
                self.log("DMM3458Aタブにアクセスできません", "ERROR")
                self.dmm_mode = "Error"
                self.dmm_range = "Error"
                self.dmm_nplc = "Error"
                self.dmm_mode_label.config(text=self.dmm_mode)
                self.dmm_range_label.config(text=self.dmm_range)
                self.dmm_nplc_label.config(text=self.dmm_nplc)
                return False
                
        except Exception as e:
            self.log(f"DMM設定取得エラー: {e}", "ERROR")
            self.dmm_mode = "Error"
            self.dmm_range = "Error"
            self.dmm_nplc = "Error"
            self.dmm_mode_label.config(text=self.dmm_mode)
            self.dmm_range_label.config(text=self.dmm_range)
            self.dmm_nplc_label.config(text=self.dmm_nplc)
            return False

    def _load_switch_delay(self):
        """スキャナ切替時間を読み込み"""
        config_file = "app_settings.json"
        default_switch = 1.0

        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                switch_delay = config.get("measurement_window", {}).get("switch_delay_sec", default_switch)
                return max(0.5, switch_delay)

        except Exception:
            pass

        return default_switch

    def _save_switch_delay(self):
        """スキャナ切替時間を保存"""
        config_file = "app_settings.json"

        try:
            config = {}
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

            if "measurement_window" not in config:
                config["measurement_window"] = {}

            config["measurement_window"]["switch_delay_sec"] = self.switch_delay_sec.get()

            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)

        except Exception as e:
            self.log(f"設定保存エラー: {e}", "WARNING")

    def _on_delay_changed(self):
        """待ち時間設定変更時の処理"""
        self._save_switch_delay()
        self._update_estimate()

    def _update_estimate(self):
        """測定間隔目安を更新

        計算式: (スキャナー切替時間 + NPLC時間) × 選択ch数
        NPLC時間 = NPLC値 / 電源周波数(50Hz) ≒ NPLC × 0.02秒
        """
        try:
            switch_delay = self.switch_delay_sec.get()

            # NPLC値を取得（数値に変換）「100 PLC」形式から数値を抽出
            nplc_time = 0.0
            if self.dmm_nplc not in ("---", "Error", ""):
                try:
                    # 「100 PLC」形式から数値部分を抽出
                    nplc_str = self.dmm_nplc.replace("PLC", "").strip()
                    nplc_val = float(nplc_str)
                    nplc_time = nplc_val * 0.02  # 50Hz基準
                except ValueError:
                    pass

            # 1チャンネルあたりの時間
            per_ch_time = switch_delay + nplc_time

            # 選択ch数をカウント（Pos/Negで2倍）
            ch_count = 0
            if hasattr(self, 'test_tab') and self.test_tab and hasattr(self.test_tab, 'def_check_vars'):
                for i, var in enumerate(self.test_tab.def_check_vars):
                    if var.get():
                        pos_ch = self.test_tab.scanner_channels_pos[i].get() if i < len(self.test_tab.scanner_channels_pos) else "ー"
                        neg_ch = self.test_tab.scanner_channels_neg[i].get() if i < len(self.test_tab.scanner_channels_neg) else "ー"
                        if pos_ch != "ー" and pos_ch != "":
                            ch_count += 1
                        if neg_ch != "ー" and neg_ch != "":
                            ch_count += 1

            if ch_count > 0:
                total_time = per_ch_time * ch_count
                self.estimate_label.config(text=f"{total_time:.1f}秒/周 ({per_ch_time:.2f}秒/ch × {ch_count}ch)")
            else:
                self.estimate_label.config(text=f"{per_ch_time:.2f}秒/ch (DEF未選択)")

        except Exception:
            self.estimate_label.config(text="---")

    def _start_estimate_update(self):
        """測定間隔目安の定期更新を開始（DEF選択やチャンネル変更を検知）"""
        self._update_estimate()
        # 1秒ごとに更新
        self.after(1000, self._start_estimate_update)

    def _load_sync_option(self):
        """パターン実行同期オプションを読み込み"""
        config_file = "app_settings.json"

        try:
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

                return config.get("measurement_window", {}).get("sync_with_pattern", False)
        except Exception:
            pass

        return False

    def _save_sync_option(self):
        """パターン実行同期オプションを保存"""
        config_file = "app_settings.json"

        try:
            config = {}
            if os.path.exists(config_file):
                with open(config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)

            if "measurement_window" not in config:
                config["measurement_window"] = {}

            config["measurement_window"]["sync_with_pattern"] = self.sync_with_pattern_var.get()

            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)

        except Exception as e:
            self.log(f"設定保存エラー: {e}", "WARNING")

    def _on_sync_option_changed(self):
        """同期オプション変更時の処理"""
        self._save_sync_option()
        if self.sync_with_pattern_var.get():
            self.log("パターン実行と保存の同期を有効化", "INFO")
        else:
            self.log("パターン実行と保存の同期を無効化", "INFO")

    # ★★★ パターン切替前の待機機能 ★★★
    def _get_nplc_time(self):
        """NPLC時間を秒単位で取得"""
        nplc_time = 0.0
        if self.dmm_nplc not in ("---", "Error", ""):
            try:
                # 「100 PLC」形式から数値部分を抽出
                nplc_str = self.dmm_nplc.replace("PLC", "").strip()
                nplc_val = float(nplc_str)
                nplc_time = nplc_val * 0.02  # 50Hz基準
            except ValueError:
                pass
        return nplc_time

    def _get_measurement_interval_seconds(self):
        """測定間隔目安を秒単位で取得"""
        try:
            selected_defs = self.get_selected_defs()
            total_channels = 0
            for def_info in selected_defs:
                if def_info['pos_channel'] and def_info['pos_channel'] != "ー":
                    total_channels += 1
                if def_info['neg_channel'] and def_info['neg_channel'] != "ー":
                    total_channels += 1

            if total_channels == 0:
                return 0

            switch_delay = self.switch_delay_sec.get()
            nplc_time = self._get_nplc_time()

            # 1周回の時間 = (スキャナー切替時間 + NPLC時間) × チャンネル数
            interval = (switch_delay + nplc_time) * total_channels
            return interval
        except Exception:
            return 0

    def _should_wait_for_pattern_change(self):
        """パターン切替前に待機すべきかどうかを判定

        Returns:
            bool: 待機すべき場合はTrue
        """
        # パターン実行中でない場合は待機不要
        if not self.test_tab.is_running:
            self.log("待機判定: パターン実行中でない", "INFO")
            return False

        # パターンの残り時間を取得
        remaining = self.test_tab.get_pattern_remaining_seconds()
        if remaining is None:
            self.log("待機判定: 残り時間取得不可", "INFO")
            return False

        # 測定間隔を取得
        interval = self._get_measurement_interval_seconds()
        if interval <= 0:
            self.log(f"待機判定: 測定間隔が0以下 ({interval})", "INFO")
            return False

        # 残り時間が測定間隔より短い場合は待機
        self.log(f"待機判定: 残り{remaining:.1f}秒 vs 測定間隔{interval:.1f}秒", "INFO")
        return remaining < interval

    def _start_waiting_for_pattern_change(self, selected_defs):
        """パターン切替待機を開始"""
        self.is_waiting_for_pattern_change = True
        self.waiting_pattern_index = self.test_tab.get_current_pattern_index()
        self.waiting_selected_defs = selected_defs
        self.waiting_def_index = 0
        self.waiting_pole = "Pos"

        remaining = self.test_tab.get_pattern_remaining_seconds()
        self.log(f"パターン切替待機開始（残り{remaining:.1f}秒）", "INFO")
        self.csv_status_label.config(text="■ パターン切替待機中", foreground="orange")

        # 500msごとにパターン切替をチェック
        self.after(500, self._check_pattern_change)

    def _check_pattern_change(self):
        """パターンが切り替わったかをチェック"""
        if not self.is_waiting_for_pattern_change:
            return

        if not self.is_measuring:
            # 計測が停止された場合は待機終了
            self._end_waiting_for_pattern_change()
            return

        # パターン実行が終了した場合
        if not self.test_tab.is_running:
            self.log("パターン実行終了を検知、待機終了", "INFO")
            self._end_waiting_for_pattern_change()
            return

        # パターンインデックスが変わったかチェック
        current_index = self.test_tab.get_current_pattern_index()
        if current_index != self.waiting_pattern_index:
            # パターンが切り替わった
            self.log(f"パターン切替を検知（No.{self.waiting_pattern_index} → No.{current_index}）、計測再開", "INFO")
            self._end_waiting_for_pattern_change()

            # 計測を再開
            self.do_one_measurement(self.waiting_selected_defs, 0, "Pos")
            return

        # まだ切り替わっていない場合は再チェック
        self.after(500, self._check_pattern_change)

    def _end_waiting_for_pattern_change(self):
        """パターン切替待機を終了"""
        self.is_waiting_for_pattern_change = False
        self.waiting_pattern_index = -1

        if self.is_csv_logging:
            self.csv_status_label.config(text="■ CSV保存中", foreground="red")
        else:
            self.csv_status_label.config(text="")