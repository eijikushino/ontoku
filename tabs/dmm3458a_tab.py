import tkinter as tk
from tkinter import ttk
import time
import threading
import queue
from utils import LoggerWidget, validate_command

class DMM3458ATab(ttk.Frame):
    def __init__(self, parent, gpib_controller):
        super().__init__(parent)
        self.gpib = gpib_controller

        self.continuous_running = False
        self.continuous_error_count = 0

        # スレッド化用: 測定結果キューとスレッド管理
        self.measurement_queue = queue.Queue()
        self.measurement_thread = None
        self.measurement_lock = threading.Lock()

        self.create_widgets()
    
    def create_widgets(self):
        # ★★★ 設定フレーム（2行構成）★★★
        config_frame = ttk.LabelFrame(self, text="設定", padding=10)
        config_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 1行目：リセット＆初期化、モード、レンジ、NPLC、設定適用
        row1_frame = ttk.Frame(config_frame)
        row1_frame.pack(fill=tk.X, pady=2)
        
        ttk.Button(row1_frame, text="リセット＆初期化", 
                   command=self.reset_and_initialize, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row1_frame, text="モード:").pack(side=tk.LEFT, padx=(15, 2))
        self.mode_var = tk.StringVar(value="DCV")
        mode_combo = ttk.Combobox(
            row1_frame,
            textvariable=self.mode_var,
            values=["DCV", "ACV", "DCI", "ACI", "OHM", "OHMF"],
            width=8,
            state="readonly"
        )
        mode_combo.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row1_frame, text="レンジ:").pack(side=tk.LEFT, padx=(10, 2))
        self.range_var = tk.StringVar(value="10")
        range_combo = ttk.Combobox(
            row1_frame,
            textvariable=self.range_var,
            values=["0.1", "1", "10", "100", "1000"],
            width=8
        )
        range_combo.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row1_frame, text="NPLC:").pack(side=tk.LEFT, padx=(10, 2))
        self.nplc_var = tk.StringVar(value="5")
        nplc_combo = ttk.Combobox(
            row1_frame,
            textvariable=self.nplc_var,
            values=["0.0001", "0.001", "0.01", "0.1", "1", "10", "100"],
            width=8
        )
        nplc_combo.pack(side=tk.LEFT, padx=2)
        
        ttk.Button(row1_frame, text="設定を適用", 
                   command=self.apply_config, width=12).pack(side=tk.LEFT, padx=(10, 5))
        
        # 2行目：設定確認ボタンと現在の設定表示
        row2_frame = ttk.Frame(config_frame)
        row2_frame.pack(fill=tk.X, pady=2)
        
        ttk.Button(row2_frame, text="設定を確認", 
                   command=self.show_current_settings, width=15).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(row2_frame, text="測定モード:").pack(side=tk.LEFT, padx=(15, 2))
        self.current_mode_label = ttk.Label(row2_frame, text="---", 
                                             foreground="blue", font=("", 9, "bold"), width=12)  # ★8→12に拡大
        self.current_mode_label.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row2_frame, text="測定レンジ:").pack(side=tk.LEFT, padx=(10, 2))
        self.current_range_label = ttk.Label(row2_frame, text="---", 
                                              foreground="blue", font=("", 9, "bold"), width=10)  # ★8→10に拡大
        self.current_range_label.pack(side=tk.LEFT, padx=2)
        
        ttk.Label(row2_frame, text="積分時間:").pack(side=tk.LEFT, padx=(10, 2))
        self.current_nplc_label = ttk.Label(row2_frame, text="---", 
                                             foreground="blue", font=("", 9, "bold"), width=12)  # ★8→12に拡大
        self.current_nplc_label.pack(side=tk.LEFT, padx=2)
        
        # 測定実行フレーム
        measure_frame = ttk.LabelFrame(self, text="測定実行", padding=10)
        measure_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(measure_frame, text="単発測定", 
                   command=self.single_measurement).pack(side=tk.LEFT, padx=5)
        ttk.Button(measure_frame, text="連続測定開始", 
                   command=self.start_continuous).pack(side=tk.LEFT, padx=5)
        ttk.Button(measure_frame, text="停止", 
                   command=self.stop_continuous).pack(side=tk.LEFT, padx=5)
        
        # 測定値表示フレーム
        result_frame = ttk.LabelFrame(self, text="測定結果", padding=10)
        result_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 生値表示
        ttk.Label(result_frame, text="生値:", font=("", 10)).grid(row=0, column=0, padx=5, pady=5, sticky=tk.E)
        self.result_label = ttk.Label(result_frame, text="-- ", 
                                       font=("", 14, "bold"), 
                                       foreground="blue")
        self.result_label.grid(row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # パネル形式表示
        ttk.Label(result_frame, text="パネル形式:", font=("", 10)).grid(row=1, column=0, padx=5, pady=5, sticky=tk.E)
        self.panel_label = ttk.Label(result_frame, text="-- ", 
                                      font=("", 16, "bold"), 
                                      foreground="green")
        self.panel_label.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        
        # デバッグフレーム
        debug_frame = ttk.LabelFrame(self, text="デバッグ", padding=10)
        debug_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(debug_frame, text="エラー確認", 
                   command=self.check_error).pack(side=tk.LEFT, padx=5)
        
        ttk.Label(debug_frame, text="コマンド:").pack(side=tk.LEFT, padx=5)
        self.custom_command_entry = ttk.Entry(debug_frame, width=20)
        self.custom_command_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Button(debug_frame, text="Write", 
                   command=self.send_write).pack(side=tk.LEFT, padx=5)
        ttk.Button(debug_frame, text="Write→Read", 
                   command=self.send_write_read).pack(side=tk.LEFT, padx=5)
        
        # ログ表示
        log_frame = ttk.LabelFrame(self, text="ログ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.logger = LoggerWidget(log_frame, height=12)
        
        # 注意事項を表示
        self.logger.log("【注意】3458Aは古い機器のため、SCPI標準コマンド(*IDN?等)は非対応です", "INFO")
        self.logger.log("使用可能: DCV, ACV, DCI, ACI, OHM, OHMF, NPLC, RANGE, TRIG SGL等", "INFO")
    
    def calculate_measurement_interval(self):
        """NPLCから測定間隔を計算（ミリ秒）"""
        try:
            nplc = float(self.nplc_var.get())
            # 50Hz電源の場合: 1 PLC = 20ms
            # 余裕を持たせて1.1倍
            interval_ms = int(nplc * 20 * 1.1)
            # 最小10ms
            return max(10, interval_ms)
        except ValueError:
            return 100
    
    def reset_and_initialize(self):
        """機器をリセットして初期化"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        self.logger.log("リセット＆初期化開始...", "INFO")
        
        # RESET
        success, _ = self.gpib.write("RESET")
        if not success:
            self.logger.log("RESET失敗", "ERROR")
            return
        
        self.logger.log("RESET完了、初期化中...", "INFO")
        
        # 少し待つ（RESETの処理時間）
        self.after(100, self.initialize_after_reset)
        
    def apply_config(self):
        """測定設定を適用"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        mode = self.mode_var.get()
        range_val = self.range_var.get()
        nplc = self.nplc_var.get()
        
        # ★★★ モード名を正しい番号に変換 ★★★
        mode_to_number = {
            "DCV": "1",
            "ACV": "2",
            "OHM": "4",      # ← 修正
            "OHMF": "5",     # ← 修正
            "DCI": "6",      # ← 修正
            "ACI": "7"       # ← 修正
        }
        
        mode_number = mode_to_number.get(mode, mode)
        
        self.logger.log(f"設定を適用中: Mode={mode}(番号{mode_number}), Range={range_val}, NPLC={nplc}", "INFO")
        
        try:
            # GPIBバッファクリア
            try:
                self.gpib.instrument.clear()
            except:
                pass
            
            # モード設定
            success, _ = self.gpib.write(f"FUNC {mode}")  # 文字列で送る
            if not success:
                self.logger.log(f"モード設定失敗: {mode}", "ERROR")
                return
            self.logger.log(f"モード設定完了: {mode}", "SUCCESS")
            
            # レンジ設定
            success, _ = self.gpib.write(f"RANGE {range_val}")
            if not success:
                self.logger.log(f"レンジ設定失敗: {range_val}", "ERROR")
                return
            self.logger.log(f"レンジ設定完了: {range_val}", "SUCCESS")
            
            # NPLC設定
            success, _ = self.gpib.write(f"NPLC {nplc}")
            if not success:
                self.logger.log(f"NPLC設定失敗: {nplc}", "ERROR")
                return
            self.logger.log(f"NPLC設定完了: {nplc}", "SUCCESS")
            
            self.logger.log("すべての設定が適用されました", "SUCCESS")
            
            # 1秒待ってから確認
            self.after(1000, self.show_current_settings)
            
        except Exception as e:
            self.logger.log(f"設定適用エラー: {e}", "ERROR")
    
    def initialize_after_reset(self):
        """RESET後の初期化処理"""
        # PRESET NORM
        success, _ = self.gpib.write("PRESET NORM")
        if not success:
            self.logger.log("PRESET失敗", "ERROR")
            return
        
        # END ALWAYS (必須: GPIB EOIラインを正しく動作させる)
        success, _ = self.gpib.write("END ALWAYS")
        if not success:
            self.logger.log("END ALWAYS失敗", "ERROR")
            return
        
        # TARM AUTO (トリガーアームを自動に設定)
        success, _ = self.gpib.write("TARM AUTO")
        if not success:
            self.logger.log("TARM AUTO失敗", "ERROR")
            return
        
        # DCVモード設定
        success, _ = self.gpib.write("DCV")
        if not success:
            self.logger.log("DCVモード設定失敗", "ERROR")
            return
        
        self.logger.log("リセット＆初期化完了", "SUCCESS")
    
    def apply_settings(self):
        """測定設定を適用(VBAのDmmRangeSet相当)"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        mode = self.mode_var.get()
        range_val = self.range_var.get()
        nplc = self.nplc_var.get()
        
        # モード+レンジ設定(例: "DCV 10")
        command = f"{mode} {range_val}"
        success, _ = self.gpib.write(command)
        if not success:
            self.logger.log(f"モード・レンジ設定失敗: {command}", "ERROR")
            return
        
        # NPLC設定
        command = f"NPLC {nplc}"
        success, _ = self.gpib.write(command)
        if not success:
            self.logger.log(f"NPLC設定失敗: {command}", "ERROR")
            return
        
        # 測定間隔を更新
        interval = self.calculate_measurement_interval()
        
        self.logger.log(f"設定完了: {mode} {range_val}, NPLC {nplc}", "SUCCESS")
    
    def format_panel_value(self, value_str):
        """パネル表示形式にフォーマット（符号付き8桁）"""
        try:
            value = float(value_str)
            range_val = float(self.range_var.get())
            
            # レンジに応じて整数部と小数部の桁数を決定（合計8桁）
            if range_val >= 1000:
                int_digits = 4  # 1000の位まで（0123）
                dec_digits = 4
            elif range_val >= 100:
                int_digits = 3  # 100の位まで（012）
                dec_digits = 5
            elif range_val >= 10:
                int_digits = 2  # 10の位まで（01）
                dec_digits = 6
            elif range_val >= 1:
                int_digits = 1  # 1の位まで（1）
                dec_digits = 7
            else:  # 0.1V
                int_digits = 1  # 0.1の位まで（0）
                dec_digits = 8
            
            # 全体幅 = 符号(1) + 整数部 + 小数点(1) + 小数部
            total_width = 1 + int_digits + 1 + dec_digits
            
            # フォーマット：符号 + 全体幅で0埋め + 小数桁数指定
            formatted = f"{value:+0{total_width}.{dec_digits}f}"
            
            # 幅を超えた場合は切り詰め（オーバーレンジ表示）
            if len(formatted) > total_width:
                formatted = "+" + "9" * int_digits + "." + "9" * dec_digits
            
            return formatted
        except:
            return value_str.strip()
    
    def single_measurement(self):
        """単発測定(VBAのDmmReadAny相当) - 非同期版"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return

        # 測定中表示
        self.result_label.config(text="測定中...")
        self.panel_label.config(text="--")

        # 別スレッドで測定を実行
        thread = threading.Thread(
            target=self._single_measure_worker,
            daemon=True
        )
        thread.start()

        # 結果のポーリング開始
        self.after(10, self._check_single_result)

    def _single_measure_worker(self):
        """単発測定のワーカースレッド"""
        result = {
            'success': False,
            'response': None,
            'error': None
        }

        try:
            with self.measurement_lock:
                success, _ = self.gpib.write("TRIG SGL")
                if success:
                    success, response = self.gpib.read()
                    if success:
                        result['success'] = True
                        result['response'] = response
                    else:
                        result['error'] = 'read_failed'
                else:
                    result['error'] = 'write_failed'
        except Exception as e:
            result['error'] = str(e)

        self.measurement_queue.put(result)

    def _check_single_result(self):
        """単発測定の結果をチェック"""
        try:
            result = self.measurement_queue.get_nowait()

            if result['success']:
                try:
                    raw_value = result['response'].strip()
                    unit = self.get_unit(self.mode_var.get())
                    self.result_label.config(text=f"{raw_value} {unit}")

                    panel_value = self.format_panel_value(result['response'])
                    self.panel_label.config(text=f"{panel_value} {unit}")

                    self.logger.log(f"測定成功: {raw_value} {unit}", "SUCCESS")
                except ValueError:
                    self.result_label.config(text="変換エラー")
                    self.panel_label.config(text="--")
                    self.logger.log(f"データ変換エラー: {result['response']}", "ERROR")
                    self.check_error()
            else:
                self.result_label.config(text="測定失敗")
                self.panel_label.config(text="--")
                self.logger.log(f"測定失敗: {result.get('error', 'unknown')}", "ERROR")
                self.check_error()

        except queue.Empty:
            # まだ結果がない → 再チェック
            self.after(20, self._check_single_result)
    
    def start_continuous(self):
        """連続測定を開始"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        if self.continuous_running:
            self.logger.log("既に連続測定中です", "ERROR")
            return
        
        # エラーカウントをリセット
        self.continuous_error_count = 0
        
        # 測定間隔を計算して表示
        interval = self.calculate_measurement_interval()
        
        self.continuous_running = True
        self.logger.log(f"連続測定開始（間隔: {interval}ms）", "INFO")
        
        # 最初の測定をスケジュール
        self.after(10, self.continuous_measure)
    
    def stop_continuous(self):
        """連続測定を停止"""
        self.continuous_running = False

        # キューをクリア（残っている結果を破棄）
        while not self.measurement_queue.empty():
            try:
                self.measurement_queue.get_nowait()
            except queue.Empty:
                break

        self.logger.log(f"連続測定停止（総エラー数: {self.continuous_error_count}）", "INFO")
    
    def continuous_measure(self):
        """連続測定のループ（非同期版）

        メインスレッドから呼ばれ、別スレッドで測定を実行。
        UIをブロックしないため、NPLC大でもラグなし。
        """
        if not self.continuous_running:
            return

        # 前回のスレッドがまだ実行中なら結果待ちへスキップ
        if self.measurement_thread and self.measurement_thread.is_alive():
            self.after(50, self._check_measurement_result)
            return

        # 別スレッドで測定を実行
        self.measurement_thread = threading.Thread(
            target=self._measure_thread_worker,
            daemon=True
        )
        self.measurement_thread.start()

        # 結果のポーリング開始（短い間隔でUIを更新可能に保つ）
        self.after(10, self._check_measurement_result)

    def _measure_thread_worker(self):
        """別スレッドで実行されるGPIB測定ワーカー

        GPIB通信（TRIG SGL → read）を実行し、結果をキューに入れる。
        このメソッドはUIをブロックしない。
        """
        result = {
            'success': False,
            'response': None,
            'error': None
        }

        try:
            with self.measurement_lock:
                # TRIG SGL でトリガー
                success, _ = self.gpib.write("TRIG SGL")
                if success:
                    # 測定値読み取り（ここでNPLC分待機するが別スレッドなのでOK）
                    success, response = self.gpib.read()
                    if success:
                        result['success'] = True
                        result['response'] = response
                    else:
                        result['error'] = 'read_failed'
                else:
                    result['error'] = 'write_failed'
        except Exception as e:
            result['error'] = str(e)

        # 結果をキューに入れる（メインスレッドで処理）
        self.measurement_queue.put(result)

    def _check_measurement_result(self):
        """メインスレッドで測定結果をチェック・UI更新

        キューから結果を取得し、UIを更新。
        結果がなければ再度ポーリングをスケジュール。
        """
        if not self.continuous_running:
            return

        try:
            # キューから結果を取得（ノンブロッキング）
            result = self.measurement_queue.get_nowait()

            if result['success']:
                try:
                    # 生値表示
                    raw_value = result['response'].strip()
                    unit = self.get_unit(self.mode_var.get())
                    self.result_label.config(text=f"{raw_value} {unit}")

                    # パネル形式表示
                    panel_value = self.format_panel_value(result['response'])
                    self.panel_label.config(text=f"{panel_value} {unit}")
                except ValueError:
                    self.continuous_error_count += 1
            else:
                self.continuous_error_count += 1
                # 10回に1回だけログ出力
                if self.continuous_error_count % 10 == 0:
                    error_msg = result.get('error', 'unknown')
                    self.logger.log(f"読み取りエラー継続中（累計: {self.continuous_error_count}回, {error_msg}）", "ERROR")

            # 次の測定をスケジュール
            interval = self.calculate_measurement_interval()
            self.after(interval, self.continuous_measure)

        except queue.Empty:
            # まだ結果がない → 短い間隔で再チェック（UIは応答可能）
            self.after(20, self._check_measurement_result)
    
    def check_error(self):
        """エラー内容を確認"""
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return
        
        self.logger.log("エラー確認中...", "INFO")
        
        # ERRSTR?でエラー詳細を取得
        success, _ = self.gpib.write("ERRSTR?")
        if not success:
            self.logger.log("ERRSTRコマンド送信失敗", "ERROR")
            return
        
        # 応答読み取り
        success, response = self.gpib.read()
        if success:
            err_msg = response.strip()
            # エラーコードが0または"NO ERROR"の場合
            if err_msg.startswith("0") or "NO ERROR" in err_msg.upper():
                self.logger.log(f"エラーなし: {err_msg}", "SUCCESS")
            else:
                self.logger.log(f"エラー内容: {err_msg}", "ERROR")
        else:
            self.logger.log("エラー応答の読み取り失敗", "ERROR")
    
    def send_write(self):
        """カスタムコマンドを送信(Write)"""
        valid, command = validate_command(self.custom_command_entry.get())
        if not valid:
            self.logger.log("コマンドが空です", "ERROR")
            return
        
        success, message = self.gpib.write(command)
        level = "SUCCESS" if success else "ERROR"
        self.logger.log(f"送信: {command}", level)
    
    def send_write_read(self):
        """カスタムコマンドを送信して読み取り"""
        valid, command = validate_command(self.custom_command_entry.get())
        if not valid:
            self.logger.log("コマンドが空です", "ERROR")
            return
        
        self.logger.log(f"送信: {command}", "INFO")
        
        # 元のタイムアウトを保存
        original_timeout = self.gpib.instrument.timeout
        
        # 短いタイムアウト(2秒)を設定
        self.gpib.instrument.timeout = 2000
        
        # query()を使用
        success, response = self.gpib.query(command)
        
        # タイムアウトを元に戻す
        self.gpib.instrument.timeout = original_timeout
        
        if success:
            self.logger.log(f"応答: {response}", "SUCCESS")
        else:
            self.logger.log("応答なし（タイムアウトまたはエラー）", "WARNING")
    
    def get_unit(self, mode):
        """測定モードに応じた単位を返す"""
        unit_map = {
            "DCV": "V",
            "ACV": "V",
            "DCI": "A",
            "ACI": "A",
            "OHM": "Ω",
            "OHMF": "Ω",
        }
        return unit_map.get(mode, "")
        
    def get_current_settings(self):
        """現在のDMM設定を取得
        
        Returns:
            dict: {'mode': str, 'range': str, 'nplc': str, 'success': bool}
                  成功時はsuccess=True、失敗時はsuccess=False
        """
        result = {
            'mode': '---',
            'range': '---',
            'nplc': '---',
            'success': False
        }
        
        if not self.gpib.connected:
            self.logger.log("機器が接続されていません", "ERROR")
            return result
        
        try:
            original_timeout = self.gpib.instrument.timeout
            self.gpib.instrument.timeout = 2000
            
            # ★★★ モード取得（FUNC?）★★★
            success, mode_response = self.gpib.query("FUNC?")
            mode_code = ""
            if success:
                mode_str = mode_response.strip().strip('"').strip()
                # カンマ区切りの場合は最初の値を取得
                if ',' in mode_str:
                    mode_str = mode_str.split(',')[0].strip()
                
                # 3458Aは番号で返す: 1=DCV, 2=ACV, 3=OHM, 4=OHMF, など
                mode_map = {
                    "1": ("DCV", "DC電圧"),
                    "2": ("ACV", "AC電圧"),
                    "3": ("ACDCV", "AC+DC電圧"),    # ← 修正
                    "4": ("OHM", "抵抗(2線)"),      # ← 修正
                    "5": ("OHMF", "抵抗(4線)"),     # ← 修正
                    "6": ("DCI", "DC電流"),         # ← 修正
                    "7": ("ACI", "AC電流"),         # ← 修正
                    "8": ("ACDCI", "AC+DC電流"),    # ← 修正
                    "9": ("FREQ", "周波数"),        # ← 修正
                    "10": ("PER", "周期"),          # ← 修正
                    # 文字列の場合も対応
                    "DCV": ("DCV", "DC電圧"),
                    "ACV": ("ACV", "AC電圧"),
                    "OHM": ("OHM", "抵抗(2線)"),
                    "OHMF": ("OHMF", "抵抗(4線)"),
                    "DCI": ("DCI", "DC電流"),
                    "ACI": ("ACI", "AC電流"),
                    "FREQ": ("FREQ", "周波数"),
                    "PER": ("PER", "周期"),
                    "ACDCV": ("ACDCV", "AC+DC電圧"),
                    "ACDCI": ("ACDCI", "AC+DC電流")
                }
                
                if mode_str in mode_map:
                    mode_code, mode_name = mode_map[mode_str]
                    result['mode'] = mode_name
                else:
                    mode_code = mode_str
                    result['mode'] = f"不明({mode_str})"
            else:
                result['mode'] = "Error"
                self.logger.log("モード取得失敗", "ERROR")
            
            # ★★★ レンジ取得（RANGE?）★★★
            success, range_response = self.gpib.query("RANGE?")
            if success:
                range_value = range_response.strip()
                
                try:
                    range_float = float(range_value)
                    
                    # モードに応じて単位を決定
                    if mode_code in ['DCV', 'ACV', 'ACDCV']:
                        # 電圧モード
                        if range_float >= 1000:
                            val = range_float / 1000
                            result['range'] = f"{val:g} kV"
                        elif range_float >= 1:
                            result['range'] = f"{range_float:g} V"
                        else:
                            val = range_float * 1000
                            result['range'] = f"{val:g} mV"
                    elif mode_code in ['DCI', 'ACI', 'ACDCI']:
                        # 電流モード
                        if range_float >= 1:
                            result['range'] = f"{range_float:g} A"
                        elif range_float >= 0.001:
                            val = range_float * 1000
                            result['range'] = f"{val:g} mA"
                        else:
                            val = range_float * 1000000
                            result['range'] = f"{val:g} uA"
                    elif mode_code in ['OHM', 'OHMF']:
                        # 抵抗モード
                        if range_float >= 1000000:
                            val = range_float / 1000000
                            result['range'] = f"{val:g} MΩ"
                        elif range_float >= 1000:
                            val = range_float / 1000
                            result['range'] = f"{val:g} kΩ"
                        else:
                            result['range'] = f"{range_float:g} Ω"
                    else:
                        # その他のモード（単位なし）
                        result['range'] = f"{range_float:g}"
                except Exception as e:
                    result['range'] = range_value
            else:
                result['range'] = "Error"
                self.logger.log("レンジ取得失敗", "ERROR")
            
            # ★★★ NPLC取得（NPLC?）★★★
            success, nplc_response = self.gpib.query("NPLC?")
            if success:
                nplc_value = nplc_response.strip()
                try:
                    nplc_float = float(nplc_value)
                    result['nplc'] = f"{nplc_float:g} PLC"
                except:
                    result['nplc'] = nplc_value
            else:
                result['nplc'] = "Error"
                self.logger.log("NPLC取得失敗", "ERROR")
            
            # タイムアウトを元に戻す
            self.gpib.instrument.timeout = original_timeout
            
            # すべて成功した場合のみsuccess=True
            if result['mode'] != "Error" and result['range'] != "Error" and result['nplc'] != "Error":
                result['success'] = True
                self.logger.log(f"設定取得成功: Mode={result['mode']}, Range={result['range']}, NPLC={result['nplc']}", "SUCCESS")
            else:
                self.logger.log("設定取得に一部失敗しました", "WARNING")
            
            return result
            
        except Exception as e:
            self.logger.log(f"設定取得エラー: {e}", "ERROR")
            result['mode'] = "Error"
            result['range'] = "Error"
            result['nplc'] = "Error"
            result['success'] = False
            return result
            
    def show_current_settings(self):
        """現在の設定を確認して表示"""
        self.logger.log("現在のDMM設定を確認中...", "INFO")
        
        # get_current_settings()を呼び出し
        settings = self.get_current_settings()
        
        # ラベルを更新
        self.current_mode_label.config(text=settings['mode'])
        self.current_range_label.config(text=settings['range'])
        self.current_nplc_label.config(text=settings['nplc'])
        
        if settings['success']:
            self.logger.log("設定確認完了", "SUCCESS")
        else:
            self.logger.log("設定確認に失敗しました", "WARNING")