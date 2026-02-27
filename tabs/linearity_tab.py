import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
import queue
import time
import os
import json
import random
import numpy as np
import matplotlib.pyplot as plt


class LinearityTab(ttk.Frame):
    """Linearity試験タブ - DACの直線性を測定し、最小二乗法で誤差を算出する"""

    # DAC仕様定数
    DAC_SPECS = {
        'Position': {'bits': 20, 'span': 20.0, 'ci': 'ci', 'center': '80000', 'dmm_range': '10'},
        'LBC':      {'bits': 16, 'span': 2.0,  'ci': 'cii', 'center': '08000', 'dmm_range': '1'},
    }

    def __init__(self, parent, gpib_3458a, gpib_3499b, datagen_manager, test_tab):
        super().__init__(parent)
        self.gpib_dmm = gpib_3458a
        self.gpib_scanner = gpib_3499b
        self.datagen = datagen_manager
        self.test_tab = test_tab
        self.scanner_slot = "1"

        self.is_running = False
        self._stop_event = threading.Event()
        self._worker_thread = None
        self._update_queue = queue.Queue()

        # 設定変数
        self.pattern_mode = tk.StringVar(value='Linear')
        self.num_points = tk.StringVar(value='64')
        self.pattern_file = tk.StringVar(value='')
        self.position_var = tk.BooleanVar(value=True)
        self.lbc_var = tk.BooleanVar(value=False)
        self.settle_time_var = tk.DoubleVar(value=0.2)
        self.th_gain = tk.DoubleVar(value=0.01)
        self.th_offset = tk.DoubleVar(value=2.0)
        self.th_error = tk.DoubleVar(value=1.5)
        self.save_dir = tk.StringVar(value='linearity_data')

        self._load_settings()
        self._create_widgets()

    # ==================== UI ====================
    def _create_widgets(self):
        main = ttk.Frame(self)
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左パネル (設定)
        left = ttk.Frame(main, width=340)
        left.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left.pack_propagate(False)
        self._create_left_panel(left)

        # 右パネル (結果)
        right = ttk.Frame(main)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        self._create_right_panel(right)

    def _create_left_panel(self, parent):
        # 色バー付き見出し
        header = tk.Frame(parent, bg='#27ae60', height=28)
        header.pack(fill=tk.X, pady=(0, 8))
        header.pack_propagate(False)
        tk.Label(header, text="Linearity試験", bg='#27ae60', fg='white',
                 font=('', 11, 'bold')).pack(expand=True)

        # --- パターン設定 ---
        pat_frame = ttk.LabelFrame(parent, text="パターン設定", padding=8)
        pat_frame.pack(fill=tk.X, pady=(0, 5))

        mode_frame = ttk.Frame(pat_frame)
        mode_frame.pack(fill=tk.X, pady=2)
        ttk.Label(mode_frame, text="モード:").pack(side=tk.LEFT)
        for mode in ['Linear', 'Random', 'File']:
            ttk.Radiobutton(mode_frame, text=mode, variable=self.pattern_mode,
                            value=mode, command=self._on_mode_changed).pack(side=tk.LEFT, padx=3)

        pts_frame = ttk.Frame(pat_frame)
        pts_frame.pack(fill=tk.X, pady=2)
        ttk.Label(pts_frame, text="計測点数:").pack(side=tk.LEFT)
        self.pts_combo = ttk.Combobox(pts_frame, textvariable=self.num_points,
                                       values=['32', '64', '128', '256', '512', '1024'],
                                       width=8, state='readonly')
        self.pts_combo.pack(side=tk.LEFT, padx=5)

        file_frame = ttk.Frame(pat_frame)
        file_frame.pack(fill=tk.X, pady=2)
        ttk.Label(file_frame, text="ファイル:").pack(side=tk.LEFT)
        self.file_entry = ttk.Entry(file_frame, textvariable=self.pattern_file, width=16)
        self.file_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        self.file_browse_btn = ttk.Button(file_frame, text="参照",
                                           command=self._browse_pattern_file, width=5)
        self.file_browse_btn.pack(side=tk.LEFT)
        self._on_mode_changed()

        # --- DAC設定 ---
        dac_frame = ttk.LabelFrame(parent, text="DAC設定", padding=8)
        dac_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Checkbutton(dac_frame, text="Position (20bit)", variable=self.position_var).pack(anchor=tk.W)
        ttk.Checkbutton(dac_frame, text="LBC (16bit)", variable=self.lbc_var).pack(anchor=tk.W)

        settle_frame = ttk.Frame(dac_frame)
        settle_frame.pack(fill=tk.X, pady=2)
        ttk.Label(settle_frame, text="DAC安定待ち:").pack(side=tk.LEFT)
        ttk.Spinbox(settle_frame, from_=0.05, to=5.0, increment=0.05,
                     textvariable=self.settle_time_var, width=6, format="%.2f").pack(side=tk.LEFT, padx=2)
        ttk.Label(settle_frame, text="sec").pack(side=tk.LEFT)

        # --- NG判定閾値 ---
        th_frame = ttk.LabelFrame(parent, text="NG判定閾値", padding=8)
        th_frame.pack(fill=tk.X, pady=(0, 5))
        for label, var, unit in [("Gain閾値:", self.th_gain, "LSB/LSB"),
                                  ("Offset閾値:", self.th_offset, "LSB"),
                                  ("Error閾値:", self.th_error, "LSB")]:
            row = ttk.Frame(th_frame)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=label, width=10).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=var, width=8).pack(side=tk.LEFT, padx=2)
            ttk.Label(row, text=unit).pack(side=tk.LEFT)

        # --- 実行制御 ---
        ctrl_frame = ttk.LabelFrame(parent, text="実行制御", padding=8)
        ctrl_frame.pack(fill=tk.X, pady=(0, 5))

        btn_row = ttk.Frame(ctrl_frame)
        btn_row.pack(fill=tk.X, pady=2)
        self.start_btn = ttk.Button(btn_row, text="開始", command=self.start_measurement, width=12)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn = ttk.Button(btn_row, text="停止", command=self.stop_measurement,
                                    state=tk.DISABLED, width=12)
        self.stop_btn.pack(side=tk.LEFT, padx=2)

        dir_frame = ttk.Frame(ctrl_frame)
        dir_frame.pack(fill=tk.X, pady=2)
        ttk.Label(dir_frame, text="保存先:").pack(side=tk.LEFT)
        ttk.Entry(dir_frame, textvariable=self.save_dir, width=18).pack(
            side=tk.LEFT, padx=2, fill=tk.X, expand=True)
        ttk.Button(dir_frame, text="参照", command=self._browse_save_dir, width=5).pack(side=tk.LEFT)

    def _create_right_panel(self, parent):
        # --- 進捗 ---
        prog_frame = ttk.LabelFrame(parent, text="進捗", padding=8)
        prog_frame.pack(fill=tk.X, pady=(0, 5))

        self.target_label = ttk.Label(prog_frame, text="待機中", font=('', 10, 'bold'))
        self.target_label.pack(anchor=tk.W)

        prog_row = ttk.Frame(prog_frame)
        prog_row.pack(fill=tk.X, pady=2)
        self.progress_label = ttk.Label(prog_row, text="0 / 0")
        self.progress_label.pack(side=tk.LEFT)
        self.progressbar = ttk.Progressbar(prog_row, length=200, mode='determinate')
        self.progressbar.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # --- 結果サマリー ---
        summary_frame = ttk.LabelFrame(parent, text="結果サマリー", padding=5)
        summary_frame.pack(fill=tk.X, pady=(0, 5))

        cols = ('def', 'dac', 'pole', 'gain', 'offset', 'maxerr', 'judge')
        self.summary_tree = ttk.Treeview(summary_frame, columns=cols, show='headings', height=8)
        for col, text, w in [('def', 'DEF', 50), ('dac', 'DAC', 70), ('pole', 'Pole', 45),
                              ('gain', 'Gain', 85), ('offset', 'Offset', 70),
                              ('maxerr', 'MaxErr', 70), ('judge', '判定', 45)]:
            self.summary_tree.heading(col, text=text)
            self.summary_tree.column(col, width=w, anchor=tk.CENTER)
        self.summary_tree.pack(fill=tk.X)
        self.summary_tree.tag_configure('ng', background='#f5b7b1')
        self.summary_tree.tag_configure('ok', background='#d5f5e3')

        # --- 実行ログ ---
        log_frame = ttk.LabelFrame(parent, text="実行ログ", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = ScrolledText(log_frame, height=12, width=50,
                                      wrap=tk.WORD, state=tk.DISABLED, font=('Courier', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config('INFO', foreground='black')
        self.log_text.tag_config('SUCCESS', foreground='green')
        self.log_text.tag_config('WARNING', foreground='orange')
        self.log_text.tag_config('ERROR', foreground='red')

    # ==================== 設定 ====================
    def _load_settings(self):
        try:
            if os.path.exists('app_settings.json'):
                with open('app_settings.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                lin = config.get('linearity', {})
                self.pattern_mode.set(lin.get('pattern_mode', 'Linear'))
                self.num_points.set(lin.get('num_points', '64'))
                self.position_var.set(lin.get('position', True))
                self.lbc_var.set(lin.get('lbc', False))
                self.settle_time_var.set(lin.get('settle_time', 0.2))
                self.th_gain.set(lin.get('th_gain', 0.01))
                self.th_offset.set(lin.get('th_offset', 2.0))
                self.th_error.set(lin.get('th_error', 1.5))
                self.save_dir.set(lin.get('save_dir', 'linearity_data'))
        except Exception:
            pass

    def _save_settings(self):
        try:
            config = {}
            if os.path.exists('app_settings.json'):
                with open('app_settings.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
            config['linearity'] = {
                'pattern_mode': self.pattern_mode.get(),
                'num_points': self.num_points.get(),
                'position': self.position_var.get(),
                'lbc': self.lbc_var.get(),
                'settle_time': self.settle_time_var.get(),
                'th_gain': self.th_gain.get(),
                'th_offset': self.th_offset.get(),
                'th_error': self.th_error.get(),
                'save_dir': self.save_dir.get(),
            }
            with open('app_settings.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    # ==================== UIコールバック ====================
    def _on_mode_changed(self):
        is_file = self.pattern_mode.get() == 'File'
        state = tk.NORMAL if is_file else tk.DISABLED
        self.file_entry.config(state=state)
        self.file_browse_btn.config(state=state)
        self.pts_combo.config(state=tk.DISABLED if is_file else 'readonly')

    def _browse_pattern_file(self):
        path = filedialog.askopenfilename(
            title="パターンファイルを選択",
            filetypes=[("テキストファイル", "*.txt *.csv"), ("すべて", "*.*")])
        if path:
            self.pattern_file.set(path)

    def _browse_save_dir(self):
        d = filedialog.askdirectory(title="保存先フォルダを選択")
        if d:
            self.save_dir.set(d)

    # ==================== パターン生成 ====================
    def _generate_pattern(self, bits):
        mode = self.pattern_mode.get()
        max_val = (1 << bits) - 1

        if mode == 'File':
            values = []
            try:
                with open(self.pattern_file.get(), 'r') as f:
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith('#'):
                            values.append(int(line))
            except Exception as e:
                raise ValueError(f"パターンファイル読込エラー: {e}")
            return values

        n = int(self.num_points.get())
        if mode == 'Linear':
            if n <= 1:
                return [0]
            step = max_val / (n - 1)
            return [min(int(round(i * step)), max_val) for i in range(n)]
        else:  # Random
            return sorted(random.sample(range(max_val + 1), min(n, max_val + 1)))

    # ==================== 計測制御 ====================
    def start_measurement(self):
        # 接続チェック
        if not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です")
            return
        if not self.gpib_dmm.connected:
            messagebox.showerror("エラー", "DMM (3458A) が未接続です")
            return
        if not self.gpib_scanner.connected:
            messagebox.showerror("エラー", "スキャナー (3499B) が未接続です")
            return
        if not self.position_var.get() and not self.lbc_var.get():
            messagebox.showwarning("警告", "Position または LBC を選択してください")
            return

        defs = self._get_selected_defs()
        if not defs:
            messagebox.showwarning("警告", "Pattern TestタブでDEFを選択してください")
            return

        self._save_settings()
        self.is_running = True
        self._stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        # サマリーをクリア
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)

        self.log("=== Linearity試験 開始 ===", "INFO")

        self._worker_thread = threading.Thread(target=self._measurement_worker, daemon=True)
        self._worker_thread.start()
        self._poll_updates()

    def stop_measurement(self):
        self.log("=== 停止要求 ===", "WARNING")
        self._stop_event.set()

    def _finish(self):
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.target_label.config(text="完了")

    # ==================== DEF情報取得 ====================
    def _get_selected_defs(self):
        selected = []
        if hasattr(self.test_tab, 'def_check_vars'):
            for i, var in enumerate(self.test_tab.def_check_vars):
                if var.get():
                    pos_ch = self.test_tab.scanner_channels_pos[i].get() \
                        if i < len(self.test_tab.scanner_channels_pos) else "ー"
                    neg_ch = self.test_tab.scanner_channels_neg[i].get() \
                        if i < len(self.test_tab.scanner_channels_neg) else "ー"
                    if (pos_ch == "ー" or not pos_ch) and (neg_ch == "ー" or not neg_ch):
                        continue
                    selected.append({
                        'index': i, 'name': f"DEF{i}",
                        'pos_channel': pos_ch, 'neg_channel': neg_ch
                    })
        return selected

    def _get_serial_number(self, def_index):
        try:
            with open('app_settings.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            sn = config.get('comm_profiles', {}).get('1', {}).get('serial_numbers', {})
            return sn.get(f'DEF{def_index}_sn', f'DEF{def_index}')
        except Exception:
            return f'DEF{def_index}'

    # ==================== ワーカースレッド ====================
    def _measurement_worker(self):
        try:
            selected_defs = self._get_selected_defs()
            switch_delay = self._load_switch_delay()
            settle_time = self.settle_time_var.get()

            # スキャナー初期化
            self._queue_update('log', ("スキャナー初期化 (cpon)...", "INFO"))
            self._scanner_cpon()
            time.sleep(0.5)

            dac_types = []
            if self.position_var.get():
                dac_types.append('Position')
            if self.lbc_var.get():
                dac_types.append('LBC')

            for dac_name in dac_types:
                if self._stop_event.is_set():
                    break

                spec = self.DAC_SPECS[dac_name]
                bits = spec['bits']
                span = spec['span']
                ci_cmd = spec['ci']
                center_hex = spec['center']
                dmm_range = spec['dmm_range']

                # パターン生成
                try:
                    pattern_values = self._generate_pattern(bits)
                except ValueError as e:
                    self._queue_update('log', (str(e), "ERROR"))
                    continue

                self._queue_update('log', (
                    f"--- {dac_name} ({bits}bit) 計測開始 ---", "INFO"))
                self._queue_update('log', (
                    f"計測点数: {len(pattern_values)}, 安定待ち: {settle_time}秒", "INFO"))

                # DataGen: A固定モード設定
                self._queue_update('log', (f"DataGen: alt s sa {ci_cmd}", "INFO"))
                self._datagen_send(f"alt s sa {ci_cmd}")
                time.sleep(0.1)

                # DMM Range設定
                self._queue_update('log', (f"DMM Range: DCV {dmm_range}", "INFO"))
                self.gpib_dmm.write(f"DCV {dmm_range}")
                time.sleep(0.3)

                for pole in ['POS', 'NEG']:
                    if self._stop_event.is_set():
                        break

                    pole_cmd = 'p' if pole == 'POS' else 'n'

                    for def_info in selected_defs:
                        if self._stop_event.is_set():
                            break

                        ch_key = 'pos_channel' if pole == 'POS' else 'neg_channel'
                        channel = def_info[ch_key]
                        if not channel or channel == 'ー':
                            self._queue_update('log', (
                                f"{def_info['name']} {pole}: CH未設定、スキップ", "WARNING"))
                            continue

                        ch_number = channel.replace("CH", "")
                        channel_addr = f"@{self.scanner_slot}{ch_number}"

                        self._queue_update('log', (
                            f"▼ {def_info['name']} {dac_name} {pole} (CH:{channel_addr})", "INFO"))
                        self._queue_update('target',
                            f"{def_info['name']} {dac_name} {pole}")

                        # スキャナーCH切替
                        self._switch_scanner(channel_addr, switch_delay)

                        # Center設定
                        self._datagen_send(f"alt a {center_hex} {ci_cmd} {pole_cmd}")
                        time.sleep(settle_time)

                        # 計測ループ
                        x_vals = []
                        y_vals = []
                        total_pts = len(pattern_values)

                        for idx, val in enumerate(pattern_values):
                            if self._stop_event.is_set():
                                break

                            mask = (1 << bits) - 1
                            hex_str = f"{val & mask:05X}"

                            # DAC値設定
                            self._datagen_send(f"alt a {hex_str} {ci_cmd} {pole_cmd}")
                            time.sleep(settle_time)

                            # DMM計測
                            voltage = self._measure_voltage()
                            if voltage is not None:
                                x_vals.append(val)
                                y_vals.append(voltage)

                            # 進捗更新
                            self._queue_update('progress', (idx + 1, total_pts))

                        # Center復帰
                        self._datagen_send(f"alt a {center_hex} {ci_cmd} {pole_cmd}")

                        if self._stop_event.is_set():
                            break

                        if len(x_vals) < 2:
                            self._queue_update('log', (
                                "計測点不足、解析スキップ", "WARNING"))
                            continue

                        # 最小二乗法計算
                        results = self._calculate_linearity(
                            np.array(x_vals, dtype=float),
                            np.array(y_vals, dtype=float),
                            bits, span, (pole == 'NEG')
                        )

                        # NG判定
                        ng_flags = []
                        if abs(results['gain'] - 1.0) > self.th_gain.get():
                            ng_flags.append('Gain')
                        if abs(results['offset']) > self.th_offset.get():
                            ng_flags.append('Offset')
                        if results['max_error'] > self.th_error.get():
                            ng_flags.append('Error')
                        results['ng'] = bool(ng_flags)
                        results['ng_detail'] = ng_flags
                        results['judge'] = 'NG' if ng_flags else 'OK'

                        detail_str = f" ({', '.join(ng_flags)})" if ng_flags else ""
                        self._queue_update('log', (
                            f"  Gain={results['gain']:.6f}, "
                            f"Offset={results['offset']:.3f} LSB, "
                            f"MaxErr={results['max_error']:.3f} LSB "
                            f"→ {results['judge']}{detail_str}",
                            "ERROR" if ng_flags else "SUCCESS"
                        ))

                        # CSV保存
                        serial_no = self._get_serial_number(def_info['index'])
                        csv_path = self._save_csv(
                            results, x_vals, y_vals,
                            dac_name, pole, def_info, serial_no, bits, span)
                        if csv_path:
                            self._queue_update('log', (
                                f"  CSV保存: {csv_path}", "SUCCESS"))

                        # グラフ・サマリー更新をキュー
                        self._queue_update('result', {
                            'def_name': def_info['name'],
                            'dac_name': dac_name,
                            'pole': pole,
                            'results': results,
                            'x_vals': x_vals,
                            'y_vals': y_vals,
                            'bits': bits,
                            'span': span,
                        })

            # 終了処理
            self._scanner_cpon()
            self._queue_update('log', ("=== Linearity試験 完了 ===", "INFO"))
            self._queue_update('done', None)

        except Exception as e:
            import traceback
            self._queue_update('log', (
                f"エラー: {e}\n{traceback.format_exc()}", "ERROR"))
            self._queue_update('done', None)

    # ==================== ハードウェア操作 ====================
    def _scanner_cpon(self):
        """スキャナー全チャンネルOPEN (cpon)"""
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000
        try:
            self.gpib_scanner.write(f":system:cpon {self.scanner_slot}")
        finally:
            self.gpib_scanner.instrument.timeout = orig_timeout

    def _switch_scanner(self, channel_addr, switch_delay):
        """スキャナーCH切替 (cpon方式 - measurement_windowと同じ)"""
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000
        try:
            # 1. cpon: 全CH OPEN
            self.gpib_scanner.write(f":system:cpon {self.scanner_slot}")
            # 2. リレー安定待ち
            time.sleep(switch_delay / 2)
            # 3. *OPC?: cpon完了確認
            self.gpib_scanner.query("*OPC?")
            # 4. CLOSE: 対象CH選択
            self.gpib_scanner.write(f"CLOSE ({channel_addr})")
            # 5. CLOSE後安定待ち
            time.sleep(switch_delay / 2)
        finally:
            self.gpib_scanner.instrument.timeout = orig_timeout

    def _measure_voltage(self):
        """DMM計測 (TRIG SGL)"""
        orig_timeout = self.gpib_dmm.instrument.timeout
        self.gpib_dmm.instrument.timeout = 5000
        try:
            success, response = self.gpib_dmm.query("TRIG SGL")
            if success:
                return float(response.strip())
        except Exception:
            pass
        finally:
            self.gpib_dmm.instrument.timeout = orig_timeout
        return None

    def _datagen_send(self, cmd):
        """DataGenコマンド送信"""
        self.datagen.send_command(cmd)
        time.sleep(0.05)

    def _load_switch_delay(self):
        """スキャナー切替時間を設定ファイルから読み込み"""
        try:
            with open('app_settings.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            return max(0.4, config.get('measurement_window', {}).get('switch_delay_sec', 0.4))
        except Exception:
            return 0.4

    # ==================== 最小二乗法・誤差計算 ====================
    def _calculate_linearity(self, x, y, bits, span, is_neg):
        """最小二乗法で直線性を計算

        Args:
            x: DAC設定値 (numpy array)
            y: 測定電圧 (numpy array)
            bits: DACビット数 (20 or 16)
            span: 電圧スパン (V)
            is_neg: NEG極性フラグ

        Returns:
            dict: gain, offset, max_error, errors, etc.
        """
        n = len(x)
        max_val = (1 << bits) - 1
        v_per_lsb = span / max_val

        # 最小二乗法: y = m*x + b
        sx = np.sum(x)
        sy = np.sum(y)
        sxx = np.sum(x * x)
        sxy = np.sum(x * y)
        denom = n * sxx - sx * sx

        m = (n * sxy - sx * sy) / denom
        b = (sy * sxx - sx * sxy) / denom

        # Gain (LSB/LSB): NEG時は極性反転
        gain = -m / v_per_lsb if is_neg else m / v_per_lsb
        # Offset (LSB)
        offset = b / v_per_lsb

        # 各点の誤差 (LSB)
        y_fit = m * x + b
        errors = (y - y_fit) / v_per_lsb
        max_error = float(np.max(np.abs(errors)))

        return {
            'gain': float(gain),
            'offset': float(offset),
            'max_error': max_error,
            'm': float(m),
            'b': float(b),
            'v_per_lsb': v_per_lsb,
            'errors': errors.tolist(),
            'y_fit': y_fit.tolist(),
        }

    # ==================== CSV保存 ====================
    def _save_csv(self, results, x_vals, y_vals, dac_name, pole,
                  def_info, serial_no, bits, span):
        """計測結果をCSVファイルに保存"""
        save_dir = self.save_dir.get()
        os.makedirs(save_dir, exist_ok=True)

        timestamp = time.strftime('%Y%m%d_%H%M%S')
        filename = f"{serial_no}_{dac_name}_{pole}_linearity_{timestamp}.csv"
        filepath = os.path.join(save_dir, filename)

        v_per_lsb = results['v_per_lsb']
        m = results['m']
        b = results['b']

        try:
            with open(filepath, 'w', encoding='utf-8', newline='') as f:
                # ヘッダ情報
                f.write(f"# Linearity Test Result\n")
                f.write(f"# Serial: {serial_no}\n")
                f.write(f"# DEF: {def_info['name']}\n")
                f.write(f"# DAC: {dac_name} ({bits}bit)\n")
                f.write(f"# Pole: {pole}\n")
                f.write(f"# Gain: {results['gain']:.8f}\n")
                f.write(f"# Offset: {results['offset']:.6f} LSB\n")
                f.write(f"# Max Error: {results['max_error']:.6f} LSB\n")
                f.write(f"# Judge: {results['judge']}\n")
                f.write(f"# Date: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"# V_per_LSB: {v_per_lsb:.10e}\n")
                f.write(f"# Slope(m): {m:.10e}\n")
                f.write(f"# Intercept(b): {b:.10e}\n")
                f.write(f"#\n")
                f.write("Order,DAC_Set,DAC_HEX,Theoretical_V,Measured_V,"
                        "FIT_V,Error_LSB,NG\n")

                mask = (1 << bits) - 1
                for i, (x, y_meas) in enumerate(zip(x_vals, y_vals)):
                    hex_str = f"{x & mask:05X}"
                    theoretical_v = -span / 2 + x * v_per_lsb
                    fit_v = m * x + b
                    error_lsb = results['errors'][i]
                    ng_mark = "*" if abs(error_lsb) > self.th_error.get() else ""
                    f.write(f"{i+1},{x},{hex_str},{theoretical_v:.6f},"
                            f"{y_meas:.6f},{fit_v:.6f},{error_lsb:.4f},{ng_mark}\n")

            return filepath
        except Exception as e:
            self._queue_update('log', (f"CSV保存エラー: {e}", "ERROR"))
            return None

    # ==================== グラフ表示 ====================
    def _show_graph(self, data):
        """誤差散布図を別ウィンドウで表示"""
        x_vals = np.array(data['x_vals'])
        errors = np.array(data['results']['errors'])
        dac_name = data['dac_name']
        pole = data['pole']
        def_name = data['def_name']
        results = data['results']
        th_err = self.th_error.get()

        fig, ax = plt.subplots(figsize=(9, 5))

        # NG時は背景色を変更
        if results['ng']:
            fig.patch.set_facecolor('#fff5f5')

        # 誤差散布図
        ax.scatter(x_vals, errors, s=4, c='blue', alpha=0.7, label='Error')

        # NG閾値ライン（赤破線）
        ax.axhline(y=th_err, color='red', linestyle='--', linewidth=1,
                   label=f'\u00b1{th_err} LSB')
        ax.axhline(y=-th_err, color='red', linestyle='--', linewidth=1)
        ax.axhline(y=0, color='gray', linestyle='-', linewidth=0.5)

        # Gain/Offset情報テキスト
        info = (f"Gain: {results['gain']:.6f}  "
                f"Offset: {results['offset']:.3f} LSB  "
                f"MaxErr: {results['max_error']:.3f} LSB  "
                f"[{results['judge']}]")
        ax.set_title(f"Linearity: {def_name} {dac_name} {pole}\n{info}",
                     fontsize=10)
        ax.set_xlabel("DAC Set Value")
        ax.set_ylabel("Error [LSB]")
        ax.legend(loc='upper right', fontsize=8)
        ax.grid(True, alpha=0.3)

        plt.tight_layout()
        plt.show(block=False)

    # ==================== 更新キュー・ポーリング ====================
    def _queue_update(self, msg_type, data):
        """ワーカースレッドからUI更新をキューに追加"""
        self._update_queue.put((msg_type, data))

    def _poll_updates(self):
        """UIスレッドで更新キューをポーリング"""
        try:
            while not self._update_queue.empty():
                msg_type, data = self._update_queue.get_nowait()

                if msg_type == 'log':
                    self.log(data[0], data[1])
                elif msg_type == 'target':
                    self.target_label.config(text=data)
                elif msg_type == 'progress':
                    current, total = data
                    self.progress_label.config(text=f"{current} / {total}")
                    self.progressbar['maximum'] = total
                    self.progressbar['value'] = current
                elif msg_type == 'result':
                    res = data['results']
                    tag = 'ng' if res['ng'] else 'ok'
                    self.summary_tree.insert('', 'end', values=(
                        data['def_name'], data['dac_name'], data['pole'],
                        f"{res['gain']:.6f}", f"{res['offset']:.3f}",
                        f"{res['max_error']:.3f}", res['judge']
                    ), tags=(tag,))
                    self._show_graph(data)
                elif msg_type == 'done':
                    self._finish()
                    return
        except Exception:
            pass

        if self.is_running:
            self.after(50, self._poll_updates)

    # ==================== ログ ====================
    def log(self, message, level="INFO"):
        """ログ出力"""
        ms = int(time.time() * 1000) % 1000
        timestamp = time.strftime("%H:%M:%S") + f".{ms:03d}"
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
