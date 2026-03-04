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
import matplotlib
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.chart.axis import ChartLines


class LinearityTab(ttk.Frame):
    """Linearity試験タブ - DACの直線性を測定し、最小二乗法で誤差を算出する"""

    # DAC仕様定数
    DAC_SPECS = {
        'Position': {'bits': 20, 'span': 20.0, 'ci': 'ci', 'center': '80000', 'dmm_range': '10'},
        'LBC':      {'bits': 16, 'span': 2.0,  'ci': 'cii', 'center': '08000', 'dmm_range': '1'},
    }

    # 出荷試験 (Ship) NG判定基準 (Excelマクロ DacTestBench5K 準拠)
    SHIP_CRITERIA = {
        'Position': {'inl': 0.75, 'dnl': 0.50},
        'LBC':      {'inl': 0.50, 'dnl': 0.25},
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
        self._log_window = None

        # 設定変数
        self.pattern_mode = tk.StringVar(value='Ship')
        self.num_points = tk.StringVar(value='64')
        self.pattern_file = tk.StringVar(value='')
        self.position_var = tk.BooleanVar(value=True)
        self.lbc_var = tk.BooleanVar(value=False)
        self.settle_time_var = tk.DoubleVar(value=0.2)
        self.th_gain = tk.DoubleVar(value=0.01)
        self.th_offset = tk.DoubleVar(value=10.0)
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
        for value, label in [('Ship', '出荷シーケンス'), ('Random', 'Random'), ('Linear', 'Linear'), ('File', 'File')]:
            ttk.Radiobutton(mode_frame, text=label, variable=self.pattern_mode,
                            value=value, command=self._on_mode_changed).pack(side=tk.LEFT, padx=3)

        pts_frame = ttk.Frame(pat_frame)
        pts_frame.pack(fill=tk.X, pady=2)
        ttk.Label(pts_frame, text="計測点数:").pack(side=tk.LEFT)
        self.pts_combo = ttk.Combobox(pts_frame, textvariable=self.num_points,
                                       values=['32', '64', '128', '256', '512', '1024', '2048', '4096'],
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

        # --- DEF選択&スキャナーCH (test_tabの変数を直接使用→両画面同期) ---
        def_frame = ttk.LabelFrame(parent, text="DEF選択&スキャナーCH", padding=4)
        def_frame.pack(fill=tk.X, pady=(0, 5))

        channel_options = ['ー'] + [f'CH{i:02d}' for i in range(10)]
        save_cmd = lambda e=None: self.test_tab.save_settings()

        # 2列レイアウト (DEF0-2 | DEF3-5)
        col_container = ttk.Frame(def_frame)
        col_container.pack(fill=tk.X)

        for col_idx, start in enumerate([0, 3]):
            col = ttk.Frame(col_container)
            col.pack(side=tk.LEFT, padx=(0, 8) if col_idx == 0 else (0, 0))
            # ヘッダー
            hdr = ttk.Frame(col)
            hdr.pack(fill=tk.X)
            ttk.Label(hdr, text="", width=5).pack(side=tk.LEFT)
            ttk.Label(hdr, text="Pos", width=5, font=('', 8)).pack(side=tk.LEFT, padx=1)
            ttk.Label(hdr, text="Neg", width=5, font=('', 8)).pack(side=tk.LEFT, padx=1)

            for i in range(start, start + 3):
                row = ttk.Frame(col)
                row.pack(fill=tk.X, pady=1)
                ttk.Checkbutton(row, text=f"DEF{i}",
                                variable=self.test_tab.def_check_vars[i],
                                command=self.test_tab.save_settings,
                                width=5).pack(side=tk.LEFT)
                pos_cb = ttk.Combobox(row, textvariable=self.test_tab.scanner_channels_pos[i],
                                       values=channel_options, width=5, state='readonly')
                pos_cb.pack(side=tk.LEFT, padx=2)
                pos_cb.bind('<<ComboboxSelected>>', save_cmd)
                neg_cb = ttk.Combobox(row, textvariable=self.test_tab.scanner_channels_neg[i],
                                       values=channel_options, width=5, state='readonly')
                neg_cb.pack(side=tk.LEFT, padx=2)
                neg_cb.bind('<<ComboboxSelected>>', save_cmd)

        # --- NG判定閾値 ---
        th_frame = ttk.LabelFrame(parent, text="NG判定閾値", padding=4)
        th_frame.pack(fill=tk.X, pady=(0, 5))
        for label, var, unit in [("Gain:", self.th_gain, "LSB/LSB"),
                                  ("Offset:", self.th_offset, "LSB"),
                                  ("Error:", self.th_error, "LSB")]:
            row = ttk.Frame(th_frame)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=label, width=7).pack(side=tk.LEFT)
            ttk.Entry(row, textvariable=var, width=8).pack(side=tk.LEFT, padx=2)
            ttk.Label(row, text=unit).pack(side=tk.LEFT)

        # --- 実行制御 ---
        ctrl_frame = ttk.LabelFrame(parent, text="実行制御", padding=4)
        ctrl_frame.pack(fill=tk.X, pady=(0, 5))

        btn_row = ttk.Frame(ctrl_frame)
        btn_row.pack(fill=tk.X, pady=2)
        self.start_btn = ttk.Button(btn_row, text="開始", command=self.start_measurement, width=10)
        self.start_btn.pack(side=tk.LEFT, padx=2)
        self.stop_btn = ttk.Button(btn_row, text="停止", command=self.stop_measurement,
                                    state=tk.DISABLED, width=10)
        self.stop_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="実行ログ", command=self._open_log_window, width=10).pack(side=tk.LEFT, padx=2)

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

        voltage_row = ttk.Frame(prog_frame)
        voltage_row.pack(fill=tk.X, pady=2)
        ttk.Label(voltage_row, text="計測値:", font=('', 10)).pack(side=tk.LEFT)
        self.voltage_label = ttk.Label(voltage_row, text="--- V",
                                        font=('', 14, 'bold'), foreground='red')
        self.voltage_label.pack(side=tk.LEFT, padx=5)

        prog_row = ttk.Frame(prog_frame)
        prog_row.pack(fill=tk.X, pady=2)
        self.progress_label = ttk.Label(prog_row, text="0 / 0")
        self.progress_label.pack(side=tk.LEFT)
        self.progressbar = ttk.Progressbar(prog_row, length=200, mode='determinate')
        self.progressbar.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        # --- 結果サマリー ---
        summary_frame = ttk.LabelFrame(parent, text="結果サマリー", padding=5)
        summary_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        cols = ('def', 'dac', 'pole', 'gain', 'offset', 'maxerr', 'judge')
        self.summary_tree = ttk.Treeview(summary_frame, columns=cols, show='headings', height=10)
        for col, text, w in [('def', 'DEF', 50), ('dac', 'DAC', 70), ('pole', 'Pole', 45),
                              ('gain', 'Gain', 85), ('offset', 'Offset', 70),
                              ('maxerr', 'MaxErr', 70), ('judge', '判定', 45)]:
            self.summary_tree.heading(col, text=text)
            self.summary_tree.column(col, width=w, anchor=tk.CENTER)
        self.summary_tree.pack(fill=tk.BOTH, expand=True)
        self.summary_tree.tag_configure('ng', background='#f5b7b1')
        self.summary_tree.tag_configure('ok', background='#d5f5e3')

    # ==================== 実行ログウィンドウ ====================
    def _open_log_window(self):
        """実行ログを別ウィンドウで表示"""
        if self._log_window and self._log_window.winfo_exists():
            self._log_window.lift()
            self._log_window.focus_force()
            return

        self._log_window = tk.Toplevel(self)
        self._log_window.title("Linearity 実行ログ")

        # メインウィンドウの右隣に配置
        root = self.winfo_toplevel()
        root.update_idletasks()
        rx = root.winfo_x() + root.winfo_width() + 8
        ry = root.winfo_y()
        self._log_window.geometry(f"600x{root.winfo_height()}+{rx}+{ry}")

        toolbar = ttk.Frame(self._log_window)
        toolbar.pack(fill=tk.X, padx=5, pady=2)
        ttk.Button(toolbar, text="クリア", command=self._clear_log, width=8).pack(side=tk.LEFT)

        self.log_text = ScrolledText(self._log_window, wrap=tk.WORD,
                                      state=tk.DISABLED, font=('Courier', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        self.log_text.tag_config('INFO', foreground='black')
        self.log_text.tag_config('SUCCESS', foreground='green')
        self.log_text.tag_config('WARNING', foreground='orange')
        self.log_text.tag_config('ERROR', foreground='red')

        # 既存ログがあれば復元
        if hasattr(self, '_log_buffer'):
            self.log_text.config(state=tk.NORMAL)
            for msg, level in self._log_buffer:
                ms = 0
                self.log_text.insert(tk.END, f"{msg}\n", level)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)

    def _clear_log(self):
        """ログをクリア"""
        if hasattr(self, 'log_text') and self.log_text.winfo_exists():
            self.log_text.config(state=tk.NORMAL)
            self.log_text.delete('1.0', tk.END)
            self.log_text.config(state=tk.DISABLED)
        if hasattr(self, '_log_buffer'):
            self._log_buffer.clear()

    # ==================== 設定 ====================
    def _load_settings(self):
        try:
            if os.path.exists('app_settings.json'):
                with open('app_settings.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                lin = config.get('linearity', {})
                self.pattern_mode.set(lin.get('pattern_mode', 'Ship'))
                self.num_points.set(lin.get('num_points', '64'))
                self.position_var.set(lin.get('position', True))
                self.lbc_var.set(lin.get('lbc', False))
                self.settle_time_var.set(lin.get('settle_time', 0.2))
                self.th_gain.set(lin.get('th_gain', 0.01))
                self.th_offset.set(lin.get('th_offset', 10.0))
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
        mode = self.pattern_mode.get()
        is_file = mode == 'File'
        is_ship = mode == 'Ship'
        self.file_entry.config(state=tk.NORMAL if is_file else tk.DISABLED)
        self.file_browse_btn.config(state=tk.NORMAL if is_file else tk.DISABLED)
        self.pts_combo.config(state=tk.DISABLED if (is_file or is_ship) else 'readonly')

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
            self._save_settings()

    # ==================== パターン生成 ====================

    # 出荷試験用パターン (Ship mode) - POS/NEGで異なるビット境界テストポイント
    # Position POS: 0x00000→0xFFFFF (コード昇順, ビット境界はLSB側)
    SHIP_PATTERN_POSITION_POS = [
        0x00000, 0x00001, 0x00002, 0x00003, 0x00004, 0x00007, 0x00008,
        0x0000F, 0x00010, 0x0001F, 0x00020, 0x0003F, 0x00040, 0x0007F,
        0x00080, 0x000FF, 0x00100, 0x001FF, 0x00200, 0x003FF, 0x00400,
        0x007FF, 0x00800, 0x00FFF, 0x01000, 0x01FFF, 0x02000, 0x03FFF,
        0x04000, 0x07FFF, 0x08000, 0x0FFFF, 0x10000, 0x1FFFF, 0x20000,
        0x3FFFF, 0x40000, 0x5FFFF, 0x60000, 0x7FFFF, 0x80000, 0x9FFFF,
        0xA0000, 0xBFFFF, 0xC0000, 0xDFFFF, 0xE0000, 0xFFFFF,
    ]
    # Position NEG: ビット境界はMSB側 (コード昇順スイープ)
    SHIP_PATTERN_POSITION_NEG = [
        0x00000, 0x1FFFF, 0x20000, 0x3FFFF, 0x40000, 0x5FFFF, 0x60000,
        0x7FFFF, 0x80000, 0x9FFFF, 0xA0000, 0xBFFFF, 0xC0000, 0xDFFFF,
        0xE0000, 0xEFFFF, 0xF0000, 0xF7FFF, 0xF8000, 0xFBFFF, 0xFC000,
        0xFDFFF, 0xFE000, 0xFEFFF, 0xFF000, 0xFF7FF, 0xFF800, 0xFFBFF,
        0xFFC00, 0xFFDFF, 0xFFE00, 0xFFEFF, 0xFFF00, 0xFFF7F, 0xFFF80,
        0xFFFBF, 0xFFFC0, 0xFFFDF, 0xFFFE0, 0xFFFEF, 0xFFFF0, 0xFFFF7,
        0xFFFF8, 0xFFFFB, 0xFFFFC, 0xFFFFD, 0xFFFFE, 0xFFFFF,
    ]
    # LBC POS: ビット境界はMSB側 (コード昇順スイープ)
    SHIP_PATTERN_LBC_POS = [
        0x0000, 0x0FFF, 0x1000, 0x1FFF, 0x2000, 0x2FFF, 0x3000,
        0x3FFF, 0x4000, 0x4FFF, 0x5000, 0x5FFF, 0x6000, 0x6FFF,
        0x7000, 0x7FFF, 0x8000, 0x8FFF, 0x9000, 0x9FFF, 0xA000,
        0xAFFF, 0xB000, 0xBFFF, 0xC000, 0xCFFF, 0xD000, 0xDFFF,
        0xE000, 0xEFFF, 0xF000, 0xF7FF, 0xF800, 0xFBFF, 0xFC00,
        0xFDFF, 0xFE00, 0xFEFF, 0xFF00, 0xFF7F, 0xFF80, 0xFFBF,
        0xFFC0, 0xFFDF, 0xFFE0, 0xFFEF, 0xFFF0, 0xFFF7, 0xFFF8,
        0xFFFB, 0xFFFC, 0xFFFD, 0xFFFE, 0xFFFF,
    ]
    # LBC NEG: 0x0000→0xFFFF (コード昇順, ビット境界はLSB側)
    SHIP_PATTERN_LBC_NEG = [
        0x0000, 0x0001, 0x0002, 0x0003, 0x0004, 0x0007, 0x0008,
        0x000F, 0x0010, 0x001F, 0x0020, 0x003F, 0x0040, 0x007F,
        0x0080, 0x00FF, 0x0100, 0x01FF, 0x0200, 0x03FF, 0x0400,
        0x07FF, 0x0800, 0x0FFF, 0x1000, 0x1FFF, 0x2000, 0x2FFF,
        0x3000, 0x3FFF, 0x4000, 0x4FFF, 0x5000, 0x5FFF, 0x6000,
        0x6FFF, 0x7000, 0x7FFF, 0x8000, 0x8FFF, 0x9000, 0x9FFF,
        0xA000, 0xAFFF, 0xB000, 0xBFFF, 0xC000, 0xCFFF, 0xD000,
        0xDFFF, 0xE000, 0xEFFF, 0xF000, 0xFFFF,
    ]

    def _generate_pattern(self, bits, pole='POS'):
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

        if mode == 'Ship':
            if bits == 20:
                return list(self.SHIP_PATTERN_POSITION_POS if pole == 'POS'
                            else self.SHIP_PATTERN_POSITION_NEG)
            else:
                return list(self.SHIP_PATTERN_LBC_POS if pole == 'POS'
                            else self.SHIP_PATTERN_LBC_NEG)

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
            messagebox.showwarning("警告", "DEFを選択してください")
            return

        self._save_settings()
        self.is_running = True
        self._stop_event.clear()
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        # サマリーをクリア & ヘッダー切替
        for item in self.summary_tree.get_children():
            self.summary_tree.delete(item)
        if self.pattern_mode.get() == 'Ship':
            for col, text in [('gain', 'INL(max)'), ('offset', 'DNL(max)'), ('maxerr', '')]:
                self.summary_tree.heading(col, text=text)
        else:
            for col, text in [('gain', 'Gain'), ('offset', 'Offset'), ('maxerr', 'MaxErr')]:
                self.summary_tree.heading(col, text=text)

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

                self._queue_update('log', (
                    f"--- {dac_name} ({bits}bit) 計測開始 ---", "INFO"))

                # DataGen: A固定モード設定
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

                    # パターン生成 (出荷シーケンスはPOS/NEGで異なるパターン)
                    try:
                        sweep_values = self._generate_pattern(bits, pole)
                    except ValueError as e:
                        self._queue_update('log', (str(e), "ERROR"))
                        continue

                    self._queue_update('log', (
                        f"  {pole} パターン: {len(sweep_values)}点, "
                        f"安定待ち: {settle_time}秒", "INFO"))

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
                        self._datagen_set_value(center_hex, ci_cmd, pole_cmd)
                        time.sleep(settle_time)

                        # 計測ループ (POS/NEGそれぞれ専用パターンでスイープ)
                        x_vals = []
                        y_vals = []
                        total_pts = len(sweep_values)

                        for idx, val in enumerate(sweep_values):
                            if self._stop_event.is_set():
                                break

                            mask = (1 << bits) - 1
                            hex_str = f"{val & mask:05X}"

                            # DAC値設定
                            self._datagen_set_value(hex_str, ci_cmd, pole_cmd)
                            time.sleep(settle_time)

                            # DMM計測
                            voltage = self._measure_voltage()
                            if voltage is not None:
                                x_vals.append(val)
                                y_vals.append(voltage)
                                self._queue_update('voltage', f"{voltage:.6f} V")
                            else:
                                self._queue_update('voltage', "--- V")

                            # 進捗更新
                            self._queue_update('progress', (idx + 1, total_pts))

                        # Center復帰
                        self._datagen_set_value(center_hex, ci_cmd, pole_cmd)

                        if self._stop_event.is_set():
                            break

                        if len(x_vals) < 2:
                            self._queue_update('log', (
                                "計測点不足、解析スキップ", "WARNING"))
                            continue

                        serial_no = self._get_serial_number(def_info['index'])
                        is_ship = self.pattern_mode.get() == 'Ship'

                        if is_ship:
                            # --- 出荷シーケンスモード: INL/DNL (Excelマクロ準拠) ---
                            # コード昇順にソート (降順パターンの場合に対応)
                            paired = sorted(zip(x_vals, y_vals))
                            x_vals = [p[0] for p in paired]
                            y_vals = [p[1] for p in paired]

                            ship_results = self._calculate_linearity_ship(
                                x_vals, y_vals, bits, dac_name)

                            criteria = self.SHIP_CRITERIA[dac_name]
                            inl_arr = np.array(ship_results['inl'])
                            dnl_vals = [d for d in ship_results['dnl'] if d is not None]
                            dnl_arr = np.array(dnl_vals) if dnl_vals else np.array([0.0])

                            inl_max_p = float(np.max(inl_arr))
                            inl_max_n = float(np.min(inl_arr))
                            dnl_max_p = float(np.max(dnl_arr))
                            dnl_max_n = float(np.min(dnl_arr))

                            inl_ok = max(abs(inl_max_p), abs(inl_max_n)) <= criteria['inl']
                            dnl_ok = max(abs(dnl_max_p), abs(dnl_max_n)) <= criteria['dnl']
                            judge = 'OK' if (inl_ok and dnl_ok) else 'NG'

                            inl_worst = inl_max_n if abs(inl_max_n) >= abs(inl_max_p) else inl_max_p
                            dnl_worst = dnl_max_n if abs(dnl_max_n) >= abs(dnl_max_p) else dnl_max_p

                            self._queue_update('log', (
                                f"  INL: {inl_max_n:+.4f} ~ {inl_max_p:+.4f} LSB "
                                f"(基準: \u00b1{criteria['inl']}), "
                                f"DNL: {dnl_max_n:+.4f} ~ {dnl_max_p:+.4f} LSB "
                                f"(基準: \u00b1{criteria['dnl']}) \u2192 {judge}",
                                "ERROR" if judge == 'NG' else "SUCCESS"
                            ))

                            # XLSX保存
                            xlsx_path, _ = self._save_xlsx_ship(
                                ship_results, x_vals, y_vals,
                                dac_name, pole, def_info, serial_no, bits)
                            if xlsx_path:
                                self._queue_update('log', (
                                    f"  XLSX保存: {xlsx_path}", "SUCCESS"))

                            # グラフ・サマリー更新をキュー
                            self._queue_update('result_ship', {
                                'def_name': def_info['name'],
                                'dac_name': dac_name,
                                'pole': pole,
                                'inl_worst': inl_worst,
                                'dnl_worst': dnl_worst,
                                'judge': judge,
                                'ship_results': ship_results,
                                'x_vals': x_vals,
                                'serial_no': serial_no,
                                'bits': bits,
                            })

                        else:
                            # --- 通常モード: Gain/Offset/Error ---
                            results = self._calculate_linearity(
                                np.array(x_vals, dtype=float),
                                np.array(y_vals, dtype=float),
                                bits, span, (pole == 'NEG')
                            )

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
                                f"\u2192 {results['judge']}{detail_str}",
                                "ERROR" if ng_flags else "SUCCESS"
                            ))

                            # CSV保存
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
        self._queue_update('log', (f"  DG SEND: {cmd}", "INFO"))
        self.datagen.send_command(cmd)
        time.sleep(0.05)

    def _datagen_set_value(self, hex_str, ci_cmd, pole_cmd):
        """DAC値設定（Linearモード時はp/n両方に同じ値を設定）"""
        if self.pattern_mode.get() == 'Linear':
            self._datagen_send(f"alt a {hex_str} {ci_cmd} p")
            self._datagen_send(f"alt a {hex_str} {ci_cmd} n")
        else:
            self._datagen_send(f"alt a {hex_str} {ci_cmd} {pole_cmd}")

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

        # 理論値: コード0 → -span/2, コードmax → +span/2
        # 理論傾き = span / max_val = v_per_lsb
        # 理論切片 = -span / 2
        # POS: コード000000→-10V, コードFFFFF→+10V (傾き正)
        # NEG: コード000000→+10V, コードFFFFF→-10V (傾き負)
        ideal_slope = v_per_lsb if not is_neg else -v_per_lsb
        ideal_intercept = -span / 2 if not is_neg else span / 2

        # Gain (LSB/LSB): 実測傾き / 理論傾き
        gain = m / ideal_slope
        # Offset (LSB): 実測切片と理論切片の差をLSB換算
        offset = (b - ideal_intercept) / v_per_lsb

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

    def _calculate_linearity_ship(self, unsigned_vals, measured_v, bits, dac_name):
        """出荷試験用 INL/DNL計算 (Excelマクロ DacTestBench5K 準拠)

        - 係数: 端点フィット (last_V - first_V) / (last_signed - first_signed)
        - 理論電圧: first_V + coeff * (signed - first_signed)
        - INL: (measured - theoretical) / v_per_lsb
        - DNL: INL[i+1] - INL[i] (コードステップ=1 の場合のみ)
        """
        offset_val = 2 ** (bits - 1)
        signed = np.array([v - offset_val for v in unsigned_vals], dtype=float)

        y = np.array(measured_v, dtype=float)

        # 端点フィット: coeff = (V_last - V_first) / (signed_last - signed_first)
        coeff = float((y[-1] - y[0]) / (signed[-1] - signed[0]))

        # 理論電圧: first_V + coeff * (signed - first_signed)
        theoretical_v = y[0] + coeff * (signed - signed[0])

        # 誤差計算用 V/LSB
        if dac_name == 'Position':
            v_per_lsb = coeff  # Position: 係数 = V/LSB (常に正)
        else:  # LBC: Position 1LSB = 10V / 2^15 = 20V / 2^16
            v_per_lsb = 10.0 / (2 ** (bits - 1))

        # INL
        inl = ((y - theoretical_v) / v_per_lsb).tolist()

        # DNL: INL[i+1] - INL[i] (コードステップ=1 の時のみ)
        dnl = [None] * len(unsigned_vals)
        for i in range(len(unsigned_vals) - 1):
            if unsigned_vals[i + 1] - unsigned_vals[i] == 1:
                dnl[i] = inl[i + 1] - inl[i]

        return {
            'coeff': coeff,
            'signed': signed.tolist(),
            'theoretical_v': theoretical_v.tolist(),
            'inl': inl,
            'dnl': dnl,
            'v_per_lsb': v_per_lsb,
        }

    # ==================== CSV保存 ====================
    def _save_csv(self, results, x_vals, y_vals, dac_name, pole,
                  def_info, serial_no, bits, span):
        """計測結果をCSVファイルに保存"""
        save_dir = self.save_dir.get()
        os.makedirs(save_dir, exist_ok=True)

        timestamp = time.strftime('%Y%m%d_%H%M%S')
        mode = '出荷シーケンス' if self.pattern_mode.get() == 'Ship' else self.pattern_mode.get()
        filename = f"{serial_no}_{dac_name}_{pole}_linearity_{mode}_{timestamp}.csv"
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

    def _save_xlsx_ship(self, ship_results, unsigned_vals, measured_v,
                        dac_name, pole, def_info, serial_no, bits):
        """出荷試験結果をExcelマクロと同じxlsx形式で保存（セル参照式付き）"""
        wb = Workbook()
        ws = wb.active
        ws.title = f"{serial_no}{pole[0]}"[:31]

        n_pts = len(unsigned_vals)
        first_r = 7          # 最初のデータ行 (Excel 1-indexed)
        last_r = first_r + n_pts - 1
        hex_width = 5 if bits == 20 else 4
        mask = (1 << bits) - 1
        criteria = self.SHIP_CRITERIA[dac_name]

        # --- スタイル定義 ---
        bold_font = Font(size=9)
        data_font = Font(size=8)
        ng_font = Font(size=8, color='FF0000')
        coeff_font = Font(size=9)
        thin = Side(style='thin')
        medium = Side(style='medium')

        fmt_v = '0.000000_ '     # 電圧 6桁
        fmt_err = '0.00_ '       # INL/DNL 2桁
        fmt_sci = '##0.0E+0'     # 係数 科学表記

        # --- 列幅設定 ---
        col_widths = {'A': 5, 'B': 12, 'C': 12, 'D': 10, 'E': 15, 'F': 15, 'G': 15, 'H': 15}
        for col_letter, w in col_widths.items():
            ws.column_dimensions[col_letter].width = w

        # --- B,C列を非表示 ---
        ws.column_dimensions['B'].hidden = True
        ws.column_dimensions['C'].hidden = True

        # --- Row 3: 理論直線係数 (式) ---
        ws['E3'] = "理論直線係数"
        ws['E3'].font = bold_font
        ws['F3'] = f'=(E{last_r}-E{first_r})/(C{last_r}-C{first_r})'
        ws['F3'].number_format = fmt_sci
        ws['F3'].font = coeff_font
        if dac_name == 'LBC':
            ws['G3'] = "Position　１LSB"
            ws['G3'].font = bold_font
            ws['H3'] = 10.0 / (2 ** (bits - 1))
            ws['H3'].font = coeff_font

        # --- Headers (Row 5-6) ---
        headers_r5 = {
            'A': '№', 'D': 'コード', 'E': '測定電圧(V)', 'F': '理論電圧(V)',
            'G': '対Position INL' if dac_name == 'LBC' else 'INL',
            'H': '対Position DNL' if dac_name == 'LBC' else 'DNL',
        }
        headers_r6 = {'B': '設定DAC値', 'G': '誤差(LSB)', 'H': '誤差(LSB)'}
        for col, text in headers_r5.items():
            cell = ws[f'{col}5']
            cell.value = text
            cell.font = bold_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
        for col, text in headers_r6.items():
            cell = ws[f'{col}6']
            cell.value = text
            cell.font = bold_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # --- ヘッダー結合セル (参照XLSに合わせる) ---
        ws.merge_cells('A5:A6')
        ws.merge_cells('D5:D6')
        ws.merge_cells('E5:E6')
        ws.merge_cells('F5:F6')

        # --- ヘッダー枠線 (Row 5-6, A-H) ---
        for r in [5, 6]:
            for c in range(1, 9):
                cell = ws.cell(row=r, column=c)
                top = medium if r == 5 else None
                if r == 6:
                    bottom = medium
                elif c in (7, 8):
                    bottom = None      # G5,H5は下線なし
                else:
                    bottom = thin
                left = medium if c == 1 else thin
                right = medium if c == 8 else thin
                cell.border = Border(
                    top=top, bottom=bottom, left=left, right=right)

        # --- Data rows ---
        center_align = Alignment(horizontal='center')
        right_align = Alignment(horizontal='right')
        inl_ref = '$F$3' if dac_name == 'Position' else '$H$3'
        is_neg_display = (pole == 'NEG')

        # NEG: B,C,D列はPOSパターンのコードを表示 (参照XLS準拠)
        if is_neg_display:
            if bits == 20:
                display_codes = list(self.SHIP_PATTERN_POSITION_POS)
            else:
                display_codes = list(self.SHIP_PATTERN_LBC_POS)
            # 表示順のINL/DNL計算 (Excel式と同等、NG色判定用)
            disp_v = [measured_v[n_pts - 1 - j] for j in range(n_pts)]
            disp_coeff = (disp_v[-1] - disp_v[0]) / (display_codes[-1] - display_codes[0])
            disp_theoretical = [disp_v[0] + disp_coeff * (display_codes[j] - display_codes[0])
                                for j in range(n_pts)]
            v_per_lsb_disp = disp_coeff if dac_name == 'Position' else 10.0 / (2 ** (bits - 1))
            disp_inl = [(disp_v[j] - disp_theoretical[j]) / v_per_lsb_disp for j in range(n_pts)]
            disp_dnl = [None] * n_pts
            for j in range(n_pts - 1):
                if display_codes[j + 1] - display_codes[j] == 1:
                    disp_dnl[j] = disp_inl[j + 1] - disp_inl[j]
        else:
            display_codes = None
            disp_inl = ship_results['inl']
            disp_dnl = ship_results['dnl']

        offset_val = 2 ** (bits - 1)
        for i, uval in enumerate(unsigned_vals):
            r = first_r + i

            # NEG: コード列はPOSパターン、電圧列は測定順逆
            if display_codes is not None:
                disp_uval = display_codes[i]
                disp_signed = disp_uval - offset_val
                v_idx = n_pts - 1 - i
            else:
                disp_uval = uval
                disp_signed = int(ship_results['signed'][i])
                v_idx = i

            c_a = ws.cell(row=r, column=1, value=i + 1)                            # A: №
            c_a.font = data_font
            c_a.alignment = center_align
            c_b = ws.cell(row=r, column=2, value=disp_signed)                      # B: 設定DAC値
            c_b.font = data_font
            c_b.alignment = center_align
            c_c = ws.cell(row=r, column=3, value=disp_uval)                        # C: unsigned
            c_c.font = data_font
            c_c.alignment = center_align
            c_d = ws.cell(row=r, column=4, value=f"{disp_uval & mask:0{hex_width}X}")  # D: HEX
            c_d.font = data_font
            c_d.alignment = center_align

            c_e = ws.cell(row=r, column=5, value=measured_v[v_idx])                # E: 測定電圧
            c_e.number_format = fmt_v
            c_e.font = data_font
            c_e.alignment = right_align

            c_f = ws.cell(row=r, column=6)                                         # F: 理論電圧 (式)
            c_f.value = f'=$F$3*C{r}+$E${first_r}'
            c_f.number_format = fmt_v
            c_f.font = data_font
            c_f.alignment = right_align

            c_g = ws.cell(row=r, column=7)                                         # G: INL (式)
            c_g.value = f'=(E{r}-F{r})/{inl_ref}'
            c_g.number_format = fmt_err
            c_g.font = ng_font if abs(disp_inl[i]) > criteria['inl'] else data_font

            # H: DNL (式 - 表示コードステップ=1の場合のみ)
            if disp_dnl[i] is not None:
                c_h = ws.cell(row=r, column=8)
                c_h.value = f'=G{r + 1}-G{r}'
                c_h.number_format = fmt_err
                c_h.font = ng_font if abs(disp_dnl[i]) > criteria['dnl'] else data_font

            # データ行枠線
            for c in range(1, 9):
                cell = ws.cell(row=r, column=c)
                is_last = (i == n_pts - 1)
                cell.border = Border(
                    top=thin,
                    bottom=medium if is_last else thin,
                    left=medium if c == 1 else thin,
                    right=medium if c == 8 else thin)

        # --- Summary rows ---
        sum_start = last_r + 1
        g_range = f'G{first_r}:G{last_r}'
        h_range = f'H{first_r}:H{last_r}'

        summary_data = [
            ('+最大誤差', f'=MAX({g_range})', f'=MAX({h_range})'),
            ('-最大誤差', f'=MIN({g_range})', f'=MIN({h_range})'),
            ('判定基準',   criteria['inl'],      criteria['dnl']),
        ]
        for offset, (label, inl_v, dnl_v) in enumerate(summary_data):
            r = sum_start + offset
            c_f = ws.cell(row=r, column=6, value=label)
            c_f.font = data_font
            c_f.alignment = center_align
            c_g = ws.cell(row=r, column=7, value=inl_v)
            c_g.number_format = fmt_err
            c_g.font = data_font
            c_h = ws.cell(row=r, column=8, value=dnl_v)
            c_h.number_format = fmt_err
            c_h.font = data_font
            # 規格行はG,H右寄せ (参照XLSに合わせる)
            if offset == 2:
                c_g.alignment = right_align
                c_h.alignment = right_align
            # 枠線
            for c in range(6, 9):
                cell = ws.cell(row=r, column=c)
                cell.border = Border(
                    top=medium if offset == 0 else thin,
                    bottom=thin,
                    left=medium if c == 6 else thin,
                    right=medium if c == 8 else thin)

        # 合否判定行 (Excel数式 - 参照XLS準拠)
        r_judge = sum_start + 3
        r_plus = sum_start      # +最大誤差行
        r_minus = sum_start + 1  # -最大誤差行
        r_crit = sum_start + 2   # 判定基準行

        # NG色判定用 (フォント色を事前決定)
        inl_arr = np.array(disp_inl)
        dnl_vals = [d for d in disp_dnl if d is not None]
        dnl_arr = np.array(dnl_vals) if dnl_vals else np.array([0.0])
        inl_ok = float(np.max(np.abs(inl_arr))) <= criteria['inl']
        dnl_ok = float(np.max(np.abs(dnl_arr))) <= criteria['dnl'] if len(dnl_vals) > 0 else True

        judge_label_font = Font(size=11)
        judge_value_font = Font(size=26, color='FF0000') if not (inl_ok and dnl_ok) else Font(size=26)
        c_f_judge = ws.cell(row=r_judge, column=6, value='判定結果')
        c_f_judge.font = judge_label_font
        c_f_judge.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[r_judge].height = 33.6
        ws.merge_cells(f'G{r_judge}:H{r_judge}')
        judge_formula = (f'=IF((ABS(G{r_plus})<G{r_crit})'
                         f'+(ABS(H{r_plus})<H{r_crit})'
                         f'+(ABS(G{r_minus})<G{r_crit})'
                         f'+(ABS(H{r_minus})<H{r_crit})=4,"OK","NG")')
        c_judge = ws.cell(row=r_judge, column=7, value=judge_formula)
        c_judge.font = judge_value_font
        c_judge.alignment = center_align
        for c in range(6, 9):
            ws.cell(row=r_judge, column=c).border = Border(
                top=thin, bottom=medium,
                left=medium if c == 6 else thin,
                right=medium if c == 8 else thin)

        # --- グラフ作成 (参照XLS準拠) ---
        # 主チャート (INL: 左Y軸)
        chart = LineChart()
        chart.title = f"{def_info['name']} {pole} {dac_name} 直線性({n_pts}点シーケンシャル測定)"
        chart.width = 19.5
        chart.height = 17.5

        # 左Y軸: INL(LSB)
        chart.y_axis.title = "INL(LSB)"
        chart.y_axis.scaling.min = -1.5
        chart.y_axis.scaling.max = 1.5
        chart.y_axis.majorUnit = 0.25
        chart.y_axis.majorGridlines = ChartLines()
        chart.y_axis.minorGridlines = None
        chart.y_axis.numFmt = '0.00_ '
        chart.y_axis.crossBetween = "midCat"

        # X軸: ラベルを上部に配置、2つおきに表示
        chart.x_axis.title = None
        chart.x_axis.majorGridlines = None
        chart.x_axis.tickLblPos = 'high'
        chart.x_axis.tickLblSkip = 2

        # INL系列: マゼンタ線 + マゼンタ四角マーカー
        inl_values = Reference(ws, min_col=7, min_row=first_r, max_row=last_r)
        inl_series = Series(inl_values, title="INL")
        inl_series.graphicalProperties.line.solidFill = "FF00FF"
        inl_series.marker.symbol = "square"
        inl_series.marker.size = 5
        inl_series.marker.graphicalProperties.solidFill = "FF00FF"
        chart.append(inl_series)

        # 副チャート (DNL: 右Y軸)
        chart2 = LineChart()
        chart2.y_axis.title = "DNL(LSB)"
        chart2.y_axis.scaling.min = -1.5
        chart2.y_axis.scaling.max = 1.5
        chart2.y_axis.majorUnit = 0.25
        chart2.y_axis.numFmt = '0.00_ '
        chart2.y_axis.axId = 200

        # DNL系列: 線なし(マーカーのみ) + ネイビーダイヤモンドマーカー
        dnl_values = Reference(ws, min_col=8, min_row=first_r, max_row=last_r)
        dnl_series = Series(dnl_values, title="DNL")
        dnl_series.graphicalProperties.line.noFill = True
        dnl_series.marker.symbol = "diamond"
        dnl_series.marker.size = 5
        dnl_series.marker.graphicalProperties.solidFill = "000080"
        chart2.append(dnl_series)

        # チャート結合 (DNLを第2軸に)
        chart += chart2

        # 凡例: 下部
        chart.legend.position = 'b'

        ws.add_chart(chart, "J1")

        # --- ファイル保存 ---
        save_dir = self.save_dir.get()
        os.makedirs(save_dir, exist_ok=True)
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        filename = f"{serial_no}_{dac_name}_{pole}_linearity_出荷シーケンス_{timestamp}.xlsx"
        filepath = os.path.join(save_dir, filename)

        try:
            wb.save(filepath)
            return filepath, inl_ok and dnl_ok
        except Exception as e:
            self._queue_update('log', (f"XLSX保存エラー: {e}", "ERROR"))
            return None, False

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

    def _save_graph_ship_png(self, data):
        """出荷試験用 INL/DNL折れ線グラフをPNGで保存（参考ファイル準拠）"""
        n_pts = len(data['x_vals'])
        x_count = np.arange(1, n_pts + 1)  # 横軸: 回数 (1, 2, 3, ...)
        ship_results = data['ship_results']
        inl = np.array(ship_results['inl'])
        dac_name = data['dac_name']
        pole = data['pole']
        serial_no = data['serial_no']
        criteria = self.SHIP_CRITERIA[dac_name]

        # 日本語フォント設定
        matplotlib.rcParams['font.family'] = 'MS Gothic'

        fig, ax = plt.subplots(figsize=(10, 5))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('white')

        # INL 折れ線 (青、丸マーカー)
        ax.plot(x_count, inl, color='#4472C4', linewidth=1, marker='o',
                markersize=3, label='INL', zorder=3)

        # DNL 折れ線 (緑、三角マーカー)
        dnl_x, dnl_y = [], []
        for i, d in enumerate(ship_results['dnl']):
            if d is not None:
                dnl_x.append(i + 1)
                dnl_y.append(d)
        if dnl_x:
            ax.plot(dnl_x, dnl_y, color='#70AD47', linewidth=1, marker='^',
                    markersize=3, label='DNL', zorder=3)
            ax.axhline(y=criteria['dnl'], color='#70AD47', linestyle='--',
                       linewidth=0.8, alpha=0.7)
            ax.axhline(y=-criteria['dnl'], color='#70AD47', linestyle='--',
                       linewidth=0.8, alpha=0.7)

        # INL規格ライン (赤破線)
        ax.axhline(y=criteria['inl'], color='red', linestyle='--', linewidth=0.8,
                   label=f'INL \u00b1{criteria["inl"]} LSB')
        ax.axhline(y=-criteria['inl'], color='red', linestyle='--', linewidth=0.8)
        ax.axhline(y=0, color='black', linestyle='-', linewidth=0.5)

        # 軸・タイトル
        ax.set_title(f'{serial_no}  {dac_name}  {pole}', fontsize=11, fontweight='bold')
        ax.set_xlabel('回数', fontsize=10)
        ax.set_ylabel('誤差 (LSB)', fontsize=10)
        ax.legend(loc='upper right', fontsize=8, framealpha=0.9)
        ax.grid(True, linestyle='-', linewidth=0.3, alpha=0.5)
        ax.set_xlim(0.5, n_pts + 0.5)

        plt.tight_layout()

        # PNG保存
        save_dir = self.save_dir.get()
        os.makedirs(save_dir, exist_ok=True)
        timestamp = time.strftime('%Y%m%d_%H%M%S')
        png_filename = f"{serial_no}_{dac_name}_{pole}_linearity_出荷シーケンス_{timestamp}.png"
        png_path = os.path.join(save_dir, png_filename)
        fig.savefig(png_path, dpi=150, bbox_inches='tight')
        plt.close(fig)

        return png_path

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
                elif msg_type == 'voltage':
                    self.voltage_label.config(text=data)
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
                elif msg_type == 'result_ship':
                    tag = 'ng' if data['judge'] == 'NG' else 'ok'
                    self.summary_tree.insert('', 'end', values=(
                        data['def_name'], data['dac_name'], data['pole'],
                        f"{data['inl_worst']:.4f}", f"{data['dnl_worst']:.4f}",
                        '', data['judge']
                    ), tags=(tag,))
                    png_path = self._save_graph_ship_png(data)
                    if png_path:
                        self.log(f"  PNG保存: {png_path}", "SUCCESS")
                elif msg_type == 'done':
                    self._finish()
                    return
        except Exception:
            pass

        if self.is_running:
            self.after(50, self._poll_updates)

    # ==================== ログ ====================
    def log(self, message, level="INFO"):
        """ログ出力（バッファに蓄積 + ウィンドウが開いていれば表示）"""
        ms = int(time.time() * 1000) % 1000
        timestamp = time.strftime("%H:%M:%S") + f".{ms:03d}"
        formatted = f"[{timestamp}] {message}"

        # バッファに蓄積
        if not hasattr(self, '_log_buffer'):
            self._log_buffer = []
        self._log_buffer.append((formatted, level))

        # ログウィンドウが開いていれば書き込み
        if self._log_window and self._log_window.winfo_exists():
            try:
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, formatted + "\n", level)
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
            except tk.TclError:
                pass
