import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import json
import os
from utils.graph_plotter import LSBGraphPlotter


class GraphTab(ttk.Frame):
    """
    グラフ描画タブ
    
    機能:
    - CSVファイル選択・読み込み
    - LSB変換設定（bit精度、基準電圧、LSB/div）- 自動保存機能付き
    - 表示データ選択（シリアルNo.、POS/NEG）- デフォルト全選択
    - グラフ表示・自動更新
    """
    
    SETTINGS_FILE = "graph_settings.json"
    
    def __init__(self, parent, gpib_controller):
        super().__init__(parent)
        self.gpib = gpib_controller
        
        # データ保持
        self.csv_data = None
        self.temp_csv_data = None  # 温度CSVデータ
        self.serial_numbers = []
        self.checkboxes = {}
        self.graph_windows = []  # 開いているグラフウィンドウのリスト
        
        # 設定変更の監視フラグ（初期化中は保存しない）
        self._initializing = True
        
        self.create_widgets()
        self.load_settings()
        
        # 設定値の変更を監視
        self._setup_auto_save()
        
        # 初期化完了
        self._initializing = False
    
    def create_widgets(self):
        """ウィジェットを作成"""
        # ファイル選択フレーム
        self._create_file_selection_frame()

        # 設定フレーム（グラフ表示ボタン含む）
        self._create_settings_frame()

        # データ選択フレーム
        self._create_data_selection_frame()
    
    def _create_file_selection_frame(self):
        """ファイル選択フレームを作成"""
        file_frame = ttk.LabelFrame(self, text="CSVファイル選択", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=5)

        # 測定CSVファイル選択
        measurement_frame = ttk.Frame(file_frame)
        measurement_frame.pack(fill=tk.X, pady=2)
        ttk.Label(measurement_frame, text="測定CSV:", width=10).pack(side=tk.LEFT, padx=5)
        self.file_path_var = tk.StringVar()
        ttk.Entry(measurement_frame, textvariable=self.file_path_var, state='readonly',
                  width=70).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(measurement_frame, text="読み込み",
                   command=self.load_csv_file).pack(side=tk.LEFT, padx=5)

        # 温度CSVファイル選択
        temp_frame = ttk.Frame(file_frame)
        temp_frame.pack(fill=tk.X, pady=2)
        ttk.Label(temp_frame, text="温度CSV:", width=10).pack(side=tk.LEFT, padx=5)
        self.temp_file_path_var = tk.StringVar()
        ttk.Entry(temp_frame, textvariable=self.temp_file_path_var, state='readonly',
                  width=70).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(temp_frame, text="読み込み",
                   command=self.load_temp_csv_file).pack(side=tk.LEFT, padx=5)
    
    def _create_settings_frame(self):
        """設定フレームを作成"""
        # 設定フレームを横並びにするためのコンテナ（均等幅）
        settings_container = ttk.Frame(self)
        settings_container.pack(fill=tk.X, padx=10, pady=5)
        settings_container.columnconfigure(0, weight=1)
        settings_container.columnconfigure(1, weight=1)

        # 左側: 変換設定
        settings_frame = ttk.LabelFrame(settings_container, text="変換設定", padding=10)
        settings_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        # Bit精度設定
        bit_frame = ttk.Frame(settings_frame)
        bit_frame.pack(fill=tk.X, pady=2)
        ttk.Label(bit_frame, text="Bit精度:").pack(side=tk.LEFT, padx=5)
        self.bit_precision_var = tk.StringVar(value="24")
        ttk.Entry(bit_frame, textvariable=self.bit_precision_var,
                  width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(bit_frame, text="bit").pack(side=tk.LEFT)

        # 基準電圧設定（幅を2/3に）
        ref_frame = ttk.Frame(settings_frame)
        ref_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ref_frame, text="+Full基準電圧:").pack(side=tk.LEFT, padx=5)
        self.pos_full_var = tk.StringVar(value="10.0")
        ttk.Entry(ref_frame, textvariable=self.pos_full_var,
                  width=7).pack(side=tk.LEFT, padx=5)
        ttk.Label(ref_frame, text="V").pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(ref_frame, text="-Full基準電圧:").pack(side=tk.LEFT, padx=5)
        self.neg_full_var = tk.StringVar(value="-10.0")
        ttk.Entry(ref_frame, textvariable=self.neg_full_var,
                  width=7).pack(side=tk.LEFT, padx=5)
        ttk.Label(ref_frame, text="V").pack(side=tk.LEFT)

        # Y軸範囲設定（2行に分割）
        yaxis_frame1 = ttk.Frame(settings_frame)
        yaxis_frame1.pack(fill=tk.X, pady=2)
        ttk.Label(yaxis_frame1, text="Y軸範囲:", width=10).pack(side=tk.LEFT, padx=5)
        self.yaxis_mode_var = tk.StringVar(value="auto")
        ttk.Radiobutton(yaxis_frame1, text="オート", variable=self.yaxis_mode_var,
                        value="auto").pack(side=tk.LEFT)

        yaxis_frame2 = ttk.Frame(settings_frame)
        yaxis_frame2.pack(fill=tk.X, pady=2)
        ttk.Label(yaxis_frame2, text="", width=10).pack(side=tk.LEFT, padx=5)  # インデント用
        ttk.Radiobutton(yaxis_frame2, text="設定値", variable=self.yaxis_mode_var,
                        value="manual").pack(side=tk.LEFT)
        ttk.Label(yaxis_frame2, text="Min:").pack(side=tk.LEFT, padx=(10, 2))
        self.yaxis_min_var = tk.StringVar(value="-50")
        ttk.Entry(yaxis_frame2, textvariable=self.yaxis_min_var,
                  width=8).pack(side=tk.LEFT, padx=2)
        ttk.Label(yaxis_frame2, text="Max:").pack(side=tk.LEFT, padx=(10, 2))
        self.yaxis_max_var = tk.StringVar(value="50")
        ttk.Entry(yaxis_frame2, textvariable=self.yaxis_max_var,
                  width=8).pack(side=tk.LEFT, padx=2)
        ttk.Label(yaxis_frame2, text="LSB").pack(side=tk.LEFT)

        # LSB/Div設定 - 順序変更
        lsb_frame = ttk.Frame(settings_frame)
        lsb_frame.pack(fill=tk.X, pady=2)
        ttk.Label(lsb_frame, text="縦軸LSB/Div:").pack(side=tk.LEFT, padx=5)
        self.lsb_per_div_var = tk.StringVar(value="10")
        ttk.Entry(lsb_frame, textvariable=self.lsb_per_div_var,
                  width=10).pack(side=tk.LEFT, padx=5)
        ttk.Label(lsb_frame, text="LSB").pack(side=tk.LEFT)

        # 基準電圧計算方法
        ref_mode_frame = ttk.Frame(settings_frame)
        ref_mode_frame.pack(fill=tk.X, pady=2)
        ttk.Label(ref_mode_frame, text="基準電圧:").pack(side=tk.LEFT, padx=5)
        self.ref_mode_var = tk.StringVar(value="ideal")
        ttk.Radiobutton(ref_mode_frame, text="理想値", variable=self.ref_mode_var,
                        value="ideal").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(ref_mode_frame, text="全平均", variable=self.ref_mode_var,
                        value="all_avg").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(ref_mode_frame, text="区間別平均", variable=self.ref_mode_var,
                        value="section_avg").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(ref_mode_frame, text="初回平均", variable=self.ref_mode_var,
                        value="first_avg").pack(side=tk.LEFT, padx=5)

        # スキップ設定フレーム
        skip_frame = ttk.Frame(settings_frame)
        skip_frame.pack(fill=tk.X, pady=2)

        # コード切替後スキップ行数
        ttk.Label(skip_frame, text="切替後スキップ:").pack(side=tk.LEFT, padx=5)
        self.skip_after_change_var = tk.StringVar(value="0")
        ttk.Entry(skip_frame, textvariable=self.skip_after_change_var,
                  width=5).pack(side=tk.LEFT, padx=2)
        ttk.Label(skip_frame, text="行").pack(side=tk.LEFT, padx=(0, 15))

        # パターン開始最初のデータを省く
        self.skip_first_data_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(skip_frame, text="開始時スキップ",
                        variable=self.skip_first_data_var).pack(side=tk.LEFT, padx=5)

        # 切替わり1周回前を省く
        self.skip_before_change_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(skip_frame, text="切替前スキップ",
                        variable=self.skip_before_change_var).pack(side=tk.LEFT, padx=5)

        # グラフ表示ボタン（変換設定内）
        ttk.Separator(settings_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        ttk.Button(settings_frame, text="選択したデータをグラフ表示",
                   command=self.plot_selected_data).pack(pady=5)

        # 右側: 温特グラフ設定
        temp_settings_frame = ttk.LabelFrame(settings_container, text="温特グラフ設定", padding=15)
        temp_settings_frame.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        # 温特グラフで使用する設定（共通設定を参照）
        ttk.Label(temp_settings_frame, text="【共通設定を使用】",
                  font=('', 9, 'bold')).pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(temp_settings_frame, text="・Bit精度").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・基準電圧 (+Full/-Full)").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・スキップ設定").pack(anchor=tk.W, padx=15, pady=1)

        # Y軸(LSB)設定
        ttk.Separator(temp_settings_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=12)
        ttk.Label(temp_settings_frame, text="【Y軸(LSB)設定】",
                  font=('', 9, 'bold')).pack(anchor=tk.W, pady=(0, 8))

        # 選択式: デフォルト or 共通設定
        self.temp_yaxis_mode_var = tk.StringVar(value="default")
        temp_yaxis_radio_frame = ttk.Frame(temp_settings_frame)
        temp_yaxis_radio_frame.pack(fill=tk.X, pady=3)
        ttk.Radiobutton(temp_yaxis_radio_frame, text="デフォルト(±8LSB 2LSB/Div)",
                        variable=self.temp_yaxis_mode_var,
                        value="default").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(temp_yaxis_radio_frame, text="共通設定を使用",
                        variable=self.temp_yaxis_mode_var,
                        value="common").pack(side=tk.LEFT, padx=10)

        # デフォルト値は保持（内部で使用）
        self.temp_yaxis_min_var = tk.StringVar(value="-8")
        self.temp_yaxis_max_var = tk.StringVar(value="8")

        # 固定値の注記
        ttk.Separator(temp_settings_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=12)
        ttk.Label(temp_settings_frame, text="【固定値】",
                  font=('', 9, 'bold')).pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(temp_settings_frame, text="・基準電圧モード: 初回平均").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・Y軸(温度): ±8℃, 2℃/Div").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・X軸: 10分/Div").pack(anchor=tk.W, padx=15, pady=1)

        # 温特グラフ表示ボタン（温特グラフ設定内）
        ttk.Separator(temp_settings_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=8)
        ttk.Button(temp_settings_frame, text="温特グラフ表示",
                   command=self.plot_temperature_graph).pack(pady=8)
    
    def _create_data_selection_frame(self):
        """データ選択フレームを作成"""
        selection_frame = ttk.LabelFrame(self, text="表示データ選択", padding=10)
        selection_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # スクロール可能なフレーム
        canvas = tk.Canvas(selection_frame, height=100)
        scrollbar = ttk.Scrollbar(selection_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = ttk.Frame(canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def _setup_auto_save(self):
        """設定値の自動保存を設定"""
        # StringVarの変更を監視
        self.bit_precision_var.trace_add('write', self._on_setting_changed)
        self.pos_full_var.trace_add('write', self._on_setting_changed)
        self.neg_full_var.trace_add('write', self._on_setting_changed)
        self.lsb_per_div_var.trace_add('write', self._on_setting_changed)
        self.ref_mode_var.trace_add('write', self._on_setting_changed)
        self.yaxis_mode_var.trace_add('write', self._on_setting_changed)
        self.yaxis_min_var.trace_add('write', self._on_setting_changed)
        self.yaxis_max_var.trace_add('write', self._on_setting_changed)
        self.skip_after_change_var.trace_add('write', self._on_setting_changed)
        self.skip_first_data_var.trace_add('write', self._on_setting_changed)
        self.skip_before_change_var.trace_add('write', self._on_setting_changed)
        # 温特グラフ用Y軸設定
        self.temp_yaxis_mode_var.trace_add('write', self._on_setting_changed)
    
    def _on_setting_changed(self, *args):
        """設定値が変更されたときに自動保存＆グラフ更新"""
        if not self._initializing:
            self.save_settings()
            self.update_all_graphs()
    
    def load_settings(self):
        """設定を読み込み"""
        if os.path.exists(self.SETTINGS_FILE):
            try:
                with open(self.SETTINGS_FILE, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                self.bit_precision_var.set(settings.get('bit_precision', '24'))
                self.pos_full_var.set(settings.get('pos_full_voltage', '10.0'))
                self.neg_full_var.set(settings.get('neg_full_voltage', '-10.0'))
                self.lsb_per_div_var.set(settings.get('lsb_per_div', '10'))
                self.ref_mode_var.set(settings.get('ref_mode', 'ideal'))
                self.yaxis_mode_var.set(settings.get('yaxis_mode', 'auto'))
                self.yaxis_min_var.set(settings.get('yaxis_min', '-50'))
                self.yaxis_max_var.set(settings.get('yaxis_max', '50'))
                self.skip_after_change_var.set(settings.get('skip_after_change', '0'))
                self.skip_first_data_var.set(settings.get('skip_first_data', False))
                self.skip_before_change_var.set(settings.get('skip_before_change', False))
                # 温特グラフ用Y軸設定（起動時は常にデフォルト）
                # self.temp_yaxis_mode_var は初期値 "default" のまま

                # CSVファイルパスを復元
                csv_path = settings.get('csv_file_path', '')
                if csv_path and os.path.exists(csv_path):
                    self.file_path_var.set(csv_path)
                    self._load_csv_from_path(csv_path, show_message=False)

                # 温度CSVファイルパスを復元
                temp_csv_path = settings.get('temp_csv_file_path', '')
                if temp_csv_path and os.path.exists(temp_csv_path):
                    self.temp_file_path_var.set(temp_csv_path)
                    self._load_temp_csv_from_path(temp_csv_path, show_message=False)

            except Exception as e:
                print(f"設定の読み込みに失敗しました: {e}")
    
    def save_settings(self):
        """設定を保存（自動）"""
        try:
            settings = {
                'bit_precision': self.bit_precision_var.get(),
                'pos_full_voltage': self.pos_full_var.get(),
                'neg_full_voltage': self.neg_full_var.get(),
                'lsb_per_div': self.lsb_per_div_var.get(),
                'ref_mode': self.ref_mode_var.get(),
                'yaxis_mode': self.yaxis_mode_var.get(),
                'yaxis_min': self.yaxis_min_var.get(),
                'yaxis_max': self.yaxis_max_var.get(),
                'skip_after_change': self.skip_after_change_var.get(),
                'skip_first_data': self.skip_first_data_var.get(),
                'skip_before_change': self.skip_before_change_var.get(),
                'temp_yaxis_mode': self.temp_yaxis_mode_var.get(),
                'csv_file_path': self.file_path_var.get(),
                'temp_csv_file_path': self.temp_file_path_var.get()
            }
            
            with open(self.SETTINGS_FILE, 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4)
            
        except Exception as e:
            print(f"設定の保存に失敗しました: {e}")
    
    def update_all_graphs(self):
        """開いている全てのグラフを更新"""
        # 設定値を取得
        try:
            bit_precision = int(self.bit_precision_var.get())
            pos_full = float(self.pos_full_var.get())
            neg_full = float(self.neg_full_var.get())
            lsb_per_div = float(self.lsb_per_div_var.get())
            ref_mode = self.ref_mode_var.get()
            yaxis_mode = self.yaxis_mode_var.get()
            yaxis_min = float(self.yaxis_min_var.get()) if yaxis_mode == "manual" else None
            yaxis_max = float(self.yaxis_max_var.get()) if yaxis_mode == "manual" else None
            skip_after_change = int(self.skip_after_change_var.get())
            skip_first_data = self.skip_first_data_var.get()
            skip_before_change = self.skip_before_change_var.get()
        except ValueError:
            return

        # 新しいプロッターを作成
        plotter = LSBGraphPlotter(bit_precision, pos_full, neg_full, lsb_per_div, ref_mode,
                                  yaxis_mode, yaxis_min, yaxis_max,
                                  skip_after_change, skip_first_data, skip_before_change)
        
        # 存在するウィンドウのみ更新
        valid_windows = []
        for window in self.graph_windows:
            try:
                if window.winfo_exists():
                    plotter.update_plot_window(window)
                    valid_windows.append(window)
            except:
                pass
        
        # 有効なウィンドウのリストを更新
        self.graph_windows = valid_windows
    
    def load_csv_file(self):
        """CSVファイルを選択して読み込み"""
        # ファイル選択ダイアログ
        filename = filedialog.askopenfilename(
            title="CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not filename:
            return

        self.file_path_var.set(filename)
        self._load_csv_from_path(filename, show_message=True)
        # ファイルパスを保存
        self.save_settings()

    def _load_csv_from_path(self, filename, show_message=True):
        """指定パスからCSVファイルを読み込み"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                self.csv_data = list(reader)
                headers = reader.fieldnames

            if not self.csv_data:
                if show_message:
                    messagebox.showerror("エラー", "CSVファイルにデータがありません")
                return

            # シリアルNo.を抽出
            self._extract_serial_numbers(headers)

            # チェックボックスを作成（デフォルト全選択）
            self._create_checkboxes()

            if show_message:
                messagebox.showinfo("成功", f"CSVファイルを読み込みました\n{len(self.csv_data)}行のデータ")

        except Exception as e:
            if show_message:
                messagebox.showerror("エラー", f"CSVファイルの読み込みに失敗しました:\n{str(e)}")
    
    def _extract_serial_numbers(self, headers):
        """ヘッダーからシリアルNo.を抽出"""
        self.serial_numbers = []
        for header in headers:
            if header.endswith('_POS') or header.endswith('_NEG'):
                parts = header.rsplit('_', 1)
                serial = parts[0]
                pole = parts[1]
                
                if serial not in [sn for sn, _ in self.serial_numbers]:
                    self.serial_numbers.append((serial, []))
                
                for i, (sn, poles) in enumerate(self.serial_numbers):
                    if sn == serial and pole not in poles:
                        self.serial_numbers[i] = (sn, poles + [pole])
    
    def _create_checkboxes(self):
        """データ選択用チェックボックスを作成（デフォルト全選択）"""
        # 既存のチェックボックスをクリア
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.checkboxes = {}
        
        # シリアルNo.ごとにチェックボックスを作成
        for serial, poles in self.serial_numbers:
            frame = ttk.Frame(self.scrollable_frame)
            frame.pack(fill=tk.X, pady=2)
            
            ttk.Label(frame, text=f"{serial}:", width=15).pack(side=tk.LEFT, padx=5)
            
            for pole in poles:
                var = tk.BooleanVar(value=True)  # デフォルト: 選択状態
                cb = ttk.Checkbutton(frame, text=pole, variable=var)
                cb.pack(side=tk.LEFT, padx=10)
                self.checkboxes[f"{serial}_{pole}"] = var
    
    def plot_selected_data(self):
        """選択されたデータをグラフ表示"""
        if not self.csv_data:
            messagebox.showerror("エラー", "CSVファイルを読み込んでください")
            return

        # 設定値を取得
        try:
            bit_precision = int(self.bit_precision_var.get())
            pos_full = float(self.pos_full_var.get())
            neg_full = float(self.neg_full_var.get())
            lsb_per_div = float(self.lsb_per_div_var.get())
            ref_mode = self.ref_mode_var.get()
            yaxis_mode = self.yaxis_mode_var.get()
            yaxis_min = float(self.yaxis_min_var.get()) if yaxis_mode == "manual" else None
            yaxis_max = float(self.yaxis_max_var.get()) if yaxis_mode == "manual" else None
            skip_after_change = int(self.skip_after_change_var.get())
            skip_first_data = self.skip_first_data_var.get()
            skip_before_change = self.skip_before_change_var.get()
        except ValueError:
            messagebox.showerror("エラー", "設定値が不正です")
            return

        # LSBGraphPlotterを作成
        plotter = LSBGraphPlotter(bit_precision, pos_full, neg_full, lsb_per_div, ref_mode,
                                  yaxis_mode, yaxis_min, yaxis_max,
                                  skip_after_change, skip_first_data, skip_before_change)
        
        # 選択されたデータをプロット
        plot_count = 0
        for key, var in self.checkboxes.items():
            if var.get():
                serial, pole = key.rsplit('_', 1)
                window = plotter.plot_csv_data(self, self.csv_data, serial, pole)
                if window:
                    self.graph_windows.append(window)
                    plot_count += 1
        
        if plot_count == 0:
            messagebox.showwarning("警告", "表示するデータを選択してください")

    def load_temp_csv_file(self):
        """温度CSVファイルを選択して読み込み"""
        filename = filedialog.askopenfilename(
            title="温度CSVファイルを選択",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )

        if not filename:
            return

        self.temp_file_path_var.set(filename)
        self._load_temp_csv_from_path(filename, show_message=True)
        self.save_settings()

    def _load_temp_csv_from_path(self, filename, show_message=True):
        """指定パスから温度CSVファイルを読み込み"""
        try:
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                self.temp_csv_data = list(reader)

            if not self.temp_csv_data:
                if show_message:
                    messagebox.showerror("エラー", "温度CSVファイルにデータがありません")
                return

            if show_message:
                messagebox.showinfo("成功", f"温度CSVを読み込みました\n{len(self.temp_csv_data)}行のデータ")

        except Exception as e:
            if show_message:
                messagebox.showerror("エラー", f"温度CSVの読み込みに失敗しました:\n{str(e)}")

    def plot_temperature_graph(self):
        """温特グラフを表示（2軸: LSB変動 + 温度差）"""
        if not self.csv_data:
            messagebox.showerror("エラー", "測定CSVファイルを読み込んでください")
            return

        if not self.temp_csv_data:
            messagebox.showerror("エラー", "温度CSVファイルを読み込んでください")
            return

        # 設定値を取得（温特グラフ用）
        try:
            bit_precision = int(self.bit_precision_var.get())
            pos_full = float(self.pos_full_var.get())
            neg_full = float(self.neg_full_var.get())
            ref_mode = "first_avg"  # 温特グラフは初回平均固定

            # Y軸設定: デフォルト or 共通設定
            if self.temp_yaxis_mode_var.get() == "default":
                # デフォルト: ±8LSB 2LSB/Div
                lsb_per_div = 2
                temp_yaxis_mode = "manual"
                temp_yaxis_min = -8.0
                temp_yaxis_max = 8.0
            else:
                # 共通設定を使用
                lsb_per_div = float(self.lsb_per_div_var.get())
                temp_yaxis_mode = self.yaxis_mode_var.get()
                if temp_yaxis_mode == "manual":
                    temp_yaxis_min = float(self.yaxis_min_var.get())
                    temp_yaxis_max = float(self.yaxis_max_var.get())
                else:
                    # オートの場合はNone
                    temp_yaxis_min = None
                    temp_yaxis_max = None

            skip_after_change = int(self.skip_after_change_var.get())
            skip_first_data = self.skip_first_data_var.get()
            skip_before_change = self.skip_before_change_var.get()
        except ValueError:
            messagebox.showerror("エラー", "設定値が不正です")
            return

        # LSBGraphPlotterを作成（温特グラフ用設定）
        plotter = LSBGraphPlotter(bit_precision, pos_full, neg_full, lsb_per_div, ref_mode,
                                  temp_yaxis_mode, temp_yaxis_min, temp_yaxis_max,
                                  skip_after_change, skip_first_data, skip_before_change)

        # 選択されたデータをプロット
        plot_count = 0
        for key, var in self.checkboxes.items():
            if var.get():
                serial, pole = key.rsplit('_', 1)
                result = plotter.plot_temperature_characteristic(
                    self.csv_data, self.temp_csv_data, serial, pole,
                    temp_yaxis_mode, temp_yaxis_min, temp_yaxis_max
                )
                if result:
                    plot_count += 1

        if plot_count == 0:
            messagebox.showwarning("警告", "表示するデータを選択してください")