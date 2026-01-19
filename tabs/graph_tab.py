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

        # 固定値・初期値の注記（Y軸関連を上に）
        ttk.Label(temp_settings_frame, text="【固定値・初期値】",
                  font=('', 9, 'bold')).pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(temp_settings_frame, text="・Y軸(LSB)初期値: ±8LSB, 2LSB/Div").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・Y軸(温度): ±8℃, 2℃/Div").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・基準電圧モード: 初回平均").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・X軸: 10分/Div").pack(anchor=tk.W, padx=15, pady=1)

        # 温特グラフで使用する設定（共通設定を参照）
        ttk.Separator(temp_settings_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=12)
        ttk.Label(temp_settings_frame, text="【共通設定を使用】",
                  font=('', 9, 'bold')).pack(anchor=tk.W, pady=(0, 8))
        ttk.Label(temp_settings_frame, text="・Bit精度").pack(anchor=tk.W, padx=15, pady=1)
        ttk.Label(temp_settings_frame, text="・スキップ設定").pack(anchor=tk.W, padx=15, pady=1)

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
        """温特グラフを表示（2軸: LSB変動 + 温度差）+ 設定画面"""
        if not self.csv_data:
            messagebox.showerror("エラー", "測定CSVファイルを読み込んでください")
            return

        if not self.temp_csv_data:
            messagebox.showerror("エラー", "温度CSVファイルを読み込んでください")
            return

        # 選択されたデータを確認
        selected_keys = [key for key, var in self.checkboxes.items() if var.get()]
        if not selected_keys:
            messagebox.showwarning("警告", "表示するデータを選択してください")
            return

        # 設定画面ウィンドウを作成
        self._create_temp_graph_settings_window(selected_keys)

    def _create_temp_graph_settings_window(self, selected_keys):
        """温特グラフ設定画面ウィンドウを作成"""
        # 既存のウィンドウがあれば閉じる
        if hasattr(self, 'temp_settings_window') and self.temp_settings_window:
            try:
                if self.temp_settings_window.winfo_exists():
                    self.temp_settings_window.destroy()
            except:
                pass

        # 設定画面ウィンドウ（温特グラフと被らないように右端に配置）
        self.temp_settings_window = tk.Toplevel(self)
        self.temp_settings_window.title("温特グラフ詳細設定")
        # 画面右端に配置
        screen_width = self.temp_settings_window.winfo_screenwidth()
        self.temp_settings_window.geometry(f"420x560+{screen_width - 450}+50")
        self.temp_settings_window.resizable(False, False)

        # 選択されたキーを保存
        self.temp_graph_selected_keys = selected_keys

        # Y軸設定用変数（デフォルト選択時用）
        self.temp_yaxis_select_var = tk.StringVar(value="default")

        main_frame = ttk.Frame(self.temp_settings_window, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 計算結果表示エリア（POS/NEG別）
        calc_frame = ttk.LabelFrame(main_frame, text="LSB電圧計算結果（実測値）", padding=10)
        calc_frame.pack(fill=tk.X, pady=(0, 10))

        # POS計算結果
        ttk.Label(calc_frame, text="【POS】", font=('', 9, 'bold')).pack(anchor=tk.W)
        self.calc_pos_fffff_label = ttk.Label(calc_frame, text="  +Full平均: ---")
        self.calc_pos_fffff_label.pack(anchor=tk.W, pady=1)
        self.calc_pos_zero_label = ttk.Label(calc_frame, text="  -Full平均: ---")
        self.calc_pos_zero_label.pack(anchor=tk.W, pady=1)
        self.calc_pos_lsb_label = ttk.Label(calc_frame, text="  LSB電圧: ---")
        self.calc_pos_lsb_label.pack(anchor=tk.W, pady=1)

        # NEG計算結果
        ttk.Label(calc_frame, text="【NEG】", font=('', 9, 'bold')).pack(anchor=tk.W, pady=(5, 0))
        self.calc_neg_fffff_label = ttk.Label(calc_frame, text="  +Full平均: ---")
        self.calc_neg_fffff_label.pack(anchor=tk.W, pady=1)
        self.calc_neg_zero_label = ttk.Label(calc_frame, text="  -Full平均: ---")
        self.calc_neg_zero_label.pack(anchor=tk.W, pady=1)
        self.calc_neg_lsb_label = ttk.Label(calc_frame, text="  LSB電圧: ---")
        self.calc_neg_lsb_label.pack(anchor=tk.W, pady=1)

        # Y軸(LSB)設定エリア
        yaxis_frame = ttk.LabelFrame(main_frame, text="Y軸(LSB)設定", padding=10)
        yaxis_frame.pack(fill=tk.X, pady=(0, 10))

        # Y軸範囲（1行目：デフォルト）
        yaxis_row0 = ttk.Frame(yaxis_frame)
        yaxis_row0.pack(fill=tk.X, pady=3)
        ttk.Label(yaxis_row0, text="Y軸範囲:", width=12).pack(side=tk.LEFT)
        ttk.Radiobutton(yaxis_row0, text="デフォルト(±8LSB 2LSB/Div)",
                        variable=self.temp_yaxis_select_var, value="default").pack(side=tk.LEFT)

        # Y軸範囲（2行目：オート）
        yaxis_row1 = ttk.Frame(yaxis_frame)
        yaxis_row1.pack(fill=tk.X, pady=3)
        ttk.Label(yaxis_row1, text="", width=12).pack(side=tk.LEFT)
        ttk.Radiobutton(yaxis_row1, text="オート", variable=self.temp_yaxis_select_var,
                        value="auto").pack(side=tk.LEFT)

        # Y軸範囲（3行目：設定値）
        yaxis_row2 = ttk.Frame(yaxis_frame)
        yaxis_row2.pack(fill=tk.X, pady=3)
        ttk.Label(yaxis_row2, text="", width=12).pack(side=tk.LEFT)
        ttk.Radiobutton(yaxis_row2, text="設定値", variable=self.temp_yaxis_select_var,
                        value="manual").pack(side=tk.LEFT)
        ttk.Label(yaxis_row2, text="Min:").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(yaxis_row2, textvariable=self.yaxis_min_var, width=7).pack(side=tk.LEFT)
        ttk.Label(yaxis_row2, text="Max:").pack(side=tk.LEFT, padx=(10, 2))
        ttk.Entry(yaxis_row2, textvariable=self.yaxis_max_var, width=7).pack(side=tk.LEFT)

        # 縦軸LSB/Div
        div_frame = ttk.Frame(yaxis_frame)
        div_frame.pack(fill=tk.X, pady=3)
        ttk.Label(div_frame, text="縦軸LSB/Div:", width=12).pack(side=tk.LEFT)
        ttk.Entry(div_frame, textvariable=self.lsb_per_div_var, width=10).pack(side=tk.LEFT, padx=5)

        # 縦軸更新ボタン（Y軸設定内）
        ttk.Button(yaxis_frame, text="縦軸更新",
                   command=self._apply_yaxis_to_temp_graph).pack(anchor=tk.E, pady=(8, 0))

        # PNG保存エリア
        save_frame = ttk.LabelFrame(main_frame, text="グラフ保存", padding=10)
        save_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(save_frame, text="グラフをPNG保存",
                   command=self._save_temp_graphs).pack(anchor=tk.W)

        # 閉じるボタン
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="閉じる",
                   command=self.temp_settings_window.destroy).pack(side=tk.RIGHT, padx=5)

        # 初回描画（デフォルト値で）
        self._draw_temp_graph_and_update_calc()

    def _apply_yaxis_to_temp_graph(self):
        """縦軸更新ボタン: 共通設定の値でグラフを再描画（位置保持）"""
        import matplotlib.pyplot as plt

        # 既存のグラフウィンドウの位置を保存
        graph_positions = []
        for fig_num in plt.get_fignums():
            try:
                fig = plt.figure(fig_num)
                manager = fig.canvas.manager
                x = manager.window.winfo_x()
                y = manager.window.winfo_y()
                graph_positions.append((x, y))
            except:
                pass

        plt.close('all')
        self._draw_temp_graph_and_update_calc()

        # 新しいグラフウィンドウを保存した位置に移動
        if graph_positions:
            for i, fig_num in enumerate(plt.get_fignums()):
                if i < len(graph_positions):
                    try:
                        fig = plt.figure(fig_num)
                        manager = fig.canvas.manager
                        x, y = graph_positions[i]
                        manager.window.geometry(f"+{x}+{y}")
                    except:
                        pass

    def _draw_temp_graph_and_update_calc(self):
        """温特グラフを描画し、計算結果を更新"""
        try:
            bit_precision = int(self.bit_precision_var.get())
            pos_full = float(self.pos_full_var.get())
            neg_full = float(self.neg_full_var.get())
            ref_mode = "first_avg"  # 温特グラフは初回平均固定

            # Y軸設定を選択に基づいて決定
            yaxis_select = self.temp_yaxis_select_var.get()
            if yaxis_select == "default":
                # デフォルト: ±8LSB, 2LSB/Div
                lsb_per_div = 2
                yaxis_mode = "manual"
                yaxis_min = -8.0
                yaxis_max = 8.0
            elif yaxis_select == "auto":
                # オート
                lsb_per_div = float(self.lsb_per_div_var.get())
                yaxis_mode = "auto"
                yaxis_min = None
                yaxis_max = None
            else:
                # 設定値
                lsb_per_div = float(self.lsb_per_div_var.get())
                yaxis_mode = "manual"
                yaxis_min = float(self.yaxis_min_var.get())
                yaxis_max = float(self.yaxis_max_var.get())

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

        # 選択されたデータをプロット（POS/NEG別に計算結果とfigureを保存）
        self.temp_graph_pos_info = None
        self.temp_graph_neg_info = None
        self.temp_graph_pos_serial = None
        self.temp_graph_neg_serial = None

        for key in self.temp_graph_selected_keys:
            serial, pole = key.rsplit('_', 1)
            calc_info = plotter.plot_temperature_characteristic(
                self.csv_data, self.temp_csv_data, serial, pole,
                yaxis_mode, yaxis_min, yaxis_max
            )
            if calc_info:
                if pole == "POS":
                    self.temp_graph_pos_info = calc_info
                    self.temp_graph_pos_serial = serial
                elif pole == "NEG":
                    self.temp_graph_neg_info = calc_info
                    self.temp_graph_neg_serial = serial

        pos_calc_info = self.temp_graph_pos_info
        neg_calc_info = self.temp_graph_neg_info

        # エラーチェック：必要なデータ（FFFFF/00000）があるか確認
        missing_data = []
        for calc_info, pole_name in [(pos_calc_info, "POS"), (neg_calc_info, "NEG")]:
            if calc_info:
                if not calc_info.get('has_fffff'):
                    missing_data.append(f"{pole_name}: +Full(FFFFF)データなし")
                if not calc_info.get('has_zero'):
                    missing_data.append(f"{pole_name}: -Full(00000)データなし")

        if missing_data:
            messagebox.showwarning("警告",
                "LSB電圧計算に必要なデータが不足しています:\n" + "\n".join(missing_data))

        # 計算結果を更新（POS/NEG別）
        self._update_calc_labels(pos_calc_info, neg_calc_info)

    def _update_calc_labels(self, pos_calc_info, neg_calc_info):
        """計算結果ラベルを更新（POS/NEG別、表示は小数第5位まで）"""
        # POS計算結果
        if pos_calc_info:
            fffff_avg = pos_calc_info.get('fffff_avg')
            zero_avg = pos_calc_info.get('00000_avg')
            lsb_voltage = pos_calc_info.get('lsb_voltage')

            if fffff_avg is not None:
                self.calc_pos_fffff_label.config(text=f"  +Full平均: {fffff_avg:.5f} V")
            else:
                self.calc_pos_fffff_label.config(text="  +Full平均: データなし")

            if zero_avg is not None:
                self.calc_pos_zero_label.config(text=f"  -Full平均: {zero_avg:.5f} V")
            else:
                self.calc_pos_zero_label.config(text="  -Full平均: データなし")

            if lsb_voltage is not None:
                lsb_mv = lsb_voltage * 1000  # V → mV
                self.calc_pos_lsb_label.config(text=f"  LSB電圧: {lsb_mv:.5f} mV")
            else:
                self.calc_pos_lsb_label.config(text="  LSB電圧: 計算不可")
        else:
            self.calc_pos_fffff_label.config(text="  +Full平均: ---")
            self.calc_pos_zero_label.config(text="  -Full平均: ---")
            self.calc_pos_lsb_label.config(text="  LSB電圧: ---")

        # NEG計算結果
        if neg_calc_info:
            fffff_avg = neg_calc_info.get('fffff_avg')
            zero_avg = neg_calc_info.get('00000_avg')
            lsb_voltage = neg_calc_info.get('lsb_voltage')

            if fffff_avg is not None:
                self.calc_neg_fffff_label.config(text=f"  +Full平均: {fffff_avg:.5f} V")
            else:
                self.calc_neg_fffff_label.config(text="  +Full平均: データなし")

            if zero_avg is not None:
                self.calc_neg_zero_label.config(text=f"  -Full平均: {zero_avg:.5f} V")
            else:
                self.calc_neg_zero_label.config(text="  -Full平均: データなし")

            if lsb_voltage is not None:
                lsb_mv = lsb_voltage * 1000  # V → mV
                self.calc_neg_lsb_label.config(text=f"  LSB電圧: {lsb_mv:.5f} mV")
            else:
                self.calc_neg_lsb_label.config(text="  LSB電圧: 計算不可")
        else:
            self.calc_neg_fffff_label.config(text="  +Full平均: ---")
            self.calc_neg_zero_label.config(text="  -Full平均: ---")
            self.calc_neg_lsb_label.config(text="  LSB電圧: ---")

    def _redraw_temp_graph(self):
        """温特グラフを再描画"""
        import matplotlib.pyplot as plt
        # 既存のグラフを閉じる
        plt.close('all')
        # 再描画
        self._draw_temp_graph_and_update_calc()

    def _save_temp_graphs(self):
        """温特グラフをPNG保存（POS/NEG両方）"""
        import os

        # 保存するグラフがあるか確認
        graphs_to_save = []
        if self.temp_graph_pos_info and self.temp_graph_pos_info.get('figure'):
            graphs_to_save.append(("POS", self.temp_graph_pos_info['figure'], self.temp_graph_pos_serial))
        if self.temp_graph_neg_info and self.temp_graph_neg_info.get('figure'):
            graphs_to_save.append(("NEG", self.temp_graph_neg_info['figure'], self.temp_graph_neg_serial))

        if not graphs_to_save:
            messagebox.showerror("エラー", "保存するグラフがありません")
            return

        # フォルダ選択ダイアログ
        folder_path = filedialog.askdirectory(title="保存先フォルダを選択")
        if not folder_path:
            return

        # 保存実行
        saved_files = []
        errors = []
        for pole, fig, serial in graphs_to_save:
            filename = f"温特グラフ_{serial}_{pole}.png"
            filepath = os.path.join(folder_path, filename)
            try:
                fig.savefig(filepath, dpi=150, bbox_inches='tight')
                saved_files.append(filename)
            except Exception as e:
                errors.append(f"{filename}: {str(e)}")

        # 結果表示
        if saved_files:
            msg = "保存しました:\n" + "\n".join(saved_files)
            if errors:
                msg += "\n\nエラー:\n" + "\n".join(errors)
                messagebox.showwarning("一部保存完了", msg)
            else:
                messagebox.showinfo("成功", msg)
        else:
            messagebox.showerror("エラー", "保存に失敗しました:\n" + "\n".join(errors))