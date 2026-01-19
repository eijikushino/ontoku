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
        # 画面右端に配置（高さを拡大）
        screen_width = self.temp_settings_window.winfo_screenwidth()
        self.temp_settings_window.geometry(f"420x720+{screen_width - 450}+50")
        self.temp_settings_window.resizable(False, False)

        # 選択されたキーを保存
        self.temp_graph_selected_keys = selected_keys

        # Y軸設定用変数（デフォルト選択時用）
        self.temp_yaxis_select_var = tk.StringVar(value="default")

        # NEG絶対値無効オプション
        self.neg_no_abs_var = tk.BooleanVar(value=False)

        # X軸全表示オプション（デフォルト: 25div=250分）
        self.xaxis_full_var = tk.BooleanVar(value=False)

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

        # NEG絶対値無効チェックボックス
        neg_abs_frame = ttk.Frame(calc_frame)
        neg_abs_frame.pack(anchor=tk.W, pady=(5, 0))
        self.neg_no_abs_check = ttk.Checkbutton(
            neg_abs_frame,
            text="NEG: 絶対値を使用しない（LSB電圧が負になる）",
            variable=self.neg_no_abs_var,
            command=self._on_neg_no_abs_changed
        )
        self.neg_no_abs_check.pack(anchor=tk.W, padx=10)

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

        # Y軸選択変更時に自動更新
        self.temp_yaxis_select_var.trace_add('write', lambda *args: self._redraw_temp_graph_preserve_position())

        # X軸設定エリア（別枠）
        xaxis_frame = ttk.LabelFrame(main_frame, text="X軸設定", padding=10)
        xaxis_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(xaxis_frame, text="全表示（デフォルト: 25div=250分）",
                        variable=self.xaxis_full_var,
                        command=self._on_xaxis_full_changed).pack(anchor=tk.W)

        # PNG保存エリア
        save_frame = ttk.LabelFrame(main_frame, text="グラフ保存", padding=10)
        save_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(save_frame, text="グラフをPNG保存",
                   command=self._save_temp_graphs).pack(anchor=tk.W)

        # 区間別平均電圧表示エリア
        analysis_frame = ttk.LabelFrame(main_frame, text="データ解析", padding=10)
        analysis_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Button(analysis_frame, text="温度係数表示",
                   command=self._show_section_averages).pack(anchor=tk.W)

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

        # NEG絶対値無効オプションを取得
        neg_no_abs = self.neg_no_abs_var.get()
        # X軸全表示オプションを取得
        xaxis_full = self.xaxis_full_var.get()

        for key in self.temp_graph_selected_keys:
            serial, pole = key.rsplit('_', 1)
            # NEGの場合のみneg_no_absを適用
            use_no_abs = neg_no_abs if pole == "NEG" else False
            calc_info = plotter.plot_temperature_characteristic(
                self.csv_data, self.temp_csv_data, serial, pole,
                yaxis_mode, yaxis_min, yaxis_max, use_no_abs, xaxis_full
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

    def _on_neg_no_abs_changed(self):
        """NEG絶対値無効チェックボックスの変更時: グラフを再描画（位置保持）"""
        self._redraw_temp_graph_preserve_position()

    def _on_xaxis_full_changed(self):
        """X軸全表示チェックボックスの変更時: グラフを再描画（位置保持）"""
        self._redraw_temp_graph_preserve_position()

    def _redraw_temp_graph_preserve_position(self):
        """温特グラフを位置を保持して再描画"""
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

    def _show_section_averages(self):
        """区間別平均電圧を表示するウィンドウを開く"""
        if not self.csv_data:
            messagebox.showerror("エラー", "測定CSVファイルを読み込んでください")
            return

        # 設定値を取得
        try:
            bit_precision = int(self.bit_precision_var.get())
            pos_full = float(self.pos_full_var.get())
            neg_full = float(self.neg_full_var.get())
            skip_after_change = int(self.skip_after_change_var.get())
            skip_first_data = self.skip_first_data_var.get()
            skip_before_change = self.skip_before_change_var.get()
        except ValueError:
            messagebox.showerror("エラー", "設定値が不正です")
            return

        # LSBGraphPlotterを作成
        plotter = LSBGraphPlotter(bit_precision, pos_full, neg_full, 2, "first_avg",
                                  "auto", None, None,
                                  skip_after_change, skip_first_data, skip_before_change)

        # 選択されたデータの区間平均を取得
        pos_sections = None
        neg_sections = None
        pos_serial = None
        neg_serial = None

        for key in self.temp_graph_selected_keys:
            serial, pole = key.rsplit('_', 1)
            column_name = f"{serial}_{pole}"
            sections = plotter.extract_section_averages(
                self.csv_data, serial, pole, column_name, last_minutes=10
            )
            if pole == "POS":
                pos_sections = sections
                pos_serial = serial
            elif pole == "NEG":
                neg_sections = sections
                neg_serial = serial

        # 結果表示ウィンドウを作成
        self._create_section_averages_window(pos_sections, neg_sections, pos_serial, neg_serial)

    def _create_section_averages_window(self, pos_sections, neg_sections, pos_serial, neg_serial):
        """温度係数表示ウィンドウを作成"""
        # 既存のウィンドウがあれば閉じる
        if hasattr(self, 'section_avg_window') and self.section_avg_window:
            try:
                if self.section_avg_window.winfo_exists():
                    self.section_avg_window.destroy()
            except:
                pass

        # ウィンドウ作成
        self.section_avg_window = tk.Toplevel(self)
        self.section_avg_window.title("温度係数")
        self.section_avg_window.geometry("750x700")
        self.section_avg_window.resizable(True, True)

        main_frame = ttk.Frame(self.section_avg_window, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # スペック値（内部で使用）
        self.temp_coef_spec_var = tk.StringVar(value="1.9")

        # テーブル表示エリア（スクロール対応）
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)

        # キャンバスとスクロールバー
        canvas = tk.Canvas(table_frame)
        v_scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=canvas.yview)
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        self.temp_coef_table_frame = ttk.Frame(canvas)

        self.temp_coef_table_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.temp_coef_table_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)

        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # テーブル作成
        self._create_temp_coef_table(pos_sections, neg_sections, pos_serial, neg_serial)

        # 閉じるボタン
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        ttk.Button(btn_frame, text="閉じる",
                   command=self.section_avg_window.destroy).pack(side=tk.RIGHT)

    def _create_temp_coef_table(self, pos_sections, neg_sections, pos_serial, neg_serial):
        """温度係数テーブルを作成（添付画像形式）"""
        parent = self.temp_coef_table_frame

        # スペック値を取得
        try:
            spec_value = float(self.temp_coef_spec_var.get())
        except ValueError:
            spec_value = 1.9

        # ヘッダー行
        headers = ['ユニットNo.', '入力コード', '温度\n(℃)', '測定電圧\n(V)', 'ΔT\n(℃)', 'ΔV\n(V)',
                   '温度係数\n(ppm/℃)', 'スペック\n(ppm/℃)', '判定']
        header_widths = [10, 7, 4, 10, 4, 10, 9, 7, 4]

        for col, (header, width) in enumerate(zip(headers, header_widths)):
            label = tk.Label(parent, text=header, relief=tk.RIDGE, width=width,
                           bg='#D0D0D0', font=('', 9))
            label.grid(row=0, column=col, sticky='nsew', padx=0, pady=0)

        # 斜線セル作成用ヘルパー関数
        def create_diagonal_cell(row, col, width, rowspan=1):
            """斜線付きの空セルを作成（左上から右下への対角線）"""
            # 文字幅をピクセルに変換（概算: 1文字 ≈ 8px）
            cell_width = width * 8
            cell_height = 15 * rowspan
            canvas = tk.Canvas(parent, width=cell_width, height=cell_height,
                             bg='white', highlightthickness=1, highlightbackground='gray')
            canvas.grid(row=row, column=col, rowspan=rowspan, sticky='nsew')
            # 左上から右下への斜線を描画
            canvas.create_line(0, 0, cell_width, cell_height, fill='gray')

        # データ行を作成
        row_idx = 1

        # 行高さを均一にするための最小高さ設定
        row_min_height = 15

        # POS/NEGそれぞれ処理
        for pole_idx, (sections, serial, pole) in enumerate([(pos_sections, pos_serial, "POS"),
                                                              (neg_sections, neg_serial, "NEG")]):
            if not sections:
                continue

            # コードごとにデータを整理（FFFFF, 80000, 00000）
            code_data = self._organize_sections_by_code(sections)

            unit_name = f"1PB397MK2\nDFH_{serial}\n{pole}"
            unit_row_start = row_idx

            # 23℃のフルスケール電圧範囲を計算
            fffff_23 = code_data.get('FFFFF', {}).get(23, {}).get('voltage')
            zero_23 = code_data.get('00000', {}).get(23, {}).get('voltage')
            full_scale_23 = (fffff_23 - zero_23) if (fffff_23 is not None and zero_23 is not None) else None

            for code_idx, code in enumerate(['FFFFF', '80000', '00000']):
                if code not in code_data:
                    continue

                temps_data = code_data[code]  # {28: voltage, 23: voltage, 18: voltage}
                code_row_start = row_idx

                # 温度データ（28, 23, 18の順）- 各温度2行で合計6行
                temp_order = [28, 23, 18]
                ref_voltage = temps_data.get(23, {}).get('voltage')  # 23℃が基準
                rows_per_temp = 2  # 各温度2行

                for temp_idx, temp in enumerate(temp_order):
                    if temp not in temps_data:
                        continue

                    voltage = temps_data[temp]['voltage']
                    temp_row = code_row_start + temp_idx * rows_per_temp

                    # 温度セル（2行結合）
                    tk.Label(parent, text=str(temp), relief=tk.RIDGE,
                            width=header_widths[2], font=('', 9), anchor='center', bg='white'
                            ).grid(row=temp_row, column=2, rowspan=rows_per_temp, sticky='nsew')

                    # 測定電圧セル（2行結合）
                    tk.Label(parent, text=f"{voltage:.6f}" if voltage else "",
                            relief=tk.RIDGE, width=header_widths[3],
                            font=('', 9), anchor='e', bg='white'
                            ).grid(row=temp_row, column=3, rowspan=rows_per_temp, sticky='nsew')

                row_idx = code_row_start + len(temp_order) * rows_per_temp

                # ΔT以降のセル - 28→23の計算（行1-2に跨る）
                if 28 in temps_data and 23 in temps_data:
                    v28 = temps_data[28]['voltage']
                    v23 = temps_data[23]['voltage']
                    delta_t_28 = 5.0
                    delta_v_28 = v28 - v23
                    temp_coef_28 = (delta_v_28 / full_scale_23) * 1e6 / delta_t_28 if full_scale_23 else 0
                    judgment_28 = "OK" if abs(temp_coef_28) <= spec_value else "NG"
                    bg_28 = '#FFCCCC' if judgment_28 == "NG" else 'white'

                    # 28℃の下半分と23℃の上半分に跨る（row 1-2）
                    delta_row = code_row_start + 1
                    tk.Label(parent, text=f"{delta_t_28:.1f}", relief=tk.RIDGE, width=header_widths[4],
                            font=('', 9), anchor='e', bg='white').grid(row=delta_row, column=4, rowspan=2, sticky='nsew')
                    tk.Label(parent, text=f"{delta_v_28:.6f}", relief=tk.RIDGE, width=header_widths[5],
                            font=('', 9), anchor='e', bg='white').grid(row=delta_row, column=5, rowspan=2, sticky='nsew')
                    tk.Label(parent, text=f"{temp_coef_28:.1f}", relief=tk.RIDGE, width=header_widths[6],
                            font=('', 9), anchor='e', bg=bg_28).grid(row=delta_row, column=6, rowspan=2, sticky='nsew')
                    tk.Label(parent, text=judgment_28, relief=tk.RIDGE, width=header_widths[8],
                            font=('', 9), anchor='center', bg=bg_28).grid(row=delta_row, column=8, rowspan=2, sticky='nsew')

                # ΔT以降のセル - 23→18の計算（行3-4に跨る）
                if 23 in temps_data and 18 in temps_data:
                    v23 = temps_data[23]['voltage']
                    v18 = temps_data[18]['voltage']
                    delta_t_18 = -5.0
                    delta_v_18 = v18 - v23
                    temp_coef_18 = (delta_v_18 / full_scale_23) * 1e6 / delta_t_18 if full_scale_23 else 0
                    judgment_18 = "OK" if abs(temp_coef_18) <= spec_value else "NG"
                    bg_18 = '#FFCCCC' if judgment_18 == "NG" else 'white'

                    # 23℃の下半分と18℃の上半分に跨る（row 3-4）
                    delta_row = code_row_start + 3
                    tk.Label(parent, text=f"{delta_t_18:.1f}", relief=tk.RIDGE, width=header_widths[4],
                            font=('', 9), anchor='e', bg='white').grid(row=delta_row, column=4, rowspan=2, sticky='nsew')
                    tk.Label(parent, text=f"{delta_v_18:.6f}", relief=tk.RIDGE, width=header_widths[5],
                            font=('', 9), anchor='e', bg='white').grid(row=delta_row, column=5, rowspan=2, sticky='nsew')
                    tk.Label(parent, text=f"{temp_coef_18:.1f}", relief=tk.RIDGE, width=header_widths[6],
                            font=('', 9), anchor='e', bg=bg_18).grid(row=delta_row, column=6, rowspan=2, sticky='nsew')
                    tk.Label(parent, text=judgment_18, relief=tk.RIDGE, width=header_widths[8],
                            font=('', 9), anchor='center', bg=bg_18).grid(row=delta_row, column=8, rowspan=2, sticky='nsew')

                # 空セル（ΔT列の上端と下端）- 斜線付き
                for col in [4, 5, 6, 8]:
                    create_diagonal_cell(code_row_start, col, header_widths[col])
                    create_diagonal_cell(code_row_start + 5, col, header_widths[col])

                # 入力コードのセル結合（縦6行）
                tk.Label(parent, text=f"{code} H", relief=tk.RIDGE, width=header_widths[1],
                        font=('', 9), bg='white').grid(row=code_row_start, column=1, rowspan=6, sticky='nsew')

                # 各行の高さを均一に設定
                for r in range(6):
                    parent.grid_rowconfigure(code_row_start + r, minsize=row_min_height)

            # 各コード6行 × 3コード = 18行
            total_rows = 6 * 3  # FFFFF, 80000, 00000
            row_idx = unit_row_start + total_rows

            # ユニットNo.のセル結合（縦）
            tk.Label(parent, text=unit_name, relief=tk.RIDGE, width=header_widths[0],
                    font=('', 9), bg='white').grid(row=unit_row_start, column=0, rowspan=total_rows, sticky='nsew')

        # スペックのセル結合（POS/NEG全体で共通）
        total_data_rows = row_idx - 1  # ヘッダー行を除いた全データ行数
        if total_data_rows > 0:
            tk.Label(parent, text=f"{spec_value:.1f}", relief=tk.RIDGE, width=header_widths[7],
                    font=('', 9), anchor='center', bg='white').grid(row=1, column=7, rowspan=total_data_rows, sticky='nsew')

    def _organize_sections_by_code(self, sections):
        """区間データをコードと温度で整理"""
        # 温特パターン: FFFFF→00000→80000 を繰り返し
        # 温度順: 23℃(1), 28℃, 18℃, 23℃(2), 23℃戻し
        temp_mapping = {
            0: 23,   # 23℃(1) - 基準値として使用
            1: 28,   # 28℃
            2: 18,   # 18℃
            3: 23,   # 23℃(2)
            4: 23,   # 23℃戻し - 参考値
        }

        code_data = {}  # {code: {temp: {'voltage': v, 'temp_set': idx}}}
        codes_per_temp = 3

        for i, section in enumerate(sections):
            temp_set_idx = i // codes_per_temp
            code = section['code']

            if temp_set_idx not in temp_mapping:
                continue

            temp = temp_mapping[temp_set_idx]

            if code not in code_data:
                code_data[code] = {}

            # 23℃は最初の値（temp_set_idx=0）を基準として使用
            # 28℃、18℃はそれぞれ1回のみ
            if temp == 23:
                # 23℃(1)のデータを基準として使用（temp_set_idx=0）
                if temp_set_idx == 0:
                    code_data[code][temp] = {
                        'voltage': section['avg_voltage'],
                        'temp_set': temp_set_idx
                    }
            else:
                code_data[code][temp] = {
                    'voltage': section['avg_voltage'],
                    'temp_set': temp_set_idx
                }

        return code_data

    def _format_minutes(self, minutes):
        """分を時:分形式にフォーマット"""
        hours = int(minutes // 60)
        mins = int(minutes % 60)
        if hours > 0:
            return f"{hours}h{mins:02d}m"
        else:
            return f"{mins}m"