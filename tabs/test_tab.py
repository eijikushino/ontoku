import tkinter as tk
import time
import csv
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

class TestTab(ttk.Frame):
    def __init__(self, parent, serial_manager):
        super().__init__(parent)
        self.serial_mgr = serial_manager
        
        # パターンデータを保持するリスト
        self.patterns = []
        for i in range(15):
            self.patterns.append({
                'enabled': tk.BooleanVar(value=False),
                'dataset': tk.StringVar(value='Position'),
                'pole': tk.StringVar(value='Pos'),
                'code': tk.StringVar(value='Center'),
                'manual_value': tk.StringVar(value=''),  # Manual入力用
                'time': tk.DoubleVar(value=1.0)  # 小数点1位まで対応
            })
        
        self.is_running = False
        self.is_holding = False  # ホールド(一時停止)フラグ
        self.skip_requested = False  # スキップフラグ
        self.def_vars = []  # DAC操作タブから共有するDEF選択状態
        self.save_folder = tk.StringVar(value="")  # 保存フォルダパス
        
        # ★★★ 計測ウィンドウ管理用 ★★★
        self.measurement_window = None  # 計測ウィンドウのインスタンス
        
        # スキャナーチャンネル設定用(変更: Pos/Negそれぞれに対応)
        self.scanner_channels_pos = []  # Pos用
        self.scanner_channels_neg = []  # Neg用

        # 選択していないDataSetにCenterコードを送信するオプション（デフォルト: True）
        self.send_opposite_center = tk.BooleanVar(value=True)
        
        # 実行状況管理用
        self.current_pattern_index = -1  # -1で初期化（未実行状態）
        self.pattern_start_time = 0
        self.total_start_time = 0
        self.total_patterns_time = 0
        self.current_pattern_time = 0
        
        # 設定ファイルから前回値を読み込み
        self.load_settings()
        
        self.create_widgets()
        
    def set_def_vars(self, def_vars):
        """DAC操作タブからDEF選択状態を共有"""
        self.def_vars = def_vars
    
    def create_widgets(self):
        # メインコンテナ
        main_container = ttk.Frame(self)
        main_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # ===== 左側全体のコンテナ =====
        left_container = ttk.Frame(main_container, width=520)  # 横幅を拡張
        left_container.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 10))
        left_container.pack_propagate(False)  # 幅を固定
        
        # ----- パターン設定テーブルエリア -----
        table_area = ttk.Frame(left_container)
        table_area.pack(fill=tk.BOTH, expand=True)
        
        # ヘッダー
        header_frame = ttk.Frame(table_area)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        # チェックボックス列用の空白(チェックボックスと同じpadx)
        ttk.Label(header_frame, text="", width=2).pack(side=tk.LEFT, padx=(5, 5))
        
        # No列(行データと同じpadx)
        header_no = ttk.Label(header_frame, text="No", width=3, anchor=tk.CENTER)
        header_no.pack(side=tk.LEFT, padx=(0, 2))
        
        # DataSet列(行データと同じpadx)
        header_dataset = ttk.Label(header_frame, text="DataSet", width=14, anchor=tk.CENTER)
        header_dataset.pack(side=tk.LEFT, padx=2)
        
        # Pole列(行データと同じpadx)
        header_pole = ttk.Label(header_frame, text="Pole", width=8, anchor=tk.CENTER)
        header_pole.pack(side=tk.LEFT, padx=2)
        
        # Code列(行データと同じpadx)
        header_code = ttk.Label(header_frame, text="Code", width=18, anchor=tk.CENTER)
        header_code.pack(side=tk.LEFT, padx=2)
        
        # Manual列(新規追加)
        header_manual = ttk.Label(header_frame, text="Manual(HEX)", width=12, anchor=tk.CENTER)
        header_manual.pack(side=tk.LEFT, padx=2)
        
        # Time列(行データと同じpadx)
        header_time = ttk.Label(header_frame, text="Time", width=10, anchor=tk.CENTER)
        header_time.pack(side=tk.LEFT, padx=2)
        
        # スクロール可能なテーブルフレーム
        table_canvas = tk.Canvas(table_area, height=300)
        scrollbar = ttk.Scrollbar(table_area, orient=tk.VERTICAL, command=table_canvas.yview)
        table_frame = ttk.Frame(table_canvas)
        
        table_canvas.create_window((0, 0), window=table_frame, anchor=tk.NW)
        table_canvas.configure(yscrollcommand=scrollbar.set)
        
        table_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 15行のパターン入力行を作成
        self.pattern_widgets = []
        for i in range(15):
            row_widgets = self.create_pattern_row(table_frame, i)
            self.pattern_widgets.append(row_widgets)
        
        table_frame.update_idletasks()
        table_canvas.configure(scrollregion=table_canvas.bbox("all"))
        
        # ----- 実行状況表示エリアとDEF選択エリア(テーブルの下) -----
        status_def_container = ttk.Frame(left_container)
        status_def_container.pack(fill=tk.X, pady=(10, 0))
        
        # 実行状況エリア(左側・最小サイズ、左詰め)
        status_frame = ttk.LabelFrame(status_def_container, text="実行状況", padding=5)
        status_frame.pack(side=tk.LEFT, padx=(0, 10))
        
        # 現在の実行番号
        current_pattern_frame = ttk.Frame(status_frame)
        current_pattern_frame.pack(anchor=tk.W, pady=1)
        ttk.Label(current_pattern_frame, text="実行中:", width=7).pack(side=tk.LEFT)
        self.current_pattern_label = ttk.Label(current_pattern_frame, text="未実行", 
                                                font=("", 9, "bold"), foreground="blue", width=8, anchor=tk.W)
        self.current_pattern_label.pack(side=tk.LEFT)
        
        # 現在のパターンの経過時間
        pattern_time_frame = ttk.Frame(status_frame)
        pattern_time_frame.pack(anchor=tk.W, pady=1)
        ttk.Label(pattern_time_frame, text="パターン:", width=7).pack(side=tk.LEFT)
        self.pattern_time_label = ttk.Label(pattern_time_frame, text="00:00/00:00", 
                                             font=("", 9), width=11, anchor=tk.W)
        self.pattern_time_label.pack(side=tk.LEFT)
        
        # トータル経過時間
        total_time_frame = ttk.Frame(status_frame)
        total_time_frame.pack(anchor=tk.W, pady=1)
        ttk.Label(total_time_frame, text="トータル:", width=7).pack(side=tk.LEFT)
        self.total_time_label = ttk.Label(total_time_frame, text="00:00/00:00", 
                                           font=("", 9), width=11, anchor=tk.W)
        self.total_time_label.pack(side=tk.LEFT)
        
        # DEF選択エリア(右側)
        def_frame = ttk.LabelFrame(status_def_container, text="DEF選択 & スキャナーチャンネル", padding=8)
        def_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # ヘッダー行
        header_frame = ttk.Frame(def_frame)
        header_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(header_frame, text="DEF", width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="Pos", width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="Neg", width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="", width=2).pack(side=tk.LEFT)  # スペーサー
        ttk.Label(header_frame, text="DEF", width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="Pos", width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        ttk.Label(header_frame, text="Neg", width=8, anchor=tk.CENTER).pack(side=tk.LEFT, padx=2)
        
        # DEFチェックボックスとスキャナーチャンネルを2列×3行で配置
        self.def_check_vars = [tk.BooleanVar(value=(i == 0)) for i in range(6)]  # 既定はDEF0のみON
        self.scanner_channels_pos = [tk.StringVar(value='ー') for i in range(6)]  # Pos用
        self.scanner_channels_neg = [tk.StringVar(value='ー') for i in range(6)]  # Neg用
        
        # 前回値を復元（DEFチェック状態も復元）
        if hasattr(self, 'saved_def_checks'):
            for i in range(min(6, len(self.saved_def_checks))):
                self.def_check_vars[i].set(self.saved_def_checks[i])
        if hasattr(self, 'saved_scanner_channels_pos'):
            for i in range(min(6, len(self.saved_scanner_channels_pos))):
                self.scanner_channels_pos[i].set(self.saved_scanner_channels_pos[i])
        if hasattr(self, 'saved_scanner_channels_neg'):
            for i in range(min(6, len(self.saved_scanner_channels_neg))):
                self.scanner_channels_neg[i].set(self.saved_scanner_channels_neg[i])
        
        # データ行のコンテナ
        data_container = ttk.Frame(def_frame)
        data_container.pack(fill=tk.BOTH, expand=True)
        
        # チャンネル選択肢(ー + CH00～CH09)
        channel_options = ['ー'] + [f'CH{i:02d}' for i in range(10)]
        
        # 1列目: DEF0, DEF1, DEF2
        left_col = ttk.Frame(data_container)
        left_col.pack(side=tk.LEFT, padx=(0, 10))
        for i in range(3):
            row = ttk.Frame(left_col)
            row.pack(fill=tk.X, pady=2)
            
            # チェックボックス（変更時に保存）
            cb = ttk.Checkbutton(row, text=f"DEF{i}", variable=self.def_check_vars[i], width=5,
                                 command=self.save_settings)
            cb.pack(side=tk.LEFT, padx=2)
            
            # Pos用チャンネル
            pos_combo = ttk.Combobox(row, textvariable=self.scanner_channels_pos[i], 
                                      values=channel_options, width=5, state='readonly')
            pos_combo.pack(side=tk.LEFT, padx=2)
            pos_combo.bind('<<ComboboxSelected>>', lambda e: self.save_settings())
            
            # Neg用チャンネル
            neg_combo = ttk.Combobox(row, textvariable=self.scanner_channels_neg[i], 
                                      values=channel_options, width=5, state='readonly')
            neg_combo.pack(side=tk.LEFT, padx=2)
            neg_combo.bind('<<ComboboxSelected>>', lambda e: self.save_settings())
        
        # 2列目: DEF3, DEF4, DEF5
        right_col = ttk.Frame(data_container)
        right_col.pack(side=tk.LEFT)
        for i in range(3, 6):
            row = ttk.Frame(right_col)
            row.pack(fill=tk.X, pady=2)
            
            # チェックボックス（変更時に保存）
            cb = ttk.Checkbutton(row, text=f"DEF{i}", variable=self.def_check_vars[i], width=5,
                                 command=self.save_settings)
            cb.pack(side=tk.LEFT, padx=2)
            
            # Pos用チャンネル
            pos_combo = ttk.Combobox(row, textvariable=self.scanner_channels_pos[i], 
                                      values=channel_options, width=5, state='readonly')
            pos_combo.pack(side=tk.LEFT, padx=2)
            pos_combo.bind('<<ComboboxSelected>>', lambda e: self.save_settings())
            
            # Neg用チャンネル
            neg_combo = ttk.Combobox(row, textvariable=self.scanner_channels_neg[i], 
                                      values=channel_options, width=5, state='readonly')
            neg_combo.pack(side=tk.LEFT, padx=2)
            neg_combo.bind('<<ComboboxSelected>>', lambda e: self.save_settings())
        
        # ===== 右側:制御パネルとログ =====
        right_frame = ttk.Frame(main_container)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 保存/読込フレーム
        file_frame = ttk.LabelFrame(right_frame, text="パターン保存/読込", padding=10)
        file_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 保存フォルダ選択
        folder_frame = ttk.Frame(file_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(folder_frame, text="保存先:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Entry(folder_frame, textvariable=self.save_folder, width=20).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(folder_frame, text="参照", command=self.select_folder, width=6).pack(side=tk.LEFT)
        
        # ファイル名入力
        name_frame = ttk.Frame(file_frame)
        name_frame.pack(fill=tk.X, pady=(0, 5))
        ttk.Label(name_frame, text="ファイル名:").pack(side=tk.LEFT, padx=(0, 5))
        self.filename_entry = ttk.Entry(name_frame, width=20)
        self.filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.filename_entry.insert(0, self.default_filename)  # 前回値を設定
        
        # 保存/読込ボタン
        button_frame = ttk.Frame(file_frame)
        button_frame.pack(fill=tk.X)
        ttk.Button(button_frame, text="保存", command=self.save_pattern, width=12).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="読込", command=self.load_pattern, width=12).pack(side=tk.LEFT)
        
        # 制御ボタン
        control_frame = ttk.LabelFrame(right_frame, text="実行制御", padding=10)
        control_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 開始、ホールド、スキップ、停止を横並び
        button_row = ttk.Frame(control_frame)
        button_row.pack(fill=tk.X, pady=2)
        
        self.start_button = ttk.Button(button_row, text="開始", 
                                        command=self.start_test, width=10)
        self.start_button.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        self.hold_button = ttk.Button(button_row, text="ホールド", 
                                        command=self.hold_test, width=10, state=tk.DISABLED)
        self.hold_button.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        self.skip_button = ttk.Button(button_row, text="スキップ", 
                                       command=self.skip_pattern, width=10, state=tk.DISABLED)
        self.skip_button.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        self.stop_button = ttk.Button(button_row, text="停止", 
                                       command=self.stop_test, width=10, state=tk.DISABLED)
        self.stop_button.pack(side=tk.LEFT, padx=2, expand=True, fill=tk.X)
        
        # ★★★ ここが修正箇所 ★★★
        # 全選択/全解除/計測ボタン
        select_frame = ttk.Frame(control_frame)
        select_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(select_frame, text="全選択", 
                   command=self.select_all, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(select_frame, text="全解除", 
                   command=self.deselect_all, width=10).pack(side=tk.LEFT, padx=2)
        
        # ★★★ 計測ボタンの参照を保持 ★★★
        self.measurement_button = ttk.Button(select_frame, text="計測",
                   command=self.open_measurement_window, width=10)
        self.measurement_button.pack(side=tk.LEFT, padx=2)
        # ★★★ 修正箇所ここまで ★★★

        # 選択していないDataSetにCenterコードを送信するオプション
        opposite_center_frame = ttk.Frame(control_frame)
        opposite_center_frame.pack(fill=tk.X, pady=5)
        opposite_center_cb = ttk.Checkbutton(
            opposite_center_frame,
            text="未選択DataSetにCenter送信",
            variable=self.send_opposite_center,
            command=self.save_settings
        )
        opposite_center_cb.pack(side=tk.LEFT, padx=2)
        
        # ログ表示
        log_frame = ttk.LabelFrame(right_frame, text="実行ログ", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = ScrolledText(log_frame, height=25, width=35,
                                     wrap=tk.WORD, state=tk.DISABLED)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # ログのタグ設定
        self.log_text.tag_config('INFO', foreground='black')
        self.log_text.tag_config('SUCCESS', foreground='green')
        self.log_text.tag_config('ERROR', foreground='red')
        self.log_text.tag_config('WARNING', foreground='orange')
        
    def create_pattern_row(self, parent, index):
        """パターン入力行を作成"""
        row_frame = ttk.Frame(parent)
        row_frame.pack(fill=tk.X, pady=2)
        
        pattern = self.patterns[index]
        
        # チェックボックス(直接配置)
        cb = ttk.Checkbutton(row_frame, variable=pattern['enabled'])
        cb.pack(side=tk.LEFT, padx=(5, 5))
        
        # 番号ラベル
        no_label = ttk.Label(row_frame, text=f"{index + 1}", width=4, anchor=tk.CENTER)
        no_label.pack(side=tk.LEFT, padx=(0, 2))
        
        # DataSet コンボボックス
        dataset_combo = ttk.Combobox(row_frame, textvariable=pattern['dataset'], 
                                      values=['Position', 'LBC'], width=10, state='readonly')
        dataset_combo.pack(side=tk.LEFT, padx=2)
        
        # Pole コンボボックス
        pole_combo = ttk.Combobox(row_frame, textvariable=pattern['pole'], 
                                   values=['Pos', 'Neg'], width=8, state='readonly')
        pole_combo.pack(side=tk.LEFT, padx=2)
        
        # Code コンボボックス（Manualを追加）
        code_combo = ttk.Combobox(row_frame, textvariable=pattern['code'], 
                                   values=['+Full', 'Center', '-Full', 'Manual'], width=10, state='readonly')
        code_combo.pack(side=tk.LEFT, padx=2)
        
        # Manual値入力欄（新規追加）
        manual_entry = ttk.Entry(row_frame, textvariable=pattern['manual_value'], width=12)
        manual_entry.pack(side=tk.LEFT, padx=2)
        
        # Time 入力 + minラベル
        time_frame = ttk.Frame(row_frame)
        time_frame.pack(side=tk.LEFT, padx=2)
        
        time_spinbox = ttk.Spinbox(time_frame, from_=0.1, to=999.9,
                                    increment=0.1, format='%.1f',
                                    textvariable=pattern['time'], width=8)
        time_spinbox.pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Label(time_frame, text="min", width=4).pack(side=tk.LEFT)
        
        return {
            'frame': row_frame,
            'checkbox': cb,
            'dataset': dataset_combo,
            'pole': pole_combo,
            'code': code_combo,
            'manual': manual_entry,
            'time': time_spinbox
        }
    
    def select_folder(self):
        """保存先フォルダを選択"""
        folder = filedialog.askdirectory(title="保存先フォルダを選択")
        if folder:
            self.save_folder.set(folder)
            self.save_settings()  # 設定を保存
    
    def save_pattern(self):
        """パターンをCSVファイルに保存"""
        if not self.save_folder.get():
            messagebox.showwarning("警告", "保存先フォルダを選択してください")
            return
        
        filename = self.filename_entry.get().strip()
        if not filename:
            messagebox.showwarning("警告", "ファイル名を入力してください")
            return
        
        # 拡張子がなければ.csvを追加
        if not filename.endswith('.csv'):
            filename += '.csv'
        
        filepath = f"{self.save_folder.get()}/{filename}"
        
        try:
            with open(filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                # ヘッダー行（Manual列を追加）
                writer.writerow(['No', 'Enabled', 'DataSet', 'Pole', 'Code', 'ManualValue', 'Time(min)'])
                
                # データ行(チェック状態も保存)
                for i, pattern in enumerate(self.patterns):
                    writer.writerow([
                        i + 1,
                        '1' if pattern['enabled'].get() else '0',  # 1=チェック有, 0=チェック無
                        pattern['dataset'].get(),
                        pattern['pole'].get(),
                        pattern['code'].get(),
                        pattern['manual_value'].get(),
                        pattern['time'].get()
                    ])
            
            # 設定を保存
            self.save_settings()
            
            messagebox.showinfo("保存完了", f"パターンを保存しました:\n{filepath}")
            self.log_message(f"パターンを保存: {filepath}", "SUCCESS")
        
        except Exception as e:
            messagebox.showerror("保存エラー", f"保存に失敗しました:\n{str(e)}")
            self.log_message(f"保存エラー: {str(e)}", "ERROR")
    
    def load_pattern(self):
        """CSVファイルからパターンを読込"""
        filepath = filedialog.askopenfilename(
            title="パターンファイルを選択",
            filetypes=[("CSVファイル", "*.csv"), ("すべてのファイル", "*.*")],
            initialdir=self.save_folder.get() if self.save_folder.get() else None
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                header = next(reader)  # ヘッダー行をスキップ
                
                for i, row in enumerate(reader):
                    if i >= 15:  # 15行まで
                        break
                    
                    if len(row) >= 7:  # Manual列を含む新形式
                        # データを読み込んで設定(チェック状態も復元)
                        # Enabled列は '1', '0' または 'Yes', 'No' の両方に対応
                        enabled_value = row[1].strip()
                        is_enabled = enabled_value in ['1', 'Yes', 'yes', 'YES', 'true', 'True', 'TRUE']
                        self.patterns[i]['enabled'].set(is_enabled)
                        
                        self.patterns[i]['dataset'].set(row[2])
                        self.patterns[i]['pole'].set(row[3])
                        self.patterns[i]['code'].set(row[4])
                        self.patterns[i]['manual_value'].set(row[5])
                        try:
                            self.patterns[i]['time'].set(float(row[6]))
                        except:
                            self.patterns[i]['time'].set(1.0)
                    elif len(row) >= 6:  # 旧形式（Manual列なし）
                        enabled_value = row[1].strip()
                        is_enabled = enabled_value in ['1', 'Yes', 'yes', 'YES', 'true', 'True', 'TRUE']
                        self.patterns[i]['enabled'].set(is_enabled)

                        self.patterns[i]['dataset'].set(row[2])
                        self.patterns[i]['pole'].set(row[3])
                        self.patterns[i]['code'].set(row[4])
                        self.patterns[i]['manual_value'].set('')  # 空にする
                        try:
                            self.patterns[i]['time'].set(float(row[5]))
                        except:
                            self.patterns[i]['time'].set(1.0)
            
            # 読み込んだファイルのパスとファイル名を設定に反映
            import os
            folder_path = os.path.dirname(filepath)
            filename = os.path.basename(filepath)
            
            # 拡張子を除いたファイル名を取得
            if filename.endswith('.csv'):
                filename_without_ext = filename[:-4]
            else:
                filename_without_ext = filename
            
            # UIを更新
            self.save_folder.set(folder_path)
            self.filename_entry.delete(0, tk.END)
            self.filename_entry.insert(0, filename_without_ext)
            
            # 設定を保存
            self.save_settings()
            
            messagebox.showinfo("読込完了", f"パターンを読み込みました:\n{filepath}")
            self.log_message(f"パターンを読込: {filepath}", "SUCCESS")
        
        except Exception as e:
            messagebox.showerror("読込エラー", f"読込に失敗しました:\n{str(e)}")
            self.log_message(f"読込エラー: {str(e)}", "ERROR")
    
    def select_all(self):
        """全パターンを選択"""
        for pattern in self.patterns:
            pattern['enabled'].set(True)
        self.log_message("全パターンを選択しました", "INFO")
    
    def deselect_all(self):
        """全パターンの選択を解除"""
        for pattern in self.patterns:
            pattern['enabled'].set(False)
        self.log_message("全パターンの選択を解除しました", "INFO")
    
    def start_test(self):
        """パターンテストを開始"""
        # シリアル通信の接続確認
        if not self.serial_mgr.is_connected():
            self.log_message("シリアルポートが接続されていません", "ERROR")
            return
        
        # DEF選択状態の確認(変更部分)
        selected_defs = [i for i, var in enumerate(self.def_check_vars) if var.get()]
        if not selected_defs:
            self.log_message("DEFが選択されていません", "ERROR")
            return
        
        # 選択されているパターンを収集
        enabled_patterns = []
        for i, pattern in enumerate(self.patterns):
            if pattern['enabled'].get():
                enabled_patterns.append({
                    'index': i + 1,
                    'dataset': pattern['dataset'].get(),
                    'pole': pattern['pole'].get(),
                    'code': pattern['code'].get(),
                    'manual_value': pattern['manual_value'].get(),
                    'time': pattern['time'].get()
                })
        
        if not enabled_patterns:
            self.log_message("実行するパターンが選択されていません", "ERROR")
            return
        
        # ★★★ パターン開始時に current_pattern_index を初期化 ★★★
        self.current_pattern_index = 0
        
        # 開始時に実行状況と経過時間をリセット
        self.current_pattern_label.config(text="未実行")
        self.pattern_time_label.config(text="00:00 / 00:00")
        self.total_time_label.config(text="00:00 / 00:00")
        
        self.is_running = True
        self.is_holding = False
        self.skip_requested = False
        self.start_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.skip_button.config(state=tk.NORMAL)
        self.hold_button.config(state=tk.NORMAL, text="ホールド")
        
        # トータル時間を計算
        self.total_patterns_time = sum(p['time'] for p in enabled_patterns) * 60  # 分を秒に
        self.total_start_time = time.time() + 3600  # 未来の時刻で初期化(異常値表示防止)
        
        self.log_message("=" * 40, "INFO")
        self.log_message(f"パターンテスト開始: {len(enabled_patterns)}個のパターン", "INFO")
        self.log_message(f"対象DEF: {selected_defs}", "INFO")
        self.log_message("=" * 40, "INFO")
        
        # 時間表示更新を開始(100msごと)
        self.update_time_display()
        
        # パターンを順次実行
        self.execute_patterns(enabled_patterns, 0)
    
    def execute_patterns(self, patterns, current_index):
        """パターンを順次実行(再帰的)"""
        if not self.is_running or current_index >= len(patterns):
            self.finish_test()
            return
        
        pattern = patterns[current_index]
        
        # ★★★ 現在のパターンインデックスを更新（0始まり） ★★★
        self.current_pattern_index = pattern['index'] - 1  # pattern['index']は1始まりなので-1
        
        # 実行状況を更新
        self.update_status_display(current_index, patterns)
        
        # パターン経過時間を0:00で表示(コマンド送信前)
        pattern_total = int(self.current_pattern_time)
        pattern_total_min = pattern_total // 60
        pattern_total_sec = pattern_total % 60
        self.pattern_time_label.config(
            text=f"00:00 / {pattern_total_min:02d}:{pattern_total_sec:02d}"
        )
        self.update_idletasks()  # 即座に画面更新
        
        # コマンド送信中は時間を進めないため、未来の時刻を設定
        self.pattern_start_time = time.time() + 3600  # 1時間後の未来時刻
        
        self.log_message(f"\nパターン {pattern['index']} 実行中...", "INFO")
        self.log_message(f"  DataSet: {pattern['dataset']}", "INFO")
        self.log_message(f"  Pole: {pattern['pole']}", "INFO")
        self.log_message(f"  Code: {pattern['code']}", "INFO")
        if pattern['code'] == 'Manual':
            self.log_message(f"  Manual値: {pattern['manual_value']}", "INFO")
        self.log_message(f"  Time: {pattern['time']} min", "INFO")
        
        # 実際のコマンド送信処理(これに2-3秒かかる)
        success = self.send_pattern_command(pattern)
        
        if success:
            self.log_message(f"パターン {pattern['index']} 送信成功", "SUCCESS")
            
            # コマンド送信完了後に時刻を設定
            current_time = time.time()
            
            # ホールド中でない場合
            if not self.is_holding:
                # パターン経過時間は常に現在時刻から新規スタート
                self.pattern_start_time = current_time
                
                # トータル経過時間も同時に設定(完全同期)
                if current_index == 0:
                    # 最初のパターンの場合、両方を現在時刻に設定
                    self.total_start_time = current_time
                elif hasattr(self, 'saved_total_elapsed_sec'):
                    # スキップ後またはパターン完了後の場合、保存した秒数を使ってtotal_start_timeを設定
                    # pattern_start_timeから秒数分引いた時刻に設定
                    self.total_start_time = self.pattern_start_time - self.saved_total_elapsed_sec
                    delattr(self, 'saved_total_elapsed_sec')
                
                # ホールド中でない場合はhold_start_timeをクリア
                if hasattr(self, 'hold_start_time'):
                    delattr(self, 'hold_start_time')
                    
            else:
                # ホールド中の場合
                # パターン経過時間もトータル経過時間も、「再開」まで進まない
                
                # pattern_start_timeを現在時刻に設定
                self.pattern_start_time = current_time
                
                # トータル経過時間の設定
                if current_index == 0:
                    self.total_start_time = current_time
                elif hasattr(self, 'saved_total_elapsed_sec'):
                    # スキップ後でホールド中の場合
                    # pattern_start_timeから秒数分引いた時刻に設定
                    self.total_start_time = self.pattern_start_time - self.saved_total_elapsed_sec
                    delattr(self, 'saved_total_elapsed_sec')
            
            # 指定時間待機(msで指定)
            wait_time = pattern['time'] * 60 * 1000  # 分→ミリ秒
            
            # スキップフラグをクリアして待機
            self.skip_requested = False
            self.wait_with_skip_check(wait_time, patterns, current_index)
        else:
            self.log_message(f"パターン {pattern['index']} 送信失敗", "ERROR")
            self.finish_test()
    
    def send_pattern_command(self, pattern):
        """パターンに基づいてコマンドを送信(選択されたDEFに対して)"""
        try:
            # コマンド文字列を生成
            dataset = pattern['dataset']
            pole = pattern['pole']
            code = pattern['code']
            manual_value = pattern['manual_value']
            
            # DACタイプを決定
            dac_type = "P" if dataset == "Position" else "L"
            
            # Manual選択時の処理
            if code == 'Manual':
                # Manual値の検証
                if not manual_value:
                    self.log_message("  エラー: Manual値が入力されていません", "ERROR")
                    return False
                
                # HEX妥当性チェック
                try:
                    int(manual_value, 16)
                except ValueError:
                    self.log_message(f"  エラー: Manual値が不正です ({manual_value})", "ERROR")
                    return False
                
                # Positionは5桁、LBCは4桁をチェック
                if dataset == 'Position' and len(manual_value) != 5:
                    self.log_message("  エラー: Positionは5桁のHEX値を入力してください", "ERROR")
                    return False
                elif dataset == 'LBC' and len(manual_value) != 4:
                    self.log_message("  エラー: LBCは4桁のHEX値を入力してください", "ERROR")
                    return False
                
                # Manual値を使用
                hex_value = manual_value.upper()
                
                # Neg選択時は反転処理
                if pole == 'Neg':
                    # HEX値を反転(ビット反転)
                    hex_int = int(hex_value, 16)
                    if dataset == 'Position':
                        hex_int = 0xFFFFF - hex_int  # 20ビット反転
                        hex_value = f"{hex_int:05X}"
                    else:  # LBC
                        hex_int = 0xFFFF - hex_int  # 16ビット反転
                        hex_value = f"{hex_int:04X}"
            
            else:
                # プリセット値の場合の処理(DAC操作タブと同じロジック)
                # プリセット値の定義
                presets = {
                    '+Full': {"P": "FFFFF", "L": "FFFF"},
                    'Center': {"P": "80000", "L": "8000"},
                    '-Full': {"P": "00000", "L": "0000"}
                }
                
                # コード値からHEX値を取得
                if code in presets:
                    hex_value = presets[code][dac_type]
                else:
                    self.log_message(f"  エラー: 不正なコード値 ({code})", "ERROR")
                    return False
                
                # Neg選択時は反転処理
                if pole == 'Neg':
                    # HEX値を反転(ビット反転)
                    hex_int = int(hex_value, 16)
                    if dataset == 'Position':
                        hex_int = 0xFFFFF - hex_int  # 20ビット反転
                        hex_value = f"{hex_int:05X}"
                    else:  # LBC
                        hex_int = 0xFFFF - hex_int  # 16ビット反転
                        hex_value = f"{hex_int:04X}"
            
            # 選択されているDEFに対してDACコマンド送信(変更部分)
            for i, var in enumerate(self.def_check_vars):
                if var.get():
                    cmd = f"DEF {i} DAC {dac_type} {hex_value}\r"
                    self.log_message(f"  送信: {cmd.strip()}", "INFO")

                    try:
                        self.serial_mgr.write(cmd.encode("utf-8"))
                        # レスポンス読み取り
                        self._read_response()
                    except Exception as e:
                        self.log_message(f"  エラー: {str(e)}", "ERROR")
                        return False

            # 選択していないDataSetにCenterコードを送信（オプション有効時）
            if self.send_opposite_center.get():
                # 選択していないDataSetのタイプとCenter値を決定
                if dataset == 'Position':
                    opposite_type = 'L'
                    opposite_center = '8000'
                else:  # LBC
                    opposite_type = 'P'
                    opposite_center = '80000'

                # Neg選択時は反転処理
                if pole == 'Neg':
                    opposite_int = int(opposite_center, 16)
                    if opposite_type == 'P':
                        opposite_int = 0xFFFFF - opposite_int  # 20ビット反転
                        opposite_center = f"{opposite_int:05X}"
                    else:  # L
                        opposite_int = 0xFFFF - opposite_int  # 16ビット反転
                        opposite_center = f"{opposite_int:04X}"

                # 選択されているDEFに対して送信
                for i, var in enumerate(self.def_check_vars):
                    if var.get():
                        cmd = f"DEF {i} DAC {opposite_type} {opposite_center}\r"
                        self.log_message(f"  送信(Center): {cmd.strip()}", "INFO")

                        try:
                            self.serial_mgr.write(cmd.encode("utf-8"))
                            self._read_response()
                        except Exception as e:
                            self.log_message(f"  エラー: {str(e)}", "ERROR")
                            return False

            return True
            
        except Exception as e:
            self.log_message(f"  例外発生: {str(e)}", "ERROR")
            return False
    
    def _read_response(self):
        """
        レスポンスを読み取り(DAC操作タブと同じロジック)
        1文字ずつ読み取り、CR/LFで即時改行。'>'で1行入れて終了。タイムアウト3秒。
        """
        line_buffer = ""
        timeout = time.time() + 3
        while time.time() < timeout:
            ch = self.serial_mgr.read()
            if not ch:
                continue
            if ch in ("\r", "\n"):
                if line_buffer.strip():
                    self.log_message(f"    レスポンス: {line_buffer}", "SUCCESS")
                    line_buffer = ""
            elif ch == ">":
                self.log_message("    レスポンス: >", "SUCCESS")
                break
            else:
                line_buffer += ch
    
    def stop_test(self):
        """パターンテストを停止"""
        self.is_running = False
        self.log_message("\nテストを停止しました", "WARNING")
        self.finish_test()
    
    def finish_test(self):
        """テスト終了処理"""
        self.is_running = False
        self.is_holding = False
        self.skip_requested = False
        
        # ★★★ テスト終了時に current_pattern_index をリセット ★★★
        self.current_pattern_index = -1
        
        self.start_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.skip_button.config(state=tk.DISABLED)
        self.hold_button.config(state=tk.DISABLED, text="ホールド")
        
        # 一時変数をクリア
        if hasattr(self, 'saved_total_elapsed_sec'):
            delattr(self, 'saved_total_elapsed_sec')
        if hasattr(self, 'held_pattern_elapsed_sec'):
            delattr(self, 'held_pattern_elapsed_sec')
        if hasattr(self, 'held_total_elapsed_sec'):
            delattr(self, 'held_total_elapsed_sec')
        if hasattr(self, 'hold_start_time'):
            delattr(self, 'hold_start_time')
        
        # 実行状況を「完了」に設定(リセットしない)
        self.current_pattern_label.config(text="完了")
        # パターン経過時間とトータル経過時間はそのまま保持
        
        self.log_message("=" * 40, "INFO")
        self.log_message("パターンテスト完了", "INFO")
        self.log_message("=" * 40, "INFO")
    
    def log_message(self, message, level="INFO"):
        """ログメッセージを表示"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n", level)
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.update_idletasks()
        
    def update_status_display(self, current_index, total_patterns):
        """実行状況の表示を更新"""
        if current_index < len(total_patterns):
            pattern = total_patterns[current_index]
            self.current_pattern_label.config(text=f"No.{pattern['index']}")
            self.current_pattern_time = pattern['time'] * 60  # 分を秒に変換
        else:
            self.current_pattern_label.config(text="完了")
    
    def update_time_display(self):
        """時間表示を更新(100msごとに呼ばれる)"""
        if not self.is_running:
            return
        
        current_time = time.time()
        
        if not self.is_holding:
            # ホールド中でなければ時間を更新
            # パターン経過時間
            pattern_elapsed = int(current_time - self.pattern_start_time + 0.5)  # 四捨五入
            if pattern_elapsed >= 0:  # マイナス値は表示しない
                pattern_total = int(self.current_pattern_time)
                pattern_elapsed_min = pattern_elapsed // 60
                pattern_elapsed_sec = pattern_elapsed % 60
                pattern_total_min = pattern_total // 60
                pattern_total_sec = pattern_total % 60
                self.pattern_time_label.config(
                    text=f"{pattern_elapsed_min:02d}:{pattern_elapsed_sec:02d} / {pattern_total_min:02d}:{pattern_total_sec:02d}"
                )
            
            # トータル経過時間
            total_elapsed = int(current_time - self.total_start_time + 0.5)  # 四捨五入
            if total_elapsed >= 0:  # マイナス値は表示しない
                total_patterns_time_int = int(self.total_patterns_time)
                total_elapsed_min = total_elapsed // 60
                total_elapsed_sec = total_elapsed % 60
                total_total_min = total_patterns_time_int // 60
                total_total_sec = total_patterns_time_int % 60
                self.total_time_label.config(
                    text=f"{total_elapsed_min:02d}:{total_elapsed_sec:02d} / {total_total_min:02d}:{total_total_sec:02d}"
                )
        # ホールド中は表示を更新しない(時間が止まって見える)

        # 100ms後に再度更新
        self.after(100, self.update_time_display)

    def get_pattern_remaining_seconds(self):
        """現在のパターンの残り秒数を取得

        Returns:
            float: 残り秒数（パターン実行中でない場合はNone）
        """
        if not self.is_running or self.is_holding:
            return None

        current_time = time.time()
        pattern_elapsed = current_time - self.pattern_start_time
        remaining = self.current_pattern_time - pattern_elapsed

        return max(0, remaining)

    def get_current_pattern_index(self):
        """現在実行中のパターンインデックスを取得

        Returns:
            int: パターンインデックス（実行中でない場合は-1）
        """
        return self.current_pattern_index if self.is_running else -1

    def load_settings(self):
        """設定ファイルから前回値を読み込み"""
        import json
        import os
        
        try:
            if os.path.exists('app_settings.json'):
                with open('app_settings.json', 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    
                    # test設定を読み込み
                    if 'test' in settings:
                        test_settings = settings['test']
                        self.save_folder.set(test_settings.get('save_folder', ''))
                        self.default_filename = test_settings.get('filename', 'pattern')
                        
                        # DEFチェック状態を読み込み（追加）
                        if 'def_checks' in test_settings:
                            self.saved_def_checks = test_settings['def_checks']
                        else:
                            self.saved_def_checks = [i == 0 for i in range(6)]  # DEF0のみTrue
                        
                        # スキャナーチャンネル設定を読み込み(Pos/Neg別々に)
                        if 'scanner_channels_pos' in test_settings:
                            self.saved_scanner_channels_pos = test_settings['scanner_channels_pos']
                        else:
                            self.saved_scanner_channels_pos = ['ー'] * 6
                            
                        if 'scanner_channels_neg' in test_settings:
                            self.saved_scanner_channels_neg = test_settings['scanner_channels_neg']
                        else:
                            self.saved_scanner_channels_neg = ['ー'] * 6

                        # 未選択DataSetにCenter送信オプション（デフォルト: True）
                        self.send_opposite_center.set(test_settings.get('send_opposite_center', True))
                    else:
                        self.default_filename = 'pattern'
                        self.saved_def_checks = [i == 0 for i in range(6)]
                        self.saved_scanner_channels_pos = ['ー'] * 6
                        self.saved_scanner_channels_neg = ['ー'] * 6
            else:
                self.default_filename = 'pattern'
                self.saved_def_checks = [i == 0 for i in range(6)]
                self.saved_scanner_channels_pos = ['ー'] * 6
                self.saved_scanner_channels_neg = ['ー'] * 6
        except Exception as e:
            print(f"設定ファイル読み込みエラー: {e}")
            self.default_filename = 'pattern'
            self.saved_def_checks = [i == 0 for i in range(6)]
            self.saved_scanner_channels_pos = ['ー'] * 6
            self.saved_scanner_channels_neg = ['ー'] * 6
    
    def save_settings(self):
        """設定ファイルに現在値を保存"""
        import json
        import os
        
        try:
            # 既存の設定を読み込み
            settings = {}
            if os.path.exists('app_settings.json'):
                with open('app_settings.json', 'r', encoding='utf-8') as f:
                    settings = json.load(f)
            
            # test設定を更新(DEFチェック状態も保存)
            settings['test'] = {
                'save_folder': self.save_folder.get(),
                'filename': self.filename_entry.get().strip(),
                'def_checks': [var.get() for var in self.def_check_vars],  # 追加
                'scanner_channels_pos': [ch.get() for ch in self.scanner_channels_pos],
                'scanner_channels_neg': [ch.get() for ch in self.scanner_channels_neg],
                'send_opposite_center': self.send_opposite_center.get()  # 未選択DataSetにCenter送信
            }
            
            # 設定ファイルに書き込み
            with open('app_settings.json', 'w', encoding='utf-8') as f:
                json.dump(settings, f, indent=4, ensure_ascii=False)
                
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")
            
    def skip_pattern(self):
        """現在のパターンをスキップ"""
        self.skip_requested = True
        self.log_message("現在のパターンをスキップします", "WARNING")
    
    def abort_test(self):
        """テストを中断(現在のパターン完了後に停止)"""
        self.abort_requested = True
        self.log_message("テストを中断します(現在のパターン完了後)", "WARNING")
        
    def wait_with_skip_check(self, wait_time_ms, patterns, current_index):
        """スキップとホールドをチェックしながら待機"""
        
        def check_skip():
            if not self.is_running:
                # 停止が要求された
                self.finish_test()
            elif self.skip_requested:
                # スキップが要求された
                self.log_message("パターンをスキップしました", "WARNING")
                self.skip_requested = False
                self._last_hold_start = None
                
                # スキップ時に現在のトータル経過時間を整数秒で保存
                current_time = time.time()
                if self.total_start_time < current_time:
                    # トータル経過時間の秒数(整数)を計算
                    total_elapsed_sec = int(current_time - self.total_start_time + 0.5)  # 四捨五入
                    self.saved_total_elapsed_sec = total_elapsed_sec
                else:
                    self.saved_total_elapsed_sec = 0
                
                # ホールド中にスキップした場合、パターン経過時間の保存をクリア
                if self.is_holding and hasattr(self, 'held_pattern_elapsed_sec'):
                    # パターン経過時間は0からスタートするため、保存値をクリア
                    delattr(self, 'held_pattern_elapsed_sec')
                
                # 次のパターンへ
                self.execute_patterns(patterns, current_index + 1)
            else:
                current_time = time.time()
                
                # pattern_start_timeからの実経過時間を計算(ホールド時間を除く)
                if self.is_holding:
                    # ホールド中は経過時間をカウントしない
                    actual_elapsed = (self.pattern_start_time - current_time) * 1000
                else:
                    # pattern_start_timeからの経過時間をミリ秒で計算
                    actual_elapsed = (current_time - self.pattern_start_time) * 1000
                
                if actual_elapsed >= wait_time_ms:
                    # 待機時間が完了
                    self._last_hold_start = None
                    
                    # パターン完了時にトータル経過時間を整数秒で保存
                    if self.total_start_time < current_time:
                        total_elapsed_sec = int(current_time - self.total_start_time + 0.5)  # 四捨五入
                        self.saved_total_elapsed_sec = total_elapsed_sec
                    
                    self.execute_patterns(patterns, current_index + 1)
                else:
                    # まだ待機中
                    self.after(100, check_skip)
        
        self._last_hold_start = None
        check_skip()
        
    def hold_test(self):
        """テストをホールド(一時停止)または再開"""
        current_time = time.time()
        
        if self.is_holding:
            # 再開
            self.is_holding = False
            self.hold_button.config(text="ホールド")
            
            # トータル経過時間の再計算(先に実行)
            if hasattr(self, 'held_total_elapsed_sec'):
                # 保存した秒数から継続
                self.total_start_time = current_time - self.held_total_elapsed_sec
                delattr(self, 'held_total_elapsed_sec')
            
            # パターン経過時間の再計算
            if hasattr(self, 'held_pattern_elapsed_sec'):
                # ホールド→再開の場合、保存した秒数から継続
                self.pattern_start_time = current_time - self.held_pattern_elapsed_sec
                delattr(self, 'held_pattern_elapsed_sec')
            else:
                # ホールド中にスキップした場合、pattern_start_timeを現在時刻に設定(0:00から)
                self.pattern_start_time = current_time
            
            self.log_message(f"テストを再開しました", "INFO")
            
            # hold_start_timeをクリア
            if hasattr(self, 'hold_start_time'):
                delattr(self, 'hold_start_time')
        else:
            # ホールド
            self.is_holding = True
            self.hold_start_time = current_time  # ホールド開始時刻を記録
            
            # ホールド時に現在の経過時間を四捨五入で保存
            # パターン経過時間
            if self.pattern_start_time < current_time:
                self.held_pattern_elapsed_sec = int(current_time - self.pattern_start_time + 0.5)  # 四捨五入
            else:
                self.held_pattern_elapsed_sec = 0
            
            # トータル経過時間
            if self.total_start_time < current_time:
                self.held_total_elapsed_sec = int(current_time - self.total_start_time + 0.5)  # 四捨五入
            else:
                self.held_total_elapsed_sec = 0
            
            self.hold_button.config(text="再開")
            self.log_message("テストをホールドしました", "WARNING")
            
    def open_measurement_window(self):
        """計測ウィンドウを開く"""
        # ★★★ 既にウィンドウが開いている場合は前面に表示 ★★★
        if self.measurement_window is not None:
            try:
                self.measurement_window.lift()
                self.measurement_window.focus_force()
                self.log_message("計測ウィンドウは既に開いています", "WARNING")
                return
            except:
                # ウィンドウが閉じられていた場合
                self.measurement_window = None
        
        try:
            from tabs.measurement_window import MeasurementWindow
            
            # 親ウィンドウ（main.pyのroot）を取得
            parent = self.winfo_toplevel()
            gpib_3458a = parent.gpib_3458a
            gpib_3499b = parent.gpib_3499b
            
            # 計測ウィンドウを開く
            self.measurement_window = MeasurementWindow(parent, gpib_3458a, gpib_3499b, self)
            
            # ★★★ ウィンドウが閉じられた時のコールバックを設定 ★★★
            self.measurement_window.protocol("WM_DELETE_WINDOW", self.on_measurement_window_close)
            
            # ★★★ 計測ボタンを無効化 ★★★
            self.measurement_button.config(state=tk.DISABLED)
            
            self.measurement_window.focus()
            
            self.log_message("計測ウィンドウを開きました", "INFO")
        except Exception as e:
            messagebox.showerror("エラー", f"計測ウィンドウを開けません: {e}")
            self.log_message(f"計測ウィンドウを開けません: {str(e)}", "ERROR")
            self.measurement_window = None
            
    def on_measurement_window_close(self):
        """計測ウィンドウが閉じられた時の処理"""
        if self.measurement_window is not None:
            try:
                # 計測が実行中の場合は停止確認
                if hasattr(self.measurement_window, 'is_measuring') and self.measurement_window.is_measuring:
                    if messagebox.askyesno("確認", "計測中です。停止しますか?"):
                        self.measurement_window.stop_measurement()
                        try:
                            self.measurement_window.open_all_used_channels()
                            self.measurement_window.gpib_scanner.instrument.control_ren(0)
                        except:
                            pass
                    else:
                        return  # キャンセルされた場合は閉じない
                else:
                    # 計測停止中の場合はチャンネルをOPEN
                    try:
                        self.measurement_window.open_all_used_channels()
                        self.measurement_window.gpib_scanner.instrument.control_ren(0)
                    except:
                        pass
                
                # タイマーをキャンセル
                if hasattr(self.measurement_window, 'update_timer_id') and self.measurement_window.update_timer_id is not None:
                    self.measurement_window.after_cancel(self.measurement_window.update_timer_id)
                
                # ウィンドウを破棄
                self.measurement_window.destroy()
            except:
                pass
            finally:
                self.measurement_window = None
                # ★★★ 計測ボタンを有効化 ★★★
                self.measurement_button.config(state=tk.NORMAL)
                self.log_message("計測ウィンドウを閉じました", "INFO")