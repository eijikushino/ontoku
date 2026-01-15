import tkinter as tk
from tkinter import ttk, filedialog
import os
import json
import re


class FileTab(ttk.Frame):
    """
    ファイル保存設定タブ（通信1のみ）
    
    ・横幅を抑えるため、UIは基本「縦並び・左寄せ・短いラベル」
    ・上段:保存先フォルダ（共通）… Entry → 参照（縦配置、左寄せ）
    ・下段:通信1の設定… S/N(DEF0〜DEF5) と CSV名 を縦並び・左寄せ
    ・通信タブと同じ配色（#81D4FA: 水色）で色分け（左側色バー）
    ・入力はすべて自動保存（フォーカスアウト／Enter／入力後300ms）
    ・設定保存先:app_settings.json
        save_config.save_dir                 # 共通フォルダ
        comm_profiles.1.serial_numbers       # 通信1のDEFシリアル
        comm_profiles.1.save_config.file_name  # 通信1のCSV名
        comm_profiles.1.device_type          # デバイスタイプ
    """
    
    # 通信1の配色（水色）
    COMM1_COLOR = "#81D4FA"
    
    def __init__(self, parent):
        super().__init__(parent)
        
        # 設定ファイルパス
        self.config_file = "app_settings.json"
        self.config = self._load_config()
        self._apply_jobs = {}  # 遅延保存ジョブ
        
        self._create_widgets()
    
    def _load_config(self):
        """設定ファイル読み込み"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"設定ファイル読み込みエラー: {e}")
                return {}
        return {}
    
    def _save_config(self):
        """設定ファイル保存"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")
    
    def _create_widgets(self):
        """ウィジェット作成"""
        # ====== 上段:保存先（共通） ======
        folder_lf = ttk.LabelFrame(self, text="保存先（共通）", padding=10)
        folder_lf.pack(fill="x", padx=10, pady=(10, 6))
        
        # ラベル（左寄せ）
        ttk.Label(folder_lf, text="保存先").pack(anchor="w")
        
        # Entry（左寄せ・幅控えめ）
        self.dir_var = tk.StringVar(value=self._get_save_dir())
        dir_entry = ttk.Entry(folder_lf, textvariable=self.dir_var, width=50)
        dir_entry.pack(anchor="w", fill="x", pady=(2, 4))
        dir_entry.bind("<FocusOut>", lambda e: self._save_global_dir())
        dir_entry.bind("<Return>", lambda e: self._save_global_dir())
        # 変更のたびにヒント文字列を更新
        self.dir_var.trace_add("write", lambda *a: self._refresh_dir_hint())
        
        # 参照ボタン（左寄せ）
        ttk.Button(folder_lf, text="参照", command=self._browse_folder).pack(anchor="w", pady=(0, 2))
        
        # ====== 下段:通信1の設定 ======
        comm_frame = self._build_comm1_frame(self)
        comm_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))
    
    def _build_comm1_frame(self, parent):
        """通信1の設定フレーム構築"""
        # 外枠は tk.Frame で背景色を設定
        root = tk.Frame(parent, bg=self.COMM1_COLOR, bd=0, highlightthickness=0)
        
        # 左の色バー（列0）は固定、右のUI（列1）を使う
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=0)
        root.grid_columnconfigure(1, weight=1)
        
        # 左バー
        left_bar = tk.Frame(root, bg=self.COMM1_COLOR, width=8)
        left_bar.grid(row=0, column=0, sticky="ns")
        left_bar.grid_propagate(False)
        
        # 内側UIは ttk
        inner = ttk.Frame(root)
        inner.grid(row=0, column=1, sticky="nsew", padx=(8, 8), pady=8)
        
        # 見出し
        ttk.Label(inner, text="ファイル保存設定（通信1）", 
                  font=("", 10, "bold")).pack(anchor="w", pady=(0, 10))
        
        # --- DEFタイプ選択(Main/Sub DEF 排他) ---
        type_lf = ttk.LabelFrame(inner, text="DEFタイプ", padding=10)
        type_lf.pack(anchor="w", fill="x", pady=(0, 6))
        
        init_device_type = self._get_device_type()
        self.device_type_var = tk.StringVar(value=init_device_type)
        
        type_row = ttk.Frame(type_lf)
        type_row.pack(anchor="w")
        
        ttk.Radiobutton(type_row, text="Main DEF (DFHxxx)", 
                       variable=self.device_type_var, 
                       value="main", 
                       command=self._on_device_type_changed).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(type_row, text="Sub DEF (SUBxxx)", 
                       variable=self.device_type_var, 
                       value="sub", 
                       command=self._on_device_type_changed).pack(side="left")
        
        # --- S/N(縦並び) ---
        sn_lf = ttk.LabelFrame(inner, text="S/N（番号のみ入力）", padding=10)
        sn_lf.pack(anchor="w", fill="x", pady=(0, 6))
        
        self.sn_vars = []
        for i in range(6):
            row = ttk.Frame(sn_lf)
            row.pack(anchor="w", fill="x", pady=2)
            
            ttk.Label(row, text=f"DEF{i}", width=6).pack(side="left", anchor="w")
            
            # 既存のS/Nからプレフィックスを除去して番号のみ表示
            full_sn = self._get_serial_number(i)
            number_only = self._extract_number_only(full_sn)
            
            v = tk.StringVar(value=number_only)
            ent = ttk.Entry(row, textvariable=v, width=16)
            ent.pack(side="left", anchor="w", padx=(6, 0))
            
            # 自動保存(遅延300ms)
            v.trace_add("write", 
                       lambda *_a, _idx=i, _var=v: self._schedule_apply(("sn", _idx, _var)))
            ent.bind("<FocusOut>", 
                    lambda e, _idx=i, _var=v: self._apply_now(("sn", _idx, _var)))
            self.sn_vars.append(v)
        
        # --- CSV名(縦並び) ---
        file_lf = ttk.LabelFrame(inner, text="CSVファイル名", padding=10)
        file_lf.pack(anchor="w", fill="x")
        
        # 行内に「ファイル名」ラベル + Entry + 「.csv」固定表示
        row = ttk.Frame(file_lf)
        row.pack(anchor="w", fill="x")
        
        ttk.Label(row, text="ファイル名").pack(side="left", anchor="w")
        
        # 既存設定から .csv を外して表示
        full = self._get_file_name()
        stem = full[:-4] if full.lower().endswith(".csv") else full
        
        self.file_var = tk.StringVar(value=stem)
        file_ent = ttk.Entry(row, textvariable=self.file_var, width=32)
        file_ent.pack(side="left", anchor="w", padx=(6, 0))
        
        ttk.Label(row, text=".csv").pack(side="left", anchor="w", padx=(6, 0))
        
        # 自動保存(遅延300ms)
        self.file_var.trace_add("write", 
                               lambda *_a: self._schedule_apply(("file", self.file_var)))
        file_ent.bind("<FocusOut>", 
                     lambda e: self._apply_now(("file", self.file_var)))
        file_ent.bind("<Return>", 
                     lambda e: self._apply_now(("file", self.file_var)))
        
        # 共通保存先のヒント(左寄せ)
        self.dir_hint_var = tk.StringVar()
        ttk.Label(file_lf, textvariable=self.dir_hint_var, 
                 foreground="gray").pack(anchor="w", pady=(4, 0))
        self._refresh_dir_hint()
        
        return root
    
    # ---------- 初期値取得 ----------
    def _get_save_dir(self):
        """保存先ディレクトリ取得"""
        return self.config.get("save_config", {}).get("save_dir", "reports")
    
    def _get_device_type(self):
        """デバイスタイプ取得"""
        return self.config.get("comm_profiles", {}).get("1", {}).get("device_type", "main")
    
    def _get_file_name(self):
        """CSVファイル名取得"""
        # 通信1の設定 → 無ければグローバル → 既定
        return (
            self.config.get("comm_profiles", {}).get("1", {}).get("save_config", {}).get("file_name")
            or self.config.get("save_config", {}).get("file_name")
            or "test_results_wide.csv"
        )
    
    def _get_serial_number(self, def_index):
        """S/N取得"""
        # 通信1の設定 → 無ければグローバル → ""
        return (
            self.config.get("comm_profiles", {}).get("1", {}).get("serial_numbers", {}).get(f"DEF{def_index}_sn")
            or self.config.get("serial_numbers", {}).get(f"DEF{def_index}_sn", "")
        )
    
    def _ensure_comm_profile_keys(self):
        """通信1のプロファイルキーを確保"""
        self.config.setdefault("comm_profiles", {})
        cp = self.config["comm_profiles"].setdefault("1", {})
        cp.setdefault("serial_numbers", {})
        cp.setdefault("save_config", {})
        return cp
    
    # ---------- S/N処理 ----------
    def _extract_number_only(self, full_sn):
        """
        完全なS/N(例: DFH903, SUBJ02, SUB025)から番号部分のみを抽出
        空文字列の場合はそのまま返す
        アルファベット+数字の組み合わせにも対応
        """
        if not full_sn:
            return ""
        
        # DFH, SUB などのプレフィックスを除去（アルファベット+数字に対応）
        match = re.match(r'^(?:DFH|SUB|H)?([A-Z]?\d+)$', full_sn, re.IGNORECASE)
        if match:
            return match.group(1).upper()  # 大文字に統一
        return full_sn  # マッチしない場合はそのまま返す
    
    # ---------- イベントハンドラ ----------
    def _on_device_type_changed(self):
        """DEFタイプ変更時の処理"""
        cp = self._ensure_comm_profile_keys()
        device_type = self.device_type_var.get()
        cp["device_type"] = device_type
        
        # 既存の全シリアルナンバーのプレフィックスを変換
        for i in range(6):
            # 現在保存されているフルS/Nを取得
            current_sn = cp["serial_numbers"].get(f"DEF{i}_sn", "")
            if current_sn:
                # 番号部分のみを抽出
                number_only = self._extract_number_only(current_sn)
                if number_only:
                    # 新しいDEFタイプに応じたプレフィックスを付与
                    if device_type == "main":
                        new_sn = f"DFH{number_only}"
                    else:  # sub
                        new_sn = f"SUB{number_only}"
                    
                    # 設定に保存
                    cp["serial_numbers"][f"DEF{i}_sn"] = new_sn
        
        self._save_config()
    
    def _apply_now(self, payload):
        """即時保存(フォーカスアウト/Enterなど)"""
        cp = self._ensure_comm_profile_keys()
        
        if payload[0] == "sn":
            # ("sn", def_index, var)
            _, i_def, var = payload
            device_type = self.device_type_var.get()
            number = (var.get() or "").strip().upper()  # ★大文字に統一
            
            # プレフィックスを自動付与
            if number:
                if device_type == "main":
                    full_sn = f"DFH{number}"
                else:  # sub
                    full_sn = f"SUB{number}"
            else:
                full_sn = ""
            
            cp["serial_numbers"][f"DEF{i_def}_sn"] = full_sn
        
        elif payload[0] == "file":
            # ("file", var)  ← Entryには拡張子なしで表示し続ける
            _, var = payload
            stem = (var.get() or "").strip() or "test_results_wide"
            save_name = stem if stem.lower().endswith(".csv") else f"{stem}.csv"
            cp["save_config"]["file_name"] = save_name
        
        self._save_config()
    
    def _schedule_apply(self, payload):
        """入力中のチラつきを抑えるため 300ms でまとめて保存"""
        key = "comm1"
        job = self._apply_jobs.get(key)
        if job:
            try:
                self.after_cancel(job)
            except Exception:
                pass
        self._apply_jobs[key] = self.after(
            300, lambda _payload=payload: self._apply_now(_payload)
        )
    
    # ---------- 共通保存先 ----------
    def _browse_folder(self):
        """フォルダ参照ダイアログ"""
        folder = filedialog.askdirectory()
        if folder:
            self.dir_var.set(folder)
            self._save_global_dir()
    
    def _save_global_dir(self):
        """保存先ディレクトリを保存"""
        path = (self.dir_var.get() or "").strip()
        if not path:
            return
        self.config.setdefault("save_config", {})["save_dir"] = path
        self._save_config()
        self._refresh_dir_hint()
    
    def _refresh_dir_hint(self):
        """保存先ヒントを更新"""
        text = f"※保存先は共通：{self.dir_var.get() or self._get_save_dir()}"
        self.dir_hint_var.set(text)