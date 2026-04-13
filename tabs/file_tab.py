import tkinter as tk
from tkinter import ttk, filedialog
import os
import json
import re


class FileTab(ttk.Frame):
    """
    ファイル保存設定タブ

    - 保存先を 3 種類（Pattern Test(温特) / Linearity / DC特性）まとめて表示
    - Linearity / DC特性 の StringVar は個別タブと共有するため、
      どちらの画面で変更しても相互に同期される
    - DEFシリアル番号・CSVファイル名はこのタブで保持（app_settings.json）
    """

    # タブ全体の額縁余白
    TAB_OUTER_PADX = 20
    TAB_OUTER_PADY = 16
    # セクション間 / 内側パディング
    SECTION_GAP = 14
    INNER_PADDING = 12
    # 行間
    ROW_GAP = 4
    ROW_GAP_TIGHT = 1          # DEF S/N 行間を詰める
    # 幅
    ENTRY_WIDTH_PATH = 42
    ENTRY_WIDTH_SHORT = 22
    LABEL_WIDTH_DIR = 18
    # フォント
    HEADING_FONT = ("", 13, "bold")

    def __init__(self, parent, linearity_tab=None, dc_char_tab=None, test_tab=None):
        super().__init__(parent)
        self.linearity_tab = linearity_tab
        self.dc_char_tab = dc_char_tab
        self.test_tab = test_tab

        self.config_file = "app_settings.json"
        self.config = self._load_config()
        self._apply_jobs = {}

        self._create_widgets()

    # ==================== 設定ファイル I/O ====================
    def _load_config(self):
        """設定ファイル読み込み"""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"設定ファイル読み込みエラー: {e}")
        return {}

    def _save_config(self):
        """
        設定ファイル保存。
        他タブ(linearity / dc_char など)の書き込みを上書きしないよう、
        毎回ディスクを読み直して FileTab 所有セクションのみ差し替える。
        """
        try:
            fresh: dict = {}
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    fresh = json.load(f)
            for key in ("save_config", "comm_profiles", "serial_numbers"):
                if key in self.config:
                    fresh[key] = self.config[key]
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(fresh, f, indent=4, ensure_ascii=False)
            self.config = fresh
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")

    # ==================== ウィジェット構築 ====================
    def _create_widgets(self):
        """ウィジェット作成"""
        container = ttk.Frame(self)
        container.pack(
            fill="both", expand=True,
            padx=self.TAB_OUTER_PADX, pady=self.TAB_OUTER_PADY,
        )

        # ====== 保存先 ======
        folder_lf = ttk.LabelFrame(container, padding=self.INNER_PADDING)
        folder_lf.configure(
            labelwidget=tk.Label(
                folder_lf, text="保存先", font=self.HEADING_FONT
            )
        )
        folder_lf.pack(fill="x", pady=(0, self.SECTION_GAP))

        # --- Pattern Test(温特) ---
        self.pattern_dir_var = tk.StringVar(value=self._get_save_dir())
        self.pattern_dir_var.trace_add(
            "write", lambda *_: self._on_pattern_dir_changed()
        )
        self._build_dir_row(
            folder_lf, "Pattern Test(温特)",
            self.pattern_dir_var,
            lambda: self._pick_dir(self.pattern_dir_var),
        )
        # Pattern Test 用 CSV ファイル名（保存先の直下に配置）
        self._build_csv_name_row(folder_lf)
        # 保存される CSV フルパスの例
        self.pattern_hint_var = tk.StringVar(value="")
        self._build_hint_row(folder_lf, self.pattern_hint_var)

        # --- Linearity: 個別タブの StringVar を共有（両方向同期） ---
        if self.linearity_tab is not None and hasattr(self.linearity_tab, "save_dir"):
            self.linearity_dir_var = self.linearity_tab.save_dir
        else:
            self.linearity_dir_var = tk.StringVar(
                value=self.config.get("linearity", {}).get("save_dir", "linearity_data")
            )
            self.linearity_dir_var.trace_add(
                "write", lambda *_: self._save_section_dir("linearity", self.linearity_dir_var)
            )
        self._build_dir_row(
            folder_lf, "Linearity",
            self.linearity_dir_var,
            lambda: self._pick_dir(self.linearity_dir_var),
        )
        self.linearity_hint_var = tk.StringVar(value="")
        self._build_hint_row(folder_lf, self.linearity_hint_var)

        # --- DC特性: 個別タブの StringVar を共有（両方向同期） ---
        if self.dc_char_tab is not None and hasattr(self.dc_char_tab, "save_dir"):
            self.dc_dir_var = self.dc_char_tab.save_dir
        else:
            self.dc_dir_var = tk.StringVar(
                value=self.config.get("dc_char", {}).get("save_dir", "dc_char_data")
            )
            self.dc_dir_var.trace_add(
                "write", lambda *_: self._save_section_dir("dc_char", self.dc_dir_var)
            )
        self._build_dir_row(
            folder_lf, "DC特性",
            self.dc_dir_var,
            lambda: self._pick_dir(self.dc_dir_var),
        )
        self.dc_hint_var = tk.StringVar(value="")
        self._build_hint_row(folder_lf, self.dc_hint_var)

        # ====== DEFシリアル設定 ======
        self._build_comm1_frame(container).pack(fill="both", expand=True)

        # ファイル名ヒントの初期描画 + タブ表示時のみ更新
        self._refresh_hints()
        self.bind("<Visibility>", self._on_visibility)

    def _build_dir_row(self, parent, label, var, on_browse):
        """保存先 1 行（ラベル + Entry + 参照ボタン）"""
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(self.ROW_GAP, 0))

        ttk.Label(
            row, text=label, width=self.LABEL_WIDTH_DIR, anchor="w"
        ).pack(side="left")
        ttk.Entry(
            row, textvariable=var, width=self.ENTRY_WIDTH_PATH
        ).pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(
            row, text="参照", command=on_browse, width=5
        ).pack(side="left")

    def _build_hint_row(self, parent, textvariable):
        """保存先の直下に表示するファイル名ヒント行（グレー文字）"""
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(1, self.ROW_GAP))
        # ラベル幅分だけインデントしてディレクトリ Entry と左端を揃える
        ttk.Label(row, text="", width=self.LABEL_WIDTH_DIR).pack(side="left")
        ttk.Label(
            row, textvariable=textvariable, foreground="gray", font=("", 8)
        ).pack(side="left", anchor="w")

    def _build_csv_name_row(self, parent):
        """Pattern Test 用 CSV ファイル名行（保存先の直下に配置）"""
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(1, self.ROW_GAP))

        # ラベル幅分だけインデントしてディレクトリ Entry と左端を揃える
        ttk.Label(row, text="", width=self.LABEL_WIDTH_DIR).pack(side="left")
        ttk.Label(row, text="CSVファイル名:").pack(side="left")

        full = self._get_file_name()
        stem = full[:-4] if full.lower().endswith(".csv") else full

        self.file_var = tk.StringVar(value=stem)
        file_ent = ttk.Entry(
            row, textvariable=self.file_var, width=self.ENTRY_WIDTH_SHORT
        )
        file_ent.pack(side="left", anchor="w", padx=(6, 0))
        ttk.Label(row, text=".csv").pack(side="left", anchor="w", padx=(2, 0))

        self.file_var.trace_add(
            "write",
            lambda *_a: self._on_file_var_changed(),
        )
        file_ent.bind(
            "<FocusOut>", lambda e: self._apply_now(("file", self.file_var))
        )
        file_ent.bind(
            "<Return>", lambda e: self._apply_now(("file", self.file_var))
        )

    def _on_file_var_changed(self):
        """CSV ファイル名入力時: 保存(遅延) + ヒント即時更新"""
        self._schedule_apply(("file", self.file_var))
        self._refresh_pattern_hint()

    def _on_pattern_dir_changed(self):
        """Pattern Test 保存先変更時: 保存 + ヒント即時更新"""
        self._save_pattern_dir()
        self._refresh_pattern_hint()

    def _build_comm1_frame(self, parent):
        """DEFシリアルの設定フレーム（装飾枠なし・LabelFrameのみ）"""
        root = ttk.LabelFrame(parent, padding=self.INNER_PADDING)
        root.configure(
            labelwidget=tk.Label(
                root, text="DEFシリアル設定", font=self.HEADING_FONT
            )
        )

        # --- DEFタイプ選択 (Main/Sub DEF 排他) ---
        type_lf = ttk.LabelFrame(
            root, text="DEFタイプ", padding=self.INNER_PADDING
        )
        type_lf.pack(anchor="w", fill="x", pady=(0, self.INNER_PADDING))

        init_device_type = self._get_device_type()
        self.device_type_var = tk.StringVar(value=init_device_type)

        type_row = ttk.Frame(type_lf)
        type_row.pack(anchor="w")
        ttk.Radiobutton(
            type_row, text="Main DEF (DFHxxx)",
            variable=self.device_type_var, value="main",
            command=self._on_device_type_changed,
        ).pack(side="left", padx=(0, 12))
        ttk.Radiobutton(
            type_row, text="Sub DEF (SUBxxx)",
            variable=self.device_type_var, value="sub",
            command=self._on_device_type_changed,
        ).pack(side="left")

        # --- S/N（縦並び・行間を詰める） ---
        sn_lf = ttk.LabelFrame(
            root, text="S/N（番号のみ入力）", padding=self.INNER_PADDING
        )
        sn_lf.pack(anchor="w", fill="x")

        self.sn_vars = []
        for i in range(6):
            row = ttk.Frame(sn_lf)
            row.pack(anchor="w", fill="x", pady=self.ROW_GAP_TIGHT)

            ttk.Label(row, text=f"DEF{i}", width=6).pack(side="left", anchor="w")

            full_sn = self._get_serial_number(i)
            number_only = self._extract_number_only(full_sn)

            v = tk.StringVar(value=number_only)
            ent = ttk.Entry(row, textvariable=v, width=self.ENTRY_WIDTH_SHORT)
            ent.pack(side="left", anchor="w", padx=(6, 0))

            v.trace_add(
                "write",
                lambda *_a, _idx=i, _var=v: self._schedule_apply(("sn", _idx, _var)),
            )
            ent.bind(
                "<FocusOut>",
                lambda e, _idx=i, _var=v: self._apply_now(("sn", _idx, _var)),
            )
            self.sn_vars.append(v)

        return root

    # ==================== 初期値取得 ====================
    def _get_save_dir(self):
        """Pattern Test(温特) 保存先"""
        return self.config.get("save_config", {}).get("save_dir", "reports")

    def _get_device_type(self):
        return self.config.get("comm_profiles", {}).get("1", {}).get("device_type", "main")

    def _get_file_name(self):
        return (
            self.config.get("comm_profiles", {}).get("1", {}).get("save_config", {}).get("file_name")
            or self.config.get("save_config", {}).get("file_name")
            or "test_results_wide.csv"
        )

    def _get_serial_number(self, def_index):
        return (
            self.config.get("comm_profiles", {}).get("1", {}).get("serial_numbers", {}).get(f"DEF{def_index}_sn")
            or self.config.get("serial_numbers", {}).get(f"DEF{def_index}_sn", "")
        )

    def _ensure_comm_profile_keys(self):
        self.config.setdefault("comm_profiles", {})
        cp = self.config["comm_profiles"].setdefault("1", {})
        cp.setdefault("serial_numbers", {})
        cp.setdefault("save_config", {})
        return cp

    # ==================== S/N処理 ====================
    def _extract_number_only(self, full_sn):
        """完全なS/N(例: DFH903, SUBJ02)から番号部分のみを抽出"""
        if not full_sn:
            return ""
        match = re.match(r'^(?:DFH|SUB|H)?([A-Z]?\d+)$', full_sn, re.IGNORECASE)
        if match:
            return match.group(1).upper()
        return full_sn

    # ==================== イベントハンドラ ====================
    def _on_device_type_changed(self):
        cp = self._ensure_comm_profile_keys()
        device_type = self.device_type_var.get()
        cp["device_type"] = device_type

        for i in range(6):
            current_sn = cp["serial_numbers"].get(f"DEF{i}_sn", "")
            if current_sn:
                number_only = self._extract_number_only(current_sn)
                if number_only:
                    if device_type == "main":
                        new_sn = f"DFH{number_only}"
                    else:
                        new_sn = f"SUB{number_only}"
                    cp["serial_numbers"][f"DEF{i}_sn"] = new_sn

        self._save_config()

    def _apply_now(self, payload):
        """即時保存(フォーカスアウト/Enter)"""
        cp = self._ensure_comm_profile_keys()

        if payload[0] == "sn":
            _, i_def, var = payload
            device_type = self.device_type_var.get()
            number = (var.get() or "").strip().upper()

            if number:
                full_sn = f"DFH{number}" if device_type == "main" else f"SUB{number}"
            else:
                full_sn = ""

            cp["serial_numbers"][f"DEF{i_def}_sn"] = full_sn

        elif payload[0] == "file":
            _, var = payload
            stem = (var.get() or "").strip() or "test_results_wide"
            save_name = stem if stem.lower().endswith(".csv") else f"{stem}.csv"
            cp["save_config"]["file_name"] = save_name

        self._save_config()

    def _schedule_apply(self, payload):
        """入力中のチラつき抑制: 300ms 遅延保存"""
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

    # ==================== 保存先ハンドラ ====================
    def _pick_dir(self, var):
        """参照ダイアログで保存先を選択（現在の指定パスを初期表示）"""
        current = (var.get() or "").strip()
        initial = ""
        if current:
            abs_path = os.path.abspath(current)
            # 指定パスが存在すればそのまま、未作成なら存在する親まで遡る
            probe = abs_path
            while probe and not os.path.isdir(probe):
                parent = os.path.dirname(probe)
                if parent == probe:
                    probe = ""
                    break
                probe = parent
            initial = probe
        folder = filedialog.askdirectory(initialdir=initial or None)
        if folder:
            var.set(folder)

    def _save_pattern_dir(self):
        """Pattern Test(温特) 保存先を保存"""
        path = (self.pattern_dir_var.get() or "").strip()
        if not path:
            return
        self.config.setdefault("save_config", {})["save_dir"] = path
        self._save_config()

    # ==================== ファイル名ヒント ====================
    def _on_visibility(self, event):
        """タブが表示状態になった時だけヒントを再計算する"""
        # 完全に隠れている時は更新しない（計算負荷を抑える）
        if getattr(event, "state", None) == "VisibilityFullyObscured":
            return
        self._refresh_hints()

    def _refresh_hints(self):
        """現在の設定値からファイル名の例文字列を再生成"""
        self._refresh_pattern_hint()
        if hasattr(self, "linearity_hint_var"):
            self.linearity_hint_var.set(self._format_linearity_hint())
        if hasattr(self, "dc_hint_var"):
            self.dc_hint_var.set(self._format_dc_hint())

    def _refresh_pattern_hint(self):
        """Pattern Test ヒントのみ更新（入力即時用）"""
        if hasattr(self, "pattern_hint_var"):
            self.pattern_hint_var.set(self._format_pattern_hint())

    def _current_sns(self):
        """
        現在有効な DEF 群の完全 S/N リストを返す。
        Pattern Test タブの DEF 選択を優先し、なければ sn_vars の非空すべて。
        """
        device_type = getattr(self, "device_type_var", None)
        prefix = "DFH"
        if device_type is not None and device_type.get() == "sub":
            prefix = "SUB"

        # DEF 選択順を決定: Pattern Test のチェック状態を優先
        indices = list(range(6))
        if self.test_tab is not None and hasattr(self.test_tab, "def_check_vars"):
            checked = [
                i for i, v in enumerate(self.test_tab.def_check_vars) if v.get()
            ]
            if checked:
                indices = checked

        sns: list[str] = []
        if hasattr(self, "sn_vars"):
            for i in indices:
                if 0 <= i < len(self.sn_vars):
                    number = (self.sn_vars[i].get() or "").strip().upper()
                    if number:
                        sns.append(f"{prefix}{number}")
        if not sns:
            sns = [f"{prefix}000"]
        return sns

    def _sn_token(self):
        """S/N が複数なら [A|B|C] 形式、単独ならそのまま"""
        sns = self._current_sns()
        if len(sns) == 1:
            return sns[0]
        return "[" + "|".join(sns) + "]"

    def _format_pattern_hint(self):
        """Pattern Test の保存先 + CSV 名を組み立てたフルパス例"""
        save_dir = (self.pattern_dir_var.get() or "").strip() or "reports"
        if hasattr(self, "file_var"):
            stem = (self.file_var.get() or "").strip()
        else:
            stem = ""
        if not stem:
            stem = "test_results_wide"
        filename = stem if stem.lower().endswith(".csv") else f"{stem}.csv"
        try:
            path = os.path.join(save_dir, filename)
        except Exception:
            path = f"{save_dir}/{filename}"
        return f"例: {path}"

    def _format_linearity_hint(self):
        """Linearity タブの設定から保存ファイル名サンプルを生成"""
        sn = self._sn_token()
        lt = self.linearity_tab
        if lt is None or not hasattr(lt, "dac_var"):
            return f"例: {sn}_Position_linearity_出荷Sequence_YYYYMMDD_HHMMSS.xlsx / .png"
        dac = lt.dac_var.get() or "Position"
        raw_mode = (lt.pattern_mode.get() or "Ship").lower()
        if raw_mode == "ship":
            mode = "出荷Sequence"
        elif raw_mode == "linear":
            mode = "sequential"
        else:
            mode = raw_mode
        pts = (lt.num_points.get() or "").strip()
        pts_str = f"_{pts}pts" if raw_mode in ("random", "linear") and pts else ""
        pole = lt.pole_select.get() or "両極"
        if pole in ("POS", "NEG"):
            return (
                f"例: {sn}_{dac}_{pole}_linearity_{mode}{pts_str}"
                f"_YYYYMMDD_HHMMSS.xlsx / .png"
            )
        return (
            f"例: {sn}_{dac}_linearity_{mode}{pts_str}"
            f"_YYYYMMDD_HHMMSS.xlsx / .png"
        )

    def _format_dc_hint(self):
        """DC特性 タブの設定から保存ファイル名サンプルを生成"""
        sn = self._sn_token()
        dt = self.dc_char_tab
        if dt is None or not hasattr(dt, "test_type"):
            return f"例: {sn}_DC特性_POSTION_YYYYMMDD_HHMMSS.xlsx / .png"
        test_type = (dt.test_type.get() or "Position").lower()
        sheet_titles = {"position": "POSTION", "lbc": "LBC", "moni": "moni"}
        title = sheet_titles.get(test_type, test_type)
        return f"例: {sn}_DC特性_{title}_YYYYMMDD_HHMMSS.xlsx / .png"

    def _save_section_dir(self, section, var):
        """
        linearity_tab / dc_char_tab が未連携のときのフォールバック。
        (通常は個別タブの StringVar を共有しているので呼ばれない)
        """
        path = (var.get() or "").strip()
        if not path:
            return
        try:
            fresh: dict = {}
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    fresh = json.load(f)
            fresh.setdefault(section, {})["save_dir"] = path
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(fresh, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"設定ファイル保存エラー: {e}")
