import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
import time
import re
import threading
import queue
from utils.config_handler import load_config, update_config_value


class DataGenTab(ttk.Frame):
    """
    DataGen制御タブ：専用SerialManagerを使用してDataGenを操作
    内部でDataGen1/DataGen2を切り替え可能
    """

    def __init__(self, parent, datagen_manager, datagen_manager2=None):
        super().__init__(parent)
        self.datagen1 = datagen_manager
        self.datagen2 = datagen_manager2
        self.datagen = datagen_manager  # 現在アクティブなマネージャ
        self.current_dg = 1  # 1 or 2
        self.initialized = False
        self.initialized2 = False  # DataGen2用

        # グリッチ制御用
        self._glitch_thread = None
        self._glitch_running = False
        self._glitch_paused = False

        # 常駐リーダー用
        self._text_queue = queue.Queue()
        self._reader_running = False
        self._reader_enabled = True        # Falseでリーダー一時停止
        self._need_recv_header = False     # 送信後、最初の受信行前に[RECV]挿入
        self._reader_thread = None

        self._build_ui()
        self._start_reader()

    def _build_ui(self):
        """UI構築"""
        # 色定義（DataGen1=青系、DataGen2=オレンジ系）
        self.COLOR_DG1_ACTIVE = "#1976D2"    # 青
        self.COLOR_DG2_ACTIVE = "#F57C00"    # オレンジ
        self.COLOR_DG_INACTIVE = "#E0E0E0"   # グレー

        # 枠線用コンテナ（gridで四方に色付きバーを配置）
        border_container = tk.Frame(self)
        border_container.pack(fill="both", expand=True, padx=5, pady=5)
        border_container.grid_rowconfigure(1, weight=1)
        border_container.grid_columnconfigure(1, weight=1)

        # 上バー
        self.border_top = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, height=4)
        self.border_top.grid(row=0, column=0, columnspan=3, sticky="ew")

        # 左バー
        self.border_left = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, width=4)
        self.border_left.grid(row=1, column=0, sticky="ns")

        # 右バー
        self.border_right = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, width=4)
        self.border_right.grid(row=1, column=2, sticky="ns")

        # 下バー
        self.border_bottom = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, height=4)
        self.border_bottom.grid(row=2, column=0, columnspan=3, sticky="ew")

        # 内側コンテンツエリア
        inner = ttk.Frame(border_container)
        inner.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        inner.grid_columnconfigure(0, weight=1)

        # === DataGen切り替え ===
        dg_switch_frame = ttk.Frame(inner)
        dg_switch_frame.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="w")

        self.var_current_dg = tk.IntVar(value=1)
        self.btn_dg1 = tk.Button(dg_switch_frame, text="DataGen1", width=10,
                                  font=("Arial", 9, "bold"),
                                  bg=self.COLOR_DG1_ACTIVE, fg="white",
                                  relief="flat", cursor="hand2",
                                  command=lambda: self._switch_datagen(1))
        self.btn_dg1.pack(side="left", padx=(0, 5))

        self.btn_dg2 = tk.Button(dg_switch_frame, text="DataGen2", width=10,
                                  font=("Arial", 9, "bold"),
                                  bg=self.COLOR_DG_INACTIVE, fg="black",
                                  relief="flat", cursor="hand2",
                                  command=lambda: self._switch_datagen(2))
        self.btn_dg2.pack(side="left")

        # DataGen2が無い場合はボタンを無効化
        if not self.datagen2:
            self.btn_dg2.config(state="disabled", bg="#CCCCCC")

        # 接続状態ラベル
        self.var_dg_status = tk.StringVar(value="未接続")
        self.lbl_dg_status = ttk.Label(dg_switch_frame, textvariable=self.var_dg_status,
                                        font=("Arial", 9), foreground="gray")
        self.lbl_dg_status.pack(side="left", padx=(15, 0))

        # レスポンス窓ボタン + RECV表示（接続状態の右横）
        ttk.Button(dg_switch_frame, text="レスポンス窓",
                   command=self._show_response_window).pack(side="left", padx=(10, 0))
        self.var_show_recv = tk.BooleanVar(value=False)
        ttk.Checkbutton(dg_switch_frame, text="RECV表示",
                        variable=self.var_show_recv).pack(side="left", padx=(5, 0))

        # CSTM/CMDボタン
        ttk.Button(dg_switch_frame, text="CSTM/CMD",
                   command=self.open_cstm_window).pack(side="left", padx=(10, 0))

        # === コネクタ設定（表形式）===
        conn_frame = ttk.LabelFrame(inner, text="コネクタ設定", padding=3)
        conn_frame.grid(row=1, column=0, padx=5, pady=(5, 2), sticky="ew")

        # 表形式
        conn_table = ttk.Frame(conn_frame)
        conn_table.pack(fill="x")

        # ヘッダー行（問合せボタン + 列タイトル）
        ttk.Button(conn_table, text="問合せ", command=self._query_connector_settings, width=10).grid(row=0, column=0, sticky="w")
        ttk.Label(conn_table, text="対象", font=("Arial", 9, "bold"), width=10).grid(row=0, column=1, sticky="w")
        ttk.Label(conn_table, text="FUNC", font=("Arial", 9, "bold"), width=10).grid(row=0, column=2, sticky="w")
        ttk.Label(conn_table, text="INV", font=("Arial", 9, "bold"), width=12).grid(row=0, column=3, sticky="w")

        # CI(DEF)行
        ttk.Label(conn_table, text="CI(DEF)", font=("Arial", 9), width=10).grid(row=1, column=0, sticky="w")
        self.var_cmode_ci = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_cmode_ci, font=("Arial", 9, "bold"), width=10).grid(row=1, column=1, sticky="w")
        self.var_func_ci = tk.StringVar(value="---")
        self.lbl_func_ci = tk.Label(conn_table, textvariable=self.var_func_ci, font=("Arial", 9, "bold"), width=10, anchor="w")
        self.lbl_func_ci.grid(row=1, column=2, sticky="w")
        self.var_inv_ci = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_inv_ci, font=("Arial", 9, "bold"), width=12).grid(row=1, column=3, sticky="w")

        # CII(INV/LBC)行
        ttk.Label(conn_table, text="CII(INV/LBC)", font=("Arial", 9), width=10).grid(row=2, column=0, sticky="w")
        self.var_cmode_cii = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_cmode_cii, font=("Arial", 9, "bold"), width=10).grid(row=2, column=1, sticky="w")
        self.var_func_cii = tk.StringVar(value="---")
        self.lbl_func_cii = tk.Label(conn_table, textvariable=self.var_func_cii, font=("Arial", 9, "bold"), width=10, anchor="w")
        self.lbl_func_cii.grid(row=2, column=2, sticky="w")
        self.var_inv_cii = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_inv_cii, font=("Arial", 9, "bold"), width=12).grid(row=2, column=3, sticky="w")

        # 設定行
        conn_set_row = ttk.Frame(conn_frame)
        conn_set_row.pack(fill="x", pady=(5, 0))

        # コネクタ選択
        ttk.Label(conn_set_row, text="コネクタ:").pack(side="left")
        self.var_conn_target = tk.StringVar(value="CI")
        ttk.Combobox(conn_set_row, textvariable=self.var_conn_target,
                     values=["CI", "CII"], state="readonly", width=4).pack(side="left", padx=(2, 10))

        # 対象(CMODE)設定
        ttk.Label(conn_set_row, text="対象:").pack(side="left")
        self.var_cmode_set = tk.StringVar(value="397PN")
        ttk.Combobox(conn_set_row, textvariable=self.var_cmode_set,
                     values=["397", "397PN", "397LBC", "398", "OS"], state="readonly", width=8).pack(side="left", padx=(2, 5))
        ttk.Button(conn_set_row, text="設定", command=self._send_cmode, width=5).pack(side="left", padx=(0, 15))

        # INV設定
        ttk.Label(conn_set_row, text="INV:").pack(side="left")
        self.var_inv_set = tk.StringVar(value="OFF")
        ttk.Combobox(conn_set_row, textvariable=self.var_inv_set,
                     values=["ON", "OFF"], state="readonly", width=5).pack(side="left", padx=(2, 5))
        ttk.Button(conn_set_row, text="設定", command=self._send_inv, width=5).pack(side="left")

        # === FUNC切替用変数（排他制御）===
        self.var_func_alt = tk.BooleanVar(value=True)
        self.var_func_rndm = tk.BooleanVar(value=False)
        self.var_func_rmp = tk.BooleanVar(value=False)
        self.var_rmp_one_side = tk.BooleanVar(value=True)
        self.var_func_cstm = tk.BooleanVar(value=False)

        # カスタムパターン送信用
        self._cstm_thread = None
        self._cstm_running = False

        # === Pattern送信設定 ===
        pattern_label = ttk.Frame(inner)
        ttk.Checkbutton(pattern_label, text="", variable=self.var_func_alt,
                        command=lambda: self._on_func_change("ALT")).pack(side="left")
        tk.Label(pattern_label, text="2値振り設定", fg="#0066CC", font=("Arial", 9, "bold")).pack(side="left")

        self.pattern_frame = ttk.LabelFrame(inner, labelwidget=pattern_label, padding=3)
        self.pattern_frame.grid(row=2, column=0, padx=5, pady=(5, 0), sticky="ew")
        pattern_frame = self.pattern_frame

        # モード選択行
        self.var_mode = tk.StringVar(value="Position")
        mode_frame = ttk.Frame(pattern_frame)
        mode_frame.grid(row=0, column=0, columnspan=7, sticky="w", padx=2, pady=(2, 5))

        ttk.Label(mode_frame, text="対象").pack(side="left")
        ttk.Radiobutton(mode_frame, text="Position", variable=self.var_mode,
                        value="Position", command=self._on_mode_change).pack(side="left", padx=(10, 0))
        ttk.Radiobutton(mode_frame, text="LBC", variable=self.var_mode,
                        value="LBC", command=self._on_mode_change).pack(side="left", padx=(5, 0))

        # 振幅
        self.var_amp = tk.StringVar(value="静止")
        ttk.Label(pattern_frame, text="振幅").grid(row=1, column=0, sticky="w", padx=2)
        self.amp_combo = ttk.Combobox(pattern_frame, textvariable=self.var_amp, state="readonly", width=12)
        self.amp_combo['values'] = ["1/32FS", "FS", "MajorCarry", "静止", "グリッチ"]
        self.amp_combo.grid(row=1, column=1, sticky="w", padx=2)
        self.amp_combo.bind("<<ComboboxSelected>>", self._on_amp_change)

        # 中心
        self.var_center = tk.StringVar(value="0V")
        ttk.Label(pattern_frame, text="中心").grid(row=1, column=2, sticky="w", padx=2)
        self.center_combo = ttk.Combobox(pattern_frame, textvariable=self.var_center, state="readonly", width=10)
        self.center_combo['values'] = ["0V", "+160V", "-160V"]
        self.center_combo.grid(row=1, column=3, sticky="w", padx=2)

        # 方向
        self.var_dir = tk.StringVar(value="up")
        ttk.Label(pattern_frame, text="方向").grid(row=2, column=0, sticky="w", padx=2)
        self.dir_combo = ttk.Combobox(pattern_frame, textvariable=self.var_dir,
                                       values=["up", "down"], state="readonly", width=12)
        self.dir_combo.grid(row=2, column=1, sticky="w", padx=2)

        # 極性
        self.var_pol = tk.StringVar(value="pos")
        ttk.Label(pattern_frame, text="極性").grid(row=2, column=2, sticky="w", padx=2)
        self.pol_combo = ttk.Combobox(pattern_frame, textvariable=self.var_pol,
                                       values=["pos", "neg"], state="readonly", width=10)
        self.pol_combo.grid(row=2, column=3, sticky="w", padx=2)

        # グリッチ設定
        self.var_glitch_sec = tk.StringVar(value="1.0")
        ttk.Label(pattern_frame, text="グリッチ(sec)").grid(row=1, column=4, sticky="w", padx=(10, 2))
        self.ent_glitch = ttk.Entry(pattern_frame, textvariable=self.var_glitch_sec, width=8, state="disabled")
        self.ent_glitch.grid(row=1, column=5, sticky="w", padx=2)

        self.var_glitch_status = tk.StringVar(value="")
        ttk.Label(pattern_frame, textvariable=self.var_glitch_status).grid(row=1, column=6, sticky="w", padx=(8, 0))

        # グリッチ制御ボタン
        glitch_btn_frame = ttk.Frame(pattern_frame)
        glitch_btn_frame.grid(row=2, column=4, columnspan=3, sticky="w", padx=(10, 0))

        self.btn_glitch_pause = ttk.Button(glitch_btn_frame, text="中断", width=6, state="disabled",
                                            command=self._glitch_pause)
        self.btn_glitch_resume = ttk.Button(glitch_btn_frame, text="再開", width=6, state="disabled",
                                             command=self._glitch_resume)
        self.btn_glitch_stop = ttk.Button(glitch_btn_frame, text="終了", width=6, state="disabled",
                                           command=self._glitch_stop)
        self.btn_glitch_pause.grid(row=0, column=0, padx=0)
        self.btn_glitch_resume.grid(row=0, column=1, padx=0)
        self.btn_glitch_stop.grid(row=0, column=2, padx=(2, 0))

        # ボタン行
        btn_frame = ttk.Frame(pattern_frame)
        btn_frame.grid(row=3, column=0, columnspan=7, pady=(2, 0), sticky="w")

        ttk.Button(btn_frame, text="パターン送信", command=self._send_pattern).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="A固定", command=self._send_hold_a).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="交互出力", command=self._start_alternating).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="イニシャル送信", command=self._send_init).pack(side="left", padx=5)

        # ALT設定表示（表形式）
        alt_table = ttk.Frame(pattern_frame)
        alt_table.grid(row=4, column=0, columnspan=7, pady=(3, 0), sticky="w")

        ttk.Button(alt_table, text="問合せ", command=self._query_alt, width=8).grid(row=0, column=0, sticky="w")
        ttk.Label(alt_table, text="A", font=("Arial", 8, "bold"), width=7).grid(row=0, column=1, sticky="w")
        ttk.Label(alt_table, text="B", font=("Arial", 8, "bold"), width=7).grid(row=0, column=2, sticky="w")
        ttk.Label(alt_table, text="出力", font=("Arial", 8, "bold"), width=6).grid(row=0, column=3, sticky="w")

        ttk.Label(alt_table, text="CI(DEF)", font=("Arial", 9), width=10).grid(row=1, column=0, sticky="w")
        self.var_alt_ci_a = tk.StringVar(value="---")
        ttk.Label(alt_table, textvariable=self.var_alt_ci_a, font=("Arial", 9, "bold"), width=7).grid(row=1, column=1, sticky="w")
        self.var_alt_ci_b = tk.StringVar(value="---")
        ttk.Label(alt_table, textvariable=self.var_alt_ci_b, font=("Arial", 9, "bold"), width=7).grid(row=1, column=2, sticky="w")
        self.var_alt_ci_sel = tk.StringVar(value="---")
        self.lbl_alt_ci_sel = tk.Label(alt_table, textvariable=self.var_alt_ci_sel, font=("Arial", 9, "bold"), width=6, anchor="w")
        self.lbl_alt_ci_sel.grid(row=1, column=3, sticky="w")

        ttk.Label(alt_table, text="CII(INV/LBC)", font=("Arial", 9), width=10).grid(row=2, column=0, sticky="w")
        self.var_alt_cii_a = tk.StringVar(value="---")
        ttk.Label(alt_table, textvariable=self.var_alt_cii_a, font=("Arial", 9, "bold"), width=7).grid(row=2, column=1, sticky="w")
        self.var_alt_cii_b = tk.StringVar(value="---")
        ttk.Label(alt_table, textvariable=self.var_alt_cii_b, font=("Arial", 9, "bold"), width=7).grid(row=2, column=2, sticky="w")
        self.var_alt_cii_sel = tk.StringVar(value="---")
        self.lbl_alt_cii_sel = tk.Label(alt_table, textvariable=self.var_alt_cii_sel, font=("Arial", 9, "bold"), width=6, anchor="w")
        self.lbl_alt_cii_sel.grid(row=2, column=3, sticky="w")

        # === ランダム設定 ===
        rndm_label = ttk.Frame(inner)
        ttk.Checkbutton(rndm_label, text="", variable=self.var_func_rndm,
                        command=lambda: self._on_func_change("RNDM")).pack(side="left")
        tk.Label(rndm_label, text="ランダム設定", fg="#CC0000", font=("Arial", 9, "bold")).pack(side="left")

        rndm_frame = ttk.LabelFrame(inner, labelwidget=rndm_label, padding=3)
        rndm_frame.grid(row=3, column=0, padx=5, pady=0, sticky="ew")

        rndm_target_row = ttk.Frame(rndm_frame)
        rndm_target_row.pack(fill="x", pady=(0, 2))
        ttk.Label(rndm_target_row, text="対象:").pack(side="left")
        self.var_rndm_target = tk.StringVar(value="Positionのみ")
        self.cmb_rndm_target = ttk.Combobox(rndm_target_row, textvariable=self.var_rndm_target,
                                             values=["Positionのみ", "LBCのみ", "Position/LBC両方"],
                                             state="readonly", width=14)
        self.cmb_rndm_target.pack(side="left", padx=(5, 0))

        rndm_set_row = ttk.Frame(rndm_frame)
        rndm_set_row.pack(fill="x", pady=(0, 2))
        ttk.Label(rndm_set_row, text="ビット数:").pack(side="left")
        self.var_rndm_bit = tk.StringVar(value="20")
        self.cmb_rndm_bit = ttk.Combobox(rndm_set_row, textvariable=self.var_rndm_bit,
                                          values=["20", "16", "14", "14H", "14L"],
                                          state="readonly", width=5)
        self.cmb_rndm_bit.pack(side="left", padx=(5, 10))
        ttk.Label(rndm_set_row, text="コネクタ:").pack(side="left")
        self.var_rndm_conn = tk.StringVar(value="CI")
        self.cmb_rndm_conn = ttk.Combobox(rndm_set_row, textvariable=self.var_rndm_conn,
                                           values=["CI", "CII"], state="readonly", width=4)
        self.cmb_rndm_conn.pack(side="left", padx=(5, 10))
        ttk.Button(rndm_set_row, text="設定", command=self._send_rndm).pack(side="left", padx=(10, 0))

        rndm_disp_row = ttk.Frame(rndm_frame)
        rndm_disp_row.pack(fill="x")
        ttk.Button(rndm_disp_row, text="問合せ", command=self._query_rndm).pack(side="left")
        ttk.Label(rndm_disp_row, text="CI(DEF):").pack(side="left", padx=(10, 0))
        self.var_rndm_ci = tk.StringVar(value="---")
        ttk.Label(rndm_disp_row, textvariable=self.var_rndm_ci, font=("Arial", 9, "bold"), width=14).pack(side="left", padx=(2, 5))
        ttk.Label(rndm_disp_row, text="CII(INV/LBC):").pack(side="left")
        self.var_rndm_cii = tk.StringVar(value="---")
        ttk.Label(rndm_disp_row, textvariable=self.var_rndm_cii, font=("Arial", 9, "bold"), width=14).pack(side="left", padx=(2, 0))

        # === ランプ設定 ===
        ramp_label = ttk.Frame(inner)
        ttk.Checkbutton(ramp_label, text="", variable=self.var_func_rmp,
                        command=lambda: self._on_func_change("RMP")).pack(side="left")
        tk.Label(ramp_label, text="ランプ設定", fg="#009900", font=("Arial", 9, "bold")).pack(side="left")

        ramp_frame = ttk.LabelFrame(inner, labelwidget=ramp_label, padding=3)
        ramp_frame.grid(row=4, column=0, padx=5, pady=0, sticky="ew")

        ramp_target_row = ttk.Frame(ramp_frame)
        ramp_target_row.pack(fill="x", pady=(0, 2))
        ttk.Label(ramp_target_row, text="対象:").pack(side="left")
        self.var_rmp_target = tk.StringVar(value="Positionのみ")
        ttk.Combobox(ramp_target_row, textvariable=self.var_rmp_target,
                     values=["Positionのみ", "LBCのみ", "Position/LBC両方"],
                     state="readonly", width=14).pack(side="left", padx=(5, 0))

        ramp_set_row = ttk.Frame(ramp_frame)
        ramp_set_row.pack(fill="x", pady=(0, 5))
        ttk.Label(ramp_set_row, text="コネクタ:").pack(side="left")
        self.var_rmp_conn = tk.StringVar(value="CI")
        ttk.Combobox(ramp_set_row, textvariable=self.var_rmp_conn,
                     values=["CI", "CII"], state="readonly", width=4).pack(side="left", padx=(2, 10))
        ttk.Label(ramp_set_row, text="BGN:").pack(side="left")
        self.var_rmp_set_bgn = tk.StringVar(value="00000")
        ttk.Entry(ramp_set_row, textvariable=self.var_rmp_set_bgn, width=7).pack(side="left", padx=(2, 10))
        ttk.Label(ramp_set_row, text="END:").pack(side="left")
        self.var_rmp_set_end = tk.StringVar(value="FFFFF")
        ttk.Entry(ramp_set_row, textvariable=self.var_rmp_set_end, width=7).pack(side="left", padx=(2, 10))
        ttk.Label(ramp_set_row, text="STEP:").pack(side="left")
        self.var_rmp_set_step = tk.StringVar(value="00001")
        ttk.Entry(ramp_set_row, textvariable=self.var_rmp_set_step, width=7).pack(side="left", padx=(2, 10))
        ttk.Button(ramp_set_row, text="設定", command=self._send_rmp).pack(side="left")

        ramp_table = ttk.Frame(ramp_frame)
        ramp_table.pack(fill="x")
        ttk.Button(ramp_table, text="問合せ", command=self._query_rmp, width=8).grid(row=0, column=0, sticky="w")
        ttk.Label(ramp_table, text="BEGIN", font=("Arial", 8, "bold"), width=8).grid(row=0, column=1, sticky="w")
        ttk.Label(ramp_table, text="END", font=("Arial", 8, "bold"), width=8).grid(row=0, column=2, sticky="w")
        ttk.Label(ramp_table, text="STEP", font=("Arial", 8, "bold"), width=8).grid(row=0, column=3, sticky="w")

        ttk.Label(ramp_table, text="CI(DEF)", font=("Arial", 9), width=10).grid(row=1, column=0, sticky="w")
        self.var_rmp_ci_bgn = tk.StringVar(value="---")
        ttk.Label(ramp_table, textvariable=self.var_rmp_ci_bgn, font=("Arial", 9, "bold"), width=8).grid(row=1, column=1, sticky="w")
        self.var_rmp_ci_end = tk.StringVar(value="---")
        ttk.Label(ramp_table, textvariable=self.var_rmp_ci_end, font=("Arial", 9, "bold"), width=8).grid(row=1, column=2, sticky="w")
        self.var_rmp_ci_step = tk.StringVar(value="---")
        ttk.Label(ramp_table, textvariable=self.var_rmp_ci_step, font=("Arial", 9, "bold"), width=8).grid(row=1, column=3, sticky="w")

        ttk.Label(ramp_table, text="CII(INV/LBC)", font=("Arial", 9), width=10).grid(row=2, column=0, sticky="w")
        self.var_rmp_cii_bgn = tk.StringVar(value="---")
        ttk.Label(ramp_table, textvariable=self.var_rmp_cii_bgn, font=("Arial", 9, "bold"), width=8).grid(row=2, column=1, sticky="w")
        self.var_rmp_cii_end = tk.StringVar(value="---")
        ttk.Label(ramp_table, textvariable=self.var_rmp_cii_end, font=("Arial", 9, "bold"), width=8).grid(row=2, column=2, sticky="w")
        self.var_rmp_cii_step = tk.StringVar(value="---")
        ttk.Label(ramp_table, textvariable=self.var_rmp_cii_step, font=("Arial", 9, "bold"), width=8).grid(row=2, column=3, sticky="w")

        # === カスタムパターン変数（別ウィンドウで使用） ===
        config = load_config()
        self.var_cstm_pos_path = tk.StringVar(value=config.get("datagen_cstm_pos_path", ""))
        self.var_cstm_neg_path = tk.StringVar(value=config.get("datagen_cstm_neg_path", ""))
        self.var_cstm_auto_start = tk.BooleanVar(value=True)
        self.var_cstm_conn = tk.StringVar(value="Positionのみ")
        self.var_cstm_progress = tk.StringVar(value="0/16374 (0%)")
        self.btn_cstm_send = None
        self.btn_cstm_stop = None
        self._cstm_window = None
        self._ent_rate_value = None
        self._cmb_rate_unit = None
        self._btn_rate_set = None
        self._btn_rate_query = None

        # === 手動コマンド・レート設定（変数のみ、UIはCSTM/CMDウィンドウに配置）===
        config = load_config()
        last_cmd = config.get("datagen_manual_cmd", "")
        self.var_manual_cmd = tk.StringVar(value=last_cmd)
        self.var_rate_value = tk.StringVar(value="10")
        self.var_rate_unit = tk.StringVar(value="msec")
        self.var_rate_display = tk.StringVar(value="---")

        # ランダムRate変更
        self.var_rate_random = tk.BooleanVar(value=False)
        self._rate_random_job = None

        # レスポンスウィンドウ用（DataGen1/2それぞれ別）
        self.response_windows = {1: None, 2: None}
        self.response_areas = {1: None, 2: None}

        # 起動時のデフォルト状態を適用（静止→グレースケール）
        self._on_amp_change()

    # ========== イベントハンドラ ==========
    def _on_func_change(self, func_type):
        """FUNC切替（ALT/RNDM/RMP/CSTM）- 排他制御とコマンド送信"""
        if func_type == "ALT":
            self.var_func_alt.set(True); self.var_func_rndm.set(False); self.var_func_rmp.set(False); self.var_func_cstm.set(False)
        elif func_type == "RNDM":
            self.var_func_alt.set(False); self.var_func_rndm.set(True); self.var_func_rmp.set(False); self.var_func_cstm.set(False)
        elif func_type == "RMP":
            self.var_func_alt.set(False); self.var_func_rndm.set(False); self.var_func_rmp.set(True); self.var_func_cstm.set(False)
        elif func_type == "CSTM":
            self.var_func_alt.set(False); self.var_func_rndm.set(False); self.var_func_rmp.set(False); self.var_func_cstm.set(True)

        if not self.datagen or not self.datagen.is_connected():
            return

        if func_type == "CSTM":
            cstm_mode = self.var_cstm_conn.get()
            self._append_log(""); self._append_log(f"【FUNC切替: {func_type} ({cstm_mode})】")
            self._send_and_log("gen stop")
            if cstm_mode == "Positionのみ":
                self._send_and_log("func cstm ci"); self._send_and_log("func alt cii")
                for cmd in ["alt a 80000 cii p", "alt b 80000 cii p", "alt a 80000 cii n", "alt b 80000 cii n"]:
                    self._send_and_log(cmd)
                self._send_and_log("alt s sa cii")
            elif cstm_mode == "LBCのみ":
                self._send_and_log("func alt ci")
                for cmd in ["alt a 80000 ci p", "alt b 80000 ci p", "alt a 80000 ci n", "alt b 80000 ci n"]:
                    self._send_and_log(cmd)
                self._send_and_log("alt s sa ci"); self._send_and_log("func cstm cii")
            else:
                self._send_and_log("func cstm ci"); self._send_and_log("func cstm cii")
            self._send_and_log("gen start")
        elif func_type == "RMP":
            target = self.var_rmp_target.get()
            self._append_log(""); self._append_log(f"【FUNC切替: {func_type} ({target})】")
            self._send_and_log("gen stop")
            if target == "Positionのみ":
                self._send_and_log("func rmp ci"); self._send_and_log("func alt cii")
                for cmd in ["alt a 80000 cii p", "alt b 80000 cii p", "alt a 80000 cii n", "alt b 80000 cii n"]:
                    self._send_and_log(cmd)
                self._send_and_log("alt s sa cii")
            elif target == "LBCのみ":
                self._send_and_log("func rmp cii"); self._send_and_log("func alt ci")
                for cmd in ["alt a 80000 ci p", "alt b 80000 ci p", "alt a 80000 ci n", "alt b 80000 ci n"]:
                    self._send_and_log(cmd)
                self._send_and_log("alt s sa ci")
            else:
                self._send_and_log("func rmp")
            self._send_and_log("gen start")
        elif func_type == "RNDM":
            target = self.var_rndm_target.get()
            self._append_log(""); self._append_log(f"【FUNC切替: {func_type} ({target})】")
            self._send_and_log("gen stop")
            if target == "Positionのみ":
                self._send_and_log("func rndm ci"); self._send_and_log("func alt cii")
                for cmd in ["alt a 80000 cii p", "alt b 80000 cii p", "alt a 80000 cii n", "alt b 80000 cii n"]:
                    self._send_and_log(cmd)
                self._send_and_log("alt s sa cii")
            elif target == "LBCのみ":
                self._send_and_log("func rndm cii"); self._send_and_log("func alt ci")
                for cmd in ["alt a 80000 ci p", "alt b 80000 ci p", "alt a 80000 ci n", "alt b 80000 ci n"]:
                    self._send_and_log(cmd)
                self._send_and_log("alt s sa ci")
            else:
                self._send_and_log("func rndm")
            self._send_and_log("gen start")
        else:
            self._append_log(""); self._append_log(f"【FUNC切替: {func_type}】")
            self._send_and_log("gen stop"); self._send_and_log(f"func {func_type.lower()}"); self._send_and_log("gen start")

        self.after(100, self._query_connector_settings)
        if func_type == "ALT":
            self.var_amp.set("静止"); self._on_amp_change(); self.after(200, self._send_pattern)
        elif func_type == "RNDM":
            self.after(200, self._query_rndm)
        elif func_type == "RMP":
            self.after(200, self._query_rmp)
        self.after(300, self._query_alt)

    def _on_rmp_one_side_change(self):
        if self.var_func_rmp.get():
            self._on_func_change("RMP")

    def _on_mode_change(self):
        mode = self.var_mode.get()
        if mode == "LBC":
            self.amp_combo['values'] = ["FS", "LBCグリッチ", "静止"]
            if self.var_amp.get() not in ["FS", "LBCグリッチ", "静止"]:
                self.var_amp.set("FS")
            self.center_combo.set(""); self.center_combo.config(state="disabled")
        else:
            self.amp_combo['values'] = ["1/32FS", "FS", "MajorCarry", "静止", "グリッチ"]
            if self.var_amp.get() not in ["1/32FS", "FS", "MajorCarry", "静止", "グリッチ"]:
                self.var_amp.set("1/32FS")
            self.center_combo.config(state="readonly")
            if not self.var_center.get():
                self.var_center.set("0V")
        self._on_amp_change()

    def _on_amp_change(self, event=None):
        amp = self.var_amp.get()
        mode = self.var_mode.get()
        self.ent_glitch.config(state="disabled")
        self.btn_glitch_pause.config(state="disabled"); self.btn_glitch_resume.config(state="disabled"); self.btn_glitch_stop.config(state="disabled")
        self.var_glitch_status.set("")
        if mode == "Position":
            self.center_combo.config(state="readonly")
        else:
            self.center_combo.set(""); self.center_combo.config(state="disabled")

        if amp == "FS":
            if mode == "Position":
                self.center_combo['values'] = ["±160V"]; self.center_combo.set("±160V")
            if not self.var_dir.get(): self.var_dir.set("up")
            if not self.var_pol.get(): self.var_pol.set("pos")
            self.dir_combo.config(state="readonly"); self.pol_combo.config(state="readonly")
        elif amp == "1/32FS":
            if mode == "Position":
                self.center_combo['values'] = ["0V", "+160V", "-160V"]
                if self.var_center.get() not in ["0V", "+160V", "-160V"]: self.var_center.set("0V")
            if not self.var_dir.get(): self.var_dir.set("up")
            if not self.var_pol.get(): self.var_pol.set("pos")
            self.dir_combo.config(state="readonly"); self.pol_combo.config(state="readonly")
        elif amp == "MajorCarry":
            self.center_combo.set(""); self.center_combo.config(state="disabled")
            if not self.var_dir.get(): self.var_dir.set("up")
            if not self.var_pol.get(): self.var_pol.set("pos")
            self.dir_combo.config(state="readonly"); self.pol_combo.config(state="readonly")
        elif amp == "静止":
            self.center_combo.set(""); self.center_combo.config(state="disabled")
            self.var_dir.set(""); self.dir_combo.set(""); self.dir_combo.config(state="disabled")
            self.var_pol.set(""); self.pol_combo.set(""); self.pol_combo.config(state="disabled")
        elif amp == "グリッチ":
            self.center_combo.set(""); self.center_combo.config(state="disabled")
            if not self.var_dir.get(): self.var_dir.set("up")
            if not self.var_pol.get(): self.var_pol.set("pos")
            self.dir_combo.config(state="readonly"); self.pol_combo.config(state="readonly")
            self.ent_glitch.config(state="normal")
        elif amp == "LBCグリッチ":
            self.center_combo.set(""); self.center_combo.config(state="disabled")
            self.dir_combo.set(""); self.dir_combo.config(state="disabled")
            self.pol_combo.set(""); self.pol_combo.config(state="disabled")

    # ========== DataGen切り替え ==========
    def _switch_datagen(self, dg_num):
        if dg_num == self.current_dg: return
        if dg_num == 2 and not self.datagen2: return
        self._reader_enabled = False; time.sleep(0.02)
        self.current_dg = dg_num; self.var_current_dg.set(dg_num)
        if dg_num == 1:
            self.datagen = self.datagen1
            self.btn_dg1.config(bg=self.COLOR_DG1_ACTIVE, fg="white"); self.btn_dg2.config(bg=self.COLOR_DG_INACTIVE, fg="black")
            border_color = self.COLOR_DG1_ACTIVE
        else:
            self.datagen = self.datagen2
            self.btn_dg1.config(bg=self.COLOR_DG_INACTIVE, fg="black"); self.btn_dg2.config(bg=self.COLOR_DG2_ACTIVE, fg="white")
            border_color = self.COLOR_DG2_ACTIVE
        for b in [self.border_top, self.border_left, self.border_right, self.border_bottom]:
            b.config(bg=border_color)
        self._update_dg_status()
        if self.datagen and self.datagen.is_connected(): self.datagen.flush_input()
        self._need_recv_header = False; self._reader_enabled = True
        self._append_log(""); self._append_log(f"【DataGen{dg_num} に切り替え】")
        if self.datagen and self.datagen.is_connected(): self.after(100, self._query_connector_settings)

    def _update_dg_status(self):
        if self.datagen and self.datagen.is_connected():
            port = getattr(self.datagen.ser, 'port', '') if self.datagen.ser else ''
            self.var_dg_status.set(f"{port} 接続中"); self.lbl_dg_status.config(foreground="green")
        else:
            self.var_dg_status.set("未接続"); self.lbl_dg_status.config(foreground="gray")

    # ========== レスポンス表示ウィンドウ ==========
    def _show_response_window(self):
        dg = self.current_dg
        if self.response_windows[dg] and self.response_windows[dg].winfo_exists():
            self.response_windows[dg].lift(); self.response_windows[dg].focus_force(); return
        border_color = self.COLOR_DG1_ACTIVE if dg == 1 else self.COLOR_DG2_ACTIVE
        win = tk.Toplevel(self)
        win.title(f"DataGen{dg} レスポンス表示")
        menubar = tk.Menu(win); win.config(menu=menubar)
        main_win = self.winfo_toplevel(); main_win.update_idletasks()
        main_x, main_y = main_win.winfo_x(), main_win.winfo_y()
        main_width, main_height = main_win.winfo_width(), main_win.winfo_height()
        x_offset = main_x + main_width + 10 + (410 if dg == 2 else 0)
        win.geometry(f"400x{main_height}+{x_offset}+{main_y}"); win.resizable(True, True)
        bc = tk.Frame(win); bc.pack(fill="both", expand=True)
        bc.grid_rowconfigure(1, weight=1); bc.grid_columnconfigure(1, weight=1)
        tk.Frame(bc, bg=border_color, height=4).grid(row=0, column=0, columnspan=3, sticky="ew")
        tk.Frame(bc, bg=border_color, width=4).grid(row=1, column=0, sticky="ns")
        tk.Frame(bc, bg=border_color, width=4).grid(row=1, column=2, sticky="ns")
        tk.Frame(bc, bg=border_color, height=4).grid(row=2, column=0, columnspan=3, sticky="ew")
        inner = ttk.Frame(bc); inner.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        response_area = ScrolledText(inner, font=("Consolas", 10), wrap="none"); response_area.pack(fill="both", expand=True)
        bf = ttk.Frame(inner); bf.pack(fill="x", pady=(5, 0))
        ttk.Button(bf, text="クリア", command=lambda ra=response_area: ra.delete(1.0, tk.END)).pack(side="right")
        self.response_windows[dg] = win; self.response_areas[dg] = response_area
        def on_close(d=dg):
            self.response_windows[d] = None; self.response_areas[d] = None; win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

    # ========== 常駐リーダー ==========
    def _start_reader(self):
        if self._reader_running: return
        self._reader_running = True
        self._reader_thread = threading.Thread(target=self._reader_loop, daemon=True); self._reader_thread.start()
        self._poll_text_queue()

    def _stop_reader(self):
        self._reader_running = False
        if self._reader_thread and self._reader_thread.is_alive():
            try: self._reader_thread.join(timeout=0.5)
            except Exception: pass
        self._reader_thread = None

    def _reader_loop(self):
        line_buffer = ""
        while self._reader_running:
            if not self._reader_enabled: time.sleep(0.05); continue
            mgr = self.datagen
            if not mgr or not mgr.is_connected(): time.sleep(0.1); continue
            chunk = self._read_serial_chunk(mgr)
            if not chunk: time.sleep(0.01); continue
            for ch in chunk:
                if ch in ("\r", "\n"):
                    if line_buffer.strip():
                        if self._need_recv_header:
                            self._append_text("[RECV]"); self._need_recv_header = False
                        self._append_text(line_buffer)
                    line_buffer = ""
                elif ch == ">" and not line_buffer.strip():
                    line_buffer = ""
                else:
                    line_buffer += ch

    def _read_serial_chunk(self, mgr):
        with mgr.lock:
            if mgr.is_connected():
                n = mgr.ser.in_waiting
                if n > 0:
                    try: return mgr.ser.read(n).decode("utf-8", errors="ignore")
                    except Exception: return ""
        return ""

    def _append_text(self, message):
        if threading.current_thread() is threading.main_thread():
            ra = self.response_areas.get(self.current_dg)
            if ra: ra.insert(tk.END, message.rstrip() + "\n"); ra.see(tk.END)
        else:
            self._text_queue.put(message)

    def _poll_text_queue(self):
        try:
            for _ in range(200):
                msg = self._text_queue.get_nowait()
                ra = self.response_areas.get(self.current_dg)
                if ra: ra.insert(tk.END, msg.rstrip() + "\n"); ra.see(tk.END)
        except queue.Empty: pass
        self.after(16, self._poll_text_queue)

    def _sync_command(self, cmd, wait_sec=0.05, read_timeout=0.05):
        self._reader_enabled = False; time.sleep(0.02)
        try:
            response = self.datagen.send_command_with_response(cmd, wait_sec=wait_sec, read_timeout=read_timeout)
        finally:
            self._reader_enabled = True
        return response

    # ========== ログ ==========
    def _append_log(self, text):
        if threading.current_thread() is threading.main_thread():
            ra = self.response_areas.get(self.current_dg)
            if ra: ra.insert(tk.END, text.rstrip() + "\n"); ra.see(tk.END)
        else:
            self._text_queue.put(text)

    def _send_and_log(self, cmd, sleep_sec=0.05, add_blank=False):
        if add_blank: self._append_log("")
        self._append_log(f"SEND: {cmd}")
        if self.var_show_recv.get():
            response = self._sync_command(cmd)
            if response: self._append_log(f"RECV:\n{response}")
        else:
            self.datagen.write(f"{cmd}\r".encode("utf-8")); time.sleep(sleep_sec)

    def _send_and_log_thread(self, cmd, sleep_sec=0.05):
        show_recv = self.var_show_recv.get()
        if show_recv:
            response = self._sync_command(cmd)
        else:
            self.datagen.write(f"{cmd}\r".encode("utf-8")); time.sleep(sleep_sec); response = None
        self._append_log(f"SEND: {cmd}")
        if show_recv and response: self._append_log(f"RECV:\n{response}")

    # ========== コマンド送信 ==========
    def _send_manual_command(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        cmd = self.var_manual_cmd.get().strip()
        if not cmd: return
        update_config_value(["datagen_manual_cmd"], cmd)
        self._append_log(""); self._append_log(f"SEND: {cmd}")
        self._need_recv_header = True
        self.datagen.write(f"{cmd}\r".encode("utf-8"))
        self.var_manual_cmd.set("")

    def _set_rate(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        try: value = float(self.var_rate_value.get())
        except ValueError: messagebox.showerror("エラー", "数値を入力してください"); return
        unit = self.var_rate_unit.get()
        if unit == "msec": ns_val = value * 1000000
        elif unit == "μsec": ns_val = value * 1000
        else: ns_val = value
        rate_val = max(3, min(int(ns_val / 10), 1048575))
        cmd = f"rate {rate_val}"
        self._append_log(""); self._append_log(f"SEND: {cmd}")
        response = self._sync_command(cmd)
        if response: self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_rate)

    def _clamp_rate_random_interval(self):
        try:
            val = float(self.var_rate_random_interval.get() or 5)
            if val < 0.1: self.var_rate_random_interval.set("0.1")
        except ValueError: self.var_rate_random_interval.set("5")

    def _toggle_rate_random(self):
        enabled = self.var_rate_random.get()
        # Rate手動設定フィールドの有効/無効切替
        rate_state = "disabled" if enabled else "normal"
        rate_combo_state = "disabled" if enabled else "readonly"
        for w in [self._ent_rate_value, self._btn_rate_set, self._btn_rate_query]:
            if w and w.winfo_exists(): w.config(state=rate_state)
        if self._cmb_rate_unit and self._cmb_rate_unit.winfo_exists():
            self._cmb_rate_unit.config(state=rate_combo_state)
        if enabled: self._rate_random_tick()
        else:
            if self._rate_random_job: self.after_cancel(self._rate_random_job); self._rate_random_job = None

    def _rate_random_tick(self):
        import random, math
        if not self.var_rate_random.get(): self._rate_random_job = None; return
        if not self.datagen or not self.datagen.is_connected():
            self._rate_random_job = None; self.var_rate_random.set(False); return
        log_min, log_max = math.log(3), math.log(1049000)
        rate_val = int(math.exp(random.uniform(log_min, log_max)))
        ns_val = rate_val * 10
        cmd = f"rate {rate_val}"
        self._append_log(f"SEND: {cmd}  [ランダム]"); self._sync_command(cmd)
        if ns_val >= 1000000: disp = f"{ns_val / 1000000:.3f}ms"
        elif ns_val >= 1000: disp = f"{ns_val / 1000:.1f}μs"
        else: disp = f"{ns_val}ns"
        self.var_rate_display.set(disp)
        try: interval_sec = max(0.1, float(self.var_rate_random_interval.get() or 5))
        except ValueError: interval_sec = 5
        self._rate_random_job = self.after(int(interval_sec * 1000), self._rate_random_tick)

    def _send_rndm(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        cmd = f"rndm {self.var_rndm_bit.get()} {self.var_rndm_conn.get()}"
        self._append_log(""); self._append_log(f"SEND: {cmd}")
        response = self._sync_command(cmd)
        if response: self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_rndm)

    def _send_cmode(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        cmd = f"cmode {self.var_cmode_set.get().lower()} {self.var_conn_target.get().lower()}"
        self._append_log(""); self._append_log(f"SEND: {cmd}")
        response = self._sync_command(cmd)
        if response: self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_connector_settings)

    def _send_inv(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        cmd = f"inv {self.var_inv_set.get().lower()} {self.var_conn_target.get().lower()}"
        self._append_log(""); self._append_log(f"SEND: {cmd}")
        response = self._sync_command(cmd)
        if response: self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_connector_settings)

    def _query_connector_settings(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        self._append_log("")
        self._append_log("SEND: cmode"); r = self._sync_command("cmode")
        if r: self._append_log(f"RECV:\n{r}"); self._update_cmode_display(r)
        self._append_log("SEND: inv"); r = self._sync_command("inv")
        if r: self._append_log(f"RECV:\n{r}"); self._update_inv_display(r)
        self._append_log("SEND: func"); r = self._sync_command("func")
        if r: self._append_log(f"RECV:\n{r}"); self._update_func_display(r)

    def _query_alt(self):
        if not self.datagen or not self.datagen.is_connected(): return
        self._append_log(""); self._append_log("SEND: alt")
        r = self._sync_command("alt")
        if r: self._append_log(f"RECV:\n{r}"); self._update_alt_display(r)

    def _query_rndm(self):
        if not self.datagen or not self.datagen.is_connected(): return
        self._append_log(""); self._append_log("SEND: rndm")
        r = self._sync_command("rndm")
        if r: self._append_log(f"RECV:\n{r}"); self._update_rndm_display(r)

    def _send_rmp(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        conn = self.var_rmp_conn.get().lower()
        self._append_log("")
        for label, var in [("bgn", self.var_rmp_set_bgn), ("end", self.var_rmp_set_end), ("step", self.var_rmp_set_step)]:
            cmd = f"rmp {label} {var.get().strip()} {conn}"
            self._append_log(f"SEND: {cmd}")
            r = self._sync_command(cmd)
            if r: self._append_log(f"RECV:\n{r}")
        self.after(100, self._query_rmp); self.after(200, self._query_connector_settings)

    def _query_rmp(self):
        if not self.datagen or not self.datagen.is_connected(): return
        self._append_log(""); self._append_log("SEND: rmp")
        r = self._sync_command("rmp")
        if r: self._append_log(f"RECV:\n{r}"); self._update_rmp_display(r)

    # ========== レスポンスパーサー ==========
    def _update_alt_display(self, response):
        ci_a, ci_b, ci_sel = "---", "---", "---"
        cii_a, cii_b, cii_sel = "---", "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector  I pos"):
                m = re.search(r':\s*(\w+)\s+(\w+)\s+(\w+)', line)
                if m: ci_a, ci_b, ci_sel = m.group(1), m.group(2), m.group(3)
            elif line.startswith("connector II pos"):
                m = re.search(r':\s*(\w+)\s+(\w+)\s+(\w+)', line)
                if m: cii_a, cii_b, cii_sel = m.group(1), m.group(2), m.group(3)
        self.var_alt_ci_a.set(ci_a); self.var_alt_ci_b.set(ci_b)
        self.var_alt_cii_a.set(cii_a); self.var_alt_cii_b.set(cii_b)
        def sel_text(s): return "A固定" if s == "sa" else ("交互" if s == "sab" else s)
        def sel_color(s): return "#CC6600" if s == "sa" else ("#0066CC" if s == "sab" else "#000000")
        self.var_alt_ci_sel.set(sel_text(ci_sel)); self.var_alt_cii_sel.set(sel_text(cii_sel))
        self.lbl_alt_ci_sel.config(fg=sel_color(ci_sel)); self.lbl_alt_cii_sel.config(fg=sel_color(cii_sel))

    def _update_rndm_display(self, response):
        ci_text, cii_text = "---", "---"
        for line in response.split('\n'):
            if line.strip().startswith("connector  I "):
                m = re.search(r':\s*(.+)', line.strip())
                if m: ci_text = m.group(1).replace(" ", ""); break
        for line in response.split('\n'):
            if line.strip().startswith("connector II "):
                m = re.search(r':\s*(.+)', line.strip())
                if m: cii_text = m.group(1).replace(" ", ""); break
        self.var_rndm_ci.set(ci_text); self.var_rndm_cii.set(cii_text)

    def _update_rmp_display(self, response):
        ci_bgn, ci_end, ci_step = "---", "---", "---"
        cii_bgn, cii_end, cii_step = "---", "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector  I pos"):
                m = re.search(r':\s*(\w+)\s+(\w+)\s+(\w+)', line)
                if m: ci_bgn, ci_end, ci_step = m.group(1), m.group(2), m.group(3)
            elif line.startswith("connector II pos"):
                m = re.search(r':\s*(\w+)\s+(\w+)\s+(\w+)', line)
                if m: cii_bgn, cii_end, cii_step = m.group(1), m.group(2), m.group(3)
        self.var_rmp_ci_bgn.set(ci_bgn); self.var_rmp_ci_end.set(ci_end); self.var_rmp_ci_step.set(ci_step)
        self.var_rmp_cii_bgn.set(cii_bgn); self.var_rmp_cii_end.set(cii_end); self.var_rmp_cii_step.set(cii_step)

    def _update_cmode_display(self, response):
        ci_text, cii_text = "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector I "):
                m = re.search(r':\s*(.+)', line)
                if m: ci_text = m.group(1).strip()
            elif line.startswith("connector II "):
                m = re.search(r':\s*(.+)', line)
                if m: cii_text = m.group(1).strip()
        self.var_cmode_ci.set(ci_text); self.var_cmode_cii.set(cii_text)

    def _update_inv_display(self, response):
        ci_text, cii_text = "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector I "):
                m = re.search(r':\s*(.+)', line)
                if m: ci_text = m.group(1).strip()
            elif line.startswith("connector II "):
                m = re.search(r':\s*(.+)', line)
                if m: cii_text = m.group(1).strip()
        self.var_inv_ci.set(ci_text); self.var_inv_cii.set(cii_text)

    def _convert_func_name(self, func_type):
        return "2値振り" if "square" in func_type.lower() else func_type

    def _get_func_color(self, func_type):
        ft = func_type.lower()
        if "square" in ft: return "#0066CC"
        elif "ramp" in ft: return "#009900"
        elif "random" in ft: return "#CC0000"
        elif "bit-check" in ft or "bitcheck" in ft: return "#FF6600"
        elif "custom" in ft: return "#9900CC"
        return "#000000"

    def _update_func_display(self, response):
        ci_p, ci_n, cii_p, cii_n = "---", "---", "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector I pos"):
                m = re.search(r':\s*(.+)', line)
                if m: ci_p = m.group(1).strip()
            elif line.startswith("connector I neg"):
                m = re.search(r':\s*(.+)', line)
                if m: ci_n = m.group(1).strip()
            elif line.startswith("connector II pos"):
                m = re.search(r':\s*(.+)', line)
                if m: cii_p = m.group(1).strip()
            elif line.startswith("connector II neg"):
                m = re.search(r':\s*(.+)', line)
                if m: cii_n = m.group(1).strip()
        warnings = []
        if ci_p != "---" and ci_n != "---" and ci_p != ci_n: warnings.append(f"CI: pos={ci_p}, neg={ci_n}")
        if cii_p != "---" and cii_n != "---" and cii_p != cii_n: warnings.append(f"CII: pos={cii_p}, neg={cii_n}")
        if warnings: messagebox.showwarning("FUNC設定警告", "P/N設定が異なっています:\n" + "\n".join(warnings))
        self.var_func_ci.set(self._convert_func_name(ci_p)); self.var_func_cii.set(self._convert_func_name(cii_p))
        self.lbl_func_ci.config(fg=self._get_func_color(ci_p)); self.lbl_func_cii.config(fg=self._get_func_color(cii_p))

    def _query_rate(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        self._append_log(""); self._append_log("SEND: rate")
        r = self._sync_command("rate")
        if r: self._append_log(f"RECV:\n{r}"); self._update_rate_display(r)

    def _update_rate_display(self, response):
        m = re.search(r'data rate\s*:\s*(\d+)\s*\[ns\]\s*\((\d+)\s*\[Hz\]\)', response)
        if m:
            ns_val, hz_val = int(m.group(1)), int(m.group(2))
            if ns_val >= 1000000: t = f"{ns_val/1000000:.2f}msec"
            elif ns_val >= 1000: t = f"{ns_val/1000:.1f}μsec"
            else: t = f"{ns_val}nsec"
            if hz_val >= 1000000: f = f"{hz_val/1000000:.2f}MHz"
            elif hz_val >= 1000: f = f"{hz_val/1000:.1f}kHz"
            else: f = f"{hz_val}Hz"
            self.var_rate_display.set(f"{t} ({f})"); return
        if "period <=" in response: self.after(100, self._query_rate)

    # ========== イニシャル・パターン送信 ==========
    def _send_init(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        init_commands = [
            "gen stop", "alt s sab ci", "alt s sab cii", "func alt ci", "func alt cii",
            "inv off ci", "inv off cii", "trg mon", "rate 1000000",
            "cmode 397pn ci", "cmode 397lbc cii",
            "stb 397pn ts 5", "stb 397pn tw 4", "stb 397pn th 5", "stb 397pn tpn 10",
            "stb 397lbc ts 8", "stb 397lbc tw 8", "stb 397lbc th 90", "stb 397lbc tpn 0",
            "stb 397 ts 10", "stb 397 tw 8", "stb 397 th 10", "stb 397 tpn 0", "gen start",
        ]
        self._append_log(""); self._append_log(f"【Init送信開始 (DataGen{self.current_dg})】")
        for cmd in init_commands:
            print(f"INIT SEND: {cmd}"); self._send_and_log(cmd, sleep_sec=0.05)
        if self.current_dg == 1: self.initialized = True
        else: self.initialized2 = True
        self._append_log("【Init送信完了】"); self.after(100, self._query_rate)

    def _send_hold_a(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        self._send_and_log("alt s sa", add_blank=True); self.after(100, self._query_alt)

    def _start_alternating(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        self._send_and_log("alt s sab", add_blank=True); self.after(100, self._query_alt)

    def _send_pattern(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        is_initialized = self.initialized if self.current_dg == 1 else self.initialized2
        if not is_initialized: self._send_init()
        amp, center, direction, polarity, mode = self.var_amp.get(), self.var_center.get(), self.var_dir.get(), self.var_pol.get(), self.var_mode.get()
        if mode == "Position" and amp == "グリッチ":
            try:
                sec = float(self.var_glitch_sec.get())
                if sec <= 0: raise ValueError
            except Exception:
                messagebox.showerror("エラー", "グリッチ遷移時間(sec)には正の数値を入力してください。"); return
            self._start_glitch_sequence(direction, polarity, sec); return
        pattern_ci, pattern_cii = self._get_patterns(mode, amp, center, direction, polarity)
        self._append_log("")
        if pattern_ci:
            for cmd in pattern_ci: self._send_and_log(cmd, sleep_sec=0.05)
        if pattern_cii:
            for cmd in pattern_cii: self._send_and_log(cmd, sleep_sec=0.05)
        self._send_and_log("alt s sab"); self.after(100, self._query_alt)

    def _get_patterns(self, mode, amp, center, direction, polarity):
        cii_static = ["alt a 80000 cii p", "alt b 80000 cii p", "alt a 80000 cii n", "alt b 80000 cii n"]
        ci_static = ["alt a 80000 ci p", "alt b 80000 ci p", "alt a 80000 ci n", "alt b 80000 ci n"]
        if amp == "静止": return ci_static, cii_static
        if mode == "LBC":
            if amp == "LBCグリッチ":
                return ci_static, ["alt a 80030 cii p", "alt b 7ffff cii p", "alt a 80030 cii n", "alt b 7ffff cii n"]
            elif amp == "FS":
                if direction == "up":
                    cii = ["alt a fffff cii p", "alt b 00000 cii p", "alt a 00000 cii n", "alt b fffff cii n"]
                else:
                    cii = ["alt a 00000 cii p", "alt b fffff cii p", "alt a fffff cii n", "alt b 00000 cii n"]
                if polarity == "neg": cii = self._flip_polarity(cii)
                return ci_static, cii
        if amp == "FS": return self._build_position_fs(direction, polarity), cii_static
        elif amp == "1/32FS": return self._build_position_132(center, direction, polarity), cii_static
        elif amp == "MajorCarry": return self._build_position_majorcarry(direction, polarity), cii_static
        return ci_static, cii_static

    def _build_position_fs(self, direction, polarity):
        d, p = (direction or "").lower(), (polarity or "").lower()
        a_val, b_val = ("fffff", "00000") if d == "up" else ("00000", "fffff")
        if p == "neg": a_val, b_val = b_val, a_val
        return [f"alt a {a_val} ci p", f"alt b {b_val} ci p", f"alt a {a_val} ci n", f"alt b {b_val} ci n"]

    def _build_position_132(self, center, direction, polarity):
        c, d, p = (center or "").upper(), (direction or "").lower(), (polarity or "").lower()
        if p == "neg":
            c_eff = "-160V" if c in ("+160V", "+FULL") else ("+160V" if c in ("-160V", "-FULL") else "0V")
        else: c_eff = c
        if c_eff in ("+160V", "+FULL"): left, right = "f7fff", "fffff"
        elif c_eff in ("-160V", "-FULL"): left, right = "00000", "08000"
        else: left, right = "7bfff", "84000"
        a_val, b_val = left, right
        if d == "up": a_val, b_val = b_val, a_val
        if p == "neg": a_val, b_val = b_val, a_val
        return [f"alt a {a_val} ci p", f"alt b {b_val} ci p", f"alt a {a_val} ci n", f"alt b {b_val} ci n"]

    def _build_position_majorcarry(self, direction, polarity):
        d, p = (direction or "").lower(), (polarity or "").lower()
        a_val, b_val = "7ffff", "80000"
        if d == "up": a_val, b_val = b_val, a_val
        if p == "neg": a_val, b_val = b_val, a_val
        return [f"alt a {a_val} ci p", f"alt b {b_val} ci p", f"alt a {a_val} ci n", f"alt b {b_val} ci n"]

    def _flip_polarity(self, cmds):
        result = []
        for cmd in cmds:
            cmd = cmd.replace(" ci p", " TMP1").replace(" ci n", " ci p").replace(" TMP1", " ci n")
            cmd = cmd.replace(" cii p", " TMP2").replace(" cii n", " cii p").replace(" TMP2", " cii n")
            result.append(cmd)
        return result

    # ========== グリッチ制御 ==========
    def _start_glitch_sequence(self, direction, polarity, interval_sec):
        self._glitch_stop(silent=True)
        self._glitch_total = 7; self._glitch_index = 0; self._glitch_remaining = interval_sec
        self._glitch_phase = "sending"; self._glitch_running = True; self._glitch_paused = False
        self._update_glitch_buttons()
        self._glitch_thread = threading.Thread(target=self._glitch_worker, args=(direction, polarity, interval_sec), daemon=True)
        self._glitch_thread.start(); self._append_log("[GLITCH] 開始")

    def _glitch_worker(self, direction, polarity, interval_sec):
        glitch_patterns = self._get_glitch_patterns()
        for idx in range(len(glitch_patterns)):
            if not self._glitch_running: break
            self._glitch_index = idx; self._glitch_remaining = interval_sec; self._glitch_phase = "sending"
            self._refresh_glitch_status(); self._update_glitch_buttons_async()
            pattern = self._apply_direction_polarity(glitch_patterns[idx], direction, polarity)
            try:
                for cmd in pattern:
                    if not self._glitch_running: break
                    self._send_and_log_thread(cmd, sleep_sec=0.03)
                if self._glitch_running and self.datagen.is_connected():
                    self._send_and_log_thread("alt s sab", sleep_sec=0.03)
            except Exception as e:
                self._append_log(f"[GLITCH][ERROR] {e}"); self._glitch_running = False; break
            if not self._glitch_running: break
            self._glitch_phase = "countdown"; self._refresh_glitch_status(); self._update_glitch_buttons_async()
            elapsed, step = 0.0, 0.1
            while self._glitch_running and elapsed < interval_sec:
                while self._glitch_running and self._glitch_paused:
                    self._glitch_remaining = max(0.0, interval_sec - elapsed); self._refresh_glitch_status(); time.sleep(0.1)
                if not self._glitch_running: break
                time.sleep(step); elapsed += step
                self._glitch_remaining = max(0.0, interval_sec - elapsed); self._refresh_glitch_status()
        self._glitch_phase = "idle"; self._glitch_running = False; self._glitch_paused = False
        self.var_glitch_status.set("GLITCH 完了"); self._update_glitch_buttons_async(); self._append_log("[GLITCH] 完了")

    def _get_glitch_patterns(self):
        return [
            ["alt a 20000 ci p", "alt b 1ffff ci p", "alt a 20000 ci n", "alt b 1ffff ci n",
             "alt a 80000 cii p", "alt b 80000 cii p", "alt a 80000 cii n", "alt b 80000 cii n"],
            ["alt a 40000 ci p", "alt b 3ffff ci p", "alt a 40000 ci n", "alt b 3ffff ci n"],
            ["alt a 60000 ci p", "alt b 5ffff ci p", "alt a 60000 ci n", "alt b 5ffff ci n"],
            ["alt a 80000 ci p", "alt b 7ffff ci p", "alt a 80000 ci n", "alt b 7ffff ci n"],
            ["alt a A0000 ci p", "alt b 9ffff ci p", "alt a A0000 ci n", "alt b 9ffff ci n"],
            ["alt a C0000 ci p", "alt b Bffff ci p", "alt a C0000 ci n", "alt b Bffff ci n"],
            ["alt a E0000 ci p", "alt b Dffff ci p", "alt a E0000 ci n", "alt b Dffff ci n"],
        ]

    def _apply_direction_polarity(self, pattern, direction, polarity):
        result = list(pattern)
        if (direction or "").lower() == "down":
            result = [cmd.replace("alt a", "TMPX").replace("alt b", "alt a").replace("TMPX", "alt b") for cmd in result]
        if (polarity or "").lower() == "neg":
            result = self._flip_polarity(result); result = self._invert_ci_codes(result)
        return result

    def _invert_ci_codes(self, cmds):
        out, MASK = [], 0xFFFFF
        pat = re.compile(r'^(alt\s+[ab]\s+)([0-9a-fA-F]{1,5})(\s+ci\s+[pn])$', re.IGNORECASE)
        for cmd in cmds:
            m = pat.match(cmd.strip())
            if not m: out.append(cmd); continue
            pre, hex_str, suf = m.groups()
            val = int(hex_str, 16); inv = MASK - val
            inv_hex = f"{inv:05x}" if inv not in (0, MASK) else ("0" if inv == 0 else "fffff")
            out.append(f"{pre}{inv_hex}{suf}")
        return out

    def _glitch_pause(self):
        if self._glitch_running and not self._glitch_paused and self._glitch_phase == "countdown":
            self._glitch_paused = True; self._refresh_glitch_status(); self._update_glitch_buttons(); self._append_log("[GLITCH] 中断")

    def _glitch_resume(self):
        if self._glitch_running and self._glitch_paused:
            self._glitch_paused = False; self._refresh_glitch_status(); self._update_glitch_buttons(); self._append_log("[GLITCH] 再開")

    def _glitch_stop(self, silent=False):
        if self._glitch_running: self._glitch_running = False
        th = self._glitch_thread; self._glitch_thread = None
        if th and th.is_alive():
            try: th.join(timeout=0.2)
            except Exception: pass
        self._glitch_phase = "idle"; self._glitch_paused = False; self._glitch_remaining = 0.0
        self._update_glitch_buttons()
        if not silent: self.var_glitch_status.set("GLITCH 停止"); self._append_log("[GLITCH] 停止")

    def _refresh_glitch_status(self):
        idx, total = getattr(self, "_glitch_index", 0), getattr(self, "_glitch_total", 7)
        remain, paused = max(0.0, float(getattr(self, "_glitch_remaining", 0.0))), bool(getattr(self, "_glitch_paused", False))
        text = f"GLITCH {idx+1}/{total}  中断中" if paused else f"GLITCH {idx+1}/{total}  次まで{remain:.1f}s"
        try: self.after(0, self.var_glitch_status.set, text)
        except Exception: self.var_glitch_status.set(text)

    def _update_glitch_buttons(self):
        is_glitch = (self.var_mode.get() == "Position" and self.var_amp.get() == "グリッチ")
        running, paused = bool(getattr(self, "_glitch_running", False)), bool(getattr(self, "_glitch_paused", False))
        phase = getattr(self, "_glitch_phase", "idle")
        self.ent_glitch.config(state=("normal" if is_glitch else "disabled"))
        if not is_glitch:
            self.btn_glitch_pause.config(state="disabled"); self.btn_glitch_resume.config(state="disabled")
            self.btn_glitch_stop.config(state="disabled"); self.var_glitch_status.set(""); return
        self.btn_glitch_pause.config(state=("normal" if running and not paused and phase == "countdown" else "disabled"))
        self.btn_glitch_resume.config(state=("normal" if running and paused else "disabled"))
        self.btn_glitch_stop.config(state=("normal" if running else "disabled"))

    def _update_glitch_buttons_async(self):
        try: self.after(0, self._update_glitch_buttons)
        except Exception: self._update_glitch_buttons()

    # ========== カスタムパターン送信 ==========
    def _browse_cstm_file(self, target):
        path = filedialog.askopenfilename(title=f"カスタムパターン({target})ファイルを選択",
                                          filetypes=[("テキストファイル", "*.txt"), ("すべて", "*.*")])
        if not path: return
        if target == "pos": self.var_cstm_pos_path.set(path); update_config_value(["datagen_cstm_pos_path"], path)
        else: self.var_cstm_neg_path.set(path); update_config_value(["datagen_cstm_neg_path"], path)

    def _send_custom_pattern(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です"); return
        pos_path, neg_path = self.var_cstm_pos_path.get().strip(), self.var_cstm_neg_path.get().strip()
        if not pos_path or not neg_path:
            messagebox.showerror("エラー", "pos/negファイルを両方選択してください"); return
        try:
            with open(pos_path, "r", encoding="utf-8") as f: pos_lines = [ln.strip() for ln in f if ln.strip()]
            with open(neg_path, "r", encoding="utf-8") as f: neg_lines = [ln.strip() for ln in f if ln.strip()]
        except Exception as e:
            messagebox.showerror("エラー", f"ファイル読込エラー: {e}"); return
        cstm_mode = self.var_cstm_conn.get()
        ci_lines = pos_lines + neg_lines
        cii_lines = [ln.replace(" ci p", " cii p").replace(" ci n", " cii n") for ln in ci_lines]
        if cstm_mode == "Positionのみ": commands = ci_lines
        elif cstm_mode == "LBCのみ": commands = cii_lines
        else: commands = ci_lines + cii_lines
        total = len(commands); self.var_cstm_progress.set(f"0/{total} (0%)")
        if self.btn_cstm_send: self.btn_cstm_send.config(state="disabled")
        if self.btn_cstm_stop: self.btn_cstm_stop.config(state="normal")
        self._cstm_running = True
        self._cstm_thread = threading.Thread(target=self._cstm_send_worker, args=(commands, total, cstm_mode), daemon=True)
        self._cstm_thread.start()

    def _cstm_send_worker(self, commands, total, cstm_mode):
        try:
            self._append_log(""); self._append_log(f"【カスタムパターン送信開始 ({cstm_mode}, {total}行)】")
            self._reader_enabled = False; time.sleep(0.02)
            self.datagen.write(b"gen stop\r"); time.sleep(0.05); self._append_log("SEND: gen stop")
            if cstm_mode == "Positionのみ":
                self.datagen.write(b"func cstm ci\r"); time.sleep(0.05); self._append_log("SEND: func cstm ci")
                self.datagen.write(b"func alt cii\r"); time.sleep(0.05); self._append_log("SEND: func alt cii")
            elif cstm_mode == "LBCのみ":
                self.datagen.write(b"func alt ci\r"); time.sleep(0.05); self._append_log("SEND: func alt ci")
                self.datagen.write(b"func cstm cii\r"); time.sleep(0.05); self._append_log("SEND: func cstm cii")
            else:
                self.datagen.write(b"func cstm ci\r"); time.sleep(0.05); self._append_log("SEND: func cstm ci")
                self.datagen.write(b"func cstm cii\r"); time.sleep(0.05); self._append_log("SEND: func cstm cii")
            if self.datagen.is_connected(): self.datagen.flush_input()
            for i, cmd in enumerate(commands):
                if not self._cstm_running: self._append_log("【カスタムパターン送信中止】"); break
                self.datagen.write(f"{cmd}\r".encode("utf-8"))
                if (i + 1) % 100 == 0 or i == total - 1:
                    time.sleep(0.005)
                    if self.datagen.is_connected(): self.datagen.flush_input()
                    pct = int((i + 1) / total * 100)
                    try: self.after(0, self.var_cstm_progress.set, f"{i+1}/{total} ({pct}%)")
                    except Exception: pass
            else:
                self._append_log(f"【カスタムパターン送信完了 ({total}行)】")
                if cstm_mode == "Positionのみ":
                    for cmd in ["alt a 80000 cii p", "alt b 80000 cii p", "alt a 80000 cii n", "alt b 80000 cii n"]:
                        self.datagen.write(f"{cmd}\r".encode("utf-8")); time.sleep(0.03)
                    self.datagen.write(b"alt s sa cii\r"); time.sleep(0.03)
                elif cstm_mode == "LBCのみ":
                    for cmd in ["alt a 80000 ci p", "alt b 80000 ci p", "alt a 80000 ci n", "alt b 80000 ci n"]:
                        self.datagen.write(f"{cmd}\r".encode("utf-8")); time.sleep(0.03)
                    self.datagen.write(b"alt s sa ci\r"); time.sleep(0.03)
                if self.var_cstm_auto_start.get():
                    self.datagen.write(b"gen start\r"); time.sleep(0.05); self._append_log("SEND: gen start")
        except Exception as e:
            self._append_log(f"【カスタムパターン送信エラー: {e}】")
        finally:
            self._reader_enabled = True; self._cstm_running = False
            try:
                if self.btn_cstm_send: self.after(0, self.btn_cstm_send.config, {"state": "normal"})
                if self.btn_cstm_stop: self.after(0, self.btn_cstm_stop.config, {"state": "disabled"})
            except Exception: pass

    def _stop_custom_send(self):
        self._cstm_running = False

    def open_cstm_window(self):
        """CSTM/Manual CMDウィンドウを開く"""
        if self._cstm_window and self._cstm_window.winfo_exists():
            self._cstm_window.lift(); self._cstm_window.focus_force(); return
        win = tk.Toplevel(self); win.title("CSTM / Manual CMD")
        root = self.winfo_toplevel(); root.update_idletasks()
        x = root.winfo_x() + root.winfo_width() + 10
        self.pattern_frame.update_idletasks(); y = self.pattern_frame.winfo_rooty()
        win.geometry(f"520x380+{x}+{y}"); win.resizable(False, False); self._cstm_window = win
        pad = ttk.Frame(win, padding=10); pad.pack(fill="both", expand=True)

        func_row = ttk.Frame(pad); func_row.pack(fill="x", pady=(0, 6))
        ttk.Checkbutton(func_row, text="CSTM有効", variable=self.var_func_cstm,
                        command=lambda: self._on_func_change("CSTM")).pack(side="left")
        tk.Label(func_row, text="カスタムパターン設定", fg="#9900CC", font=("Arial", 10, "bold")).pack(side="left", padx=(10, 0))

        target_row = ttk.Frame(pad); target_row.pack(fill="x", pady=(0, 4))
        ttk.Label(target_row, text="対象:").pack(side="left")
        ttk.Combobox(target_row, textvariable=self.var_cstm_conn, values=["Positionのみ", "LBCのみ", "Position/LBC両方"],
                     state="readonly", width=14).pack(side="left", padx=(5, 0))

        pos_row = ttk.Frame(pad); pos_row.pack(fill="x", pady=(0, 4))
        ttk.Label(pos_row, text="posファイル:").pack(side="left")
        ttk.Entry(pos_row, textvariable=self.var_cstm_pos_path, width=50, state="readonly").pack(side="left", padx=(2, 2))
        ttk.Button(pos_row, text="参照", width=4, command=lambda: self._browse_cstm_file("pos")).pack(side="left")

        neg_row = ttk.Frame(pad); neg_row.pack(fill="x", pady=(0, 4))
        ttk.Label(neg_row, text="negファイル:").pack(side="left")
        ttk.Entry(neg_row, textvariable=self.var_cstm_neg_path, width=50, state="readonly").pack(side="left", padx=(2, 2))
        ttk.Button(neg_row, text="参照", width=4, command=lambda: self._browse_cstm_file("neg")).pack(side="left")

        ctrl_row = ttk.Frame(pad); ctrl_row.pack(fill="x", pady=(0, 4))
        ttk.Checkbutton(ctrl_row, text="送信後に自動開始", variable=self.var_cstm_auto_start).pack(side="left")
        self.btn_cstm_send = ttk.Button(ctrl_row, text="送信", command=self._send_custom_pattern)
        self.btn_cstm_send.pack(side="left", padx=(10, 2))
        self.btn_cstm_stop = ttk.Button(ctrl_row, text="中止", command=self._stop_custom_send, state="disabled")
        self.btn_cstm_stop.pack(side="left")

        prog_row = ttk.Frame(pad); prog_row.pack(fill="x")
        ttk.Label(prog_row, text="進捗:").pack(side="left")
        ttk.Label(prog_row, textvariable=self.var_cstm_progress, font=("Arial", 9, "bold")).pack(side="left", padx=(5, 0))

        ttk.Separator(pad, orient="horizontal").pack(fill="x", pady=8)
        tk.Label(pad, text="Manual CMD", fg="#666666", font=("Arial", 10, "bold")).pack(anchor="w")

        cmd_row = ttk.Frame(pad); cmd_row.pack(fill="x", pady=(4, 3))
        ttk.Label(cmd_row, text="CMD:").pack(side="left")
        ent = ttk.Entry(cmd_row, textvariable=self.var_manual_cmd, width=25); ent.pack(side="left", padx=(2, 2))
        ent.bind("<Return>", lambda e: self._send_manual_command())
        ttk.Button(cmd_row, text="送信", command=self._send_manual_command, width=4).pack(side="left")

        rate_row = ttk.Frame(pad); rate_row.pack(fill="x")
        ttk.Label(rate_row, text="Rate:").pack(side="left")
        self._ent_rate_value = ttk.Entry(rate_row, textvariable=self.var_rate_value, width=8)
        self._ent_rate_value.pack(side="left", padx=(2, 2))
        self._cmb_rate_unit = ttk.Combobox(rate_row, textvariable=self.var_rate_unit, values=["msec", "μsec", "nsec"], state="readonly", width=5)
        self._cmb_rate_unit.pack(side="left", padx=(0, 2))
        self._btn_rate_set = ttk.Button(rate_row, text="設定", command=self._set_rate, width=4)
        self._btn_rate_set.pack(side="left")
        self._btn_rate_query = ttk.Button(rate_row, text="問合せ", command=self._query_rate, width=6)
        self._btn_rate_query.pack(side="left", padx=(3, 0))
        ttk.Label(rate_row, text="現在:").pack(side="left", padx=(8, 0))
        ttk.Label(rate_row, textvariable=self.var_rate_display, font=("Arial", 9, "bold")).pack(side="left", padx=(2, 0))

        rnd_row = ttk.Frame(pad); rnd_row.pack(fill="x", pady=(2, 0))
        ttk.Checkbutton(rnd_row, text="Rateランダム変更 (30ns〜10.49ms)", variable=self.var_rate_random,
                        command=self._toggle_rate_random).pack(side="left")
        ttk.Label(rnd_row, text="間隔:").pack(side="left", padx=(8, 0))
        self.var_rate_random_interval = tk.StringVar(value="5")
        ent_rnd = ttk.Entry(rnd_row, textvariable=self.var_rate_random_interval, width=4); ent_rnd.pack(side="left", padx=(2, 0))
        ent_rnd.bind("<FocusOut>", lambda e: self._clamp_rate_random_interval())
        ent_rnd.bind("<Return>", lambda e: self._clamp_rate_random_interval())
        ttk.Label(rnd_row, text="秒").pack(side="left", padx=(2, 0))

        def on_close():
            self.var_rate_random.set(False)
            if self._rate_random_job: self.after_cancel(self._rate_random_job); self._rate_random_job = None
            self.btn_cstm_send = None; self.btn_cstm_stop = None; self._cstm_window = None
            self._ent_rate_value = None; self._cmb_rate_unit = None; self._btn_rate_set = None; self._btn_rate_query = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)
