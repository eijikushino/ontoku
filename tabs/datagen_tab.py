import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText
import time
import re
import threading


class DataGenTab(ttk.Frame):
    """DataGen制御タブ：専用SerialManagerを使用してDataGenを操作"""

    def __init__(self, parent, datagen_manager, datagen_manager2=None):
        super().__init__(parent)
        self.datagen1 = datagen_manager
        self.datagen2 = datagen_manager2
        self.datagen = datagen_manager  # 現在アクティブ
        self.current_dg = 1
        self.initialized = False
        self.initialized2 = False

        # グリッチ制御用
        self._glitch_thread = None
        self._glitch_running = False
        self._glitch_paused = False

        # レスポンスウィンドウ（DataGen1/2それぞれ別）
        self.response_windows = {1: None, 2: None}
        self.response_areas = {1: None, 2: None}

        self._build_ui()

    # ==================== UI構築 ====================
    def _build_ui(self):
        # 色定義
        self.COLOR_DG1_ACTIVE = "#1976D2"
        self.COLOR_DG2_ACTIVE = "#F57C00"
        self.COLOR_DG_INACTIVE = "#E0E0E0"

        # 枠線用コンテナ
        border_container = tk.Frame(self)
        border_container.pack(fill="both", expand=True, padx=5, pady=5)
        border_container.grid_rowconfigure(1, weight=1)
        border_container.grid_columnconfigure(1, weight=1)

        self.border_top = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, height=4)
        self.border_top.grid(row=0, column=0, columnspan=3, sticky="ew")
        self.border_left = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, width=4)
        self.border_left.grid(row=1, column=0, sticky="ns")
        self.border_right = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, width=4)
        self.border_right.grid(row=1, column=2, sticky="ns")
        self.border_bottom = tk.Frame(border_container, bg=self.COLOR_DG1_ACTIVE, height=4)
        self.border_bottom.grid(row=2, column=0, columnspan=3, sticky="ew")

        inner = ttk.Frame(border_container)
        inner.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        inner.grid_columnconfigure(0, weight=1)

        # === DataGen切り替え ===
        dg_switch = ttk.Frame(inner)
        dg_switch.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="w")

        self.btn_dg1 = tk.Button(dg_switch, text="DataGen1", width=10,
                                  font=("Arial", 9, "bold"),
                                  bg=self.COLOR_DG1_ACTIVE, fg="white",
                                  relief="flat", cursor="hand2",
                                  command=lambda: self._switch_datagen(1))
        self.btn_dg1.pack(side="left", padx=(0, 5))

        self.btn_dg2 = tk.Button(dg_switch, text="DataGen2", width=10,
                                  font=("Arial", 9, "bold"),
                                  bg=self.COLOR_DG_INACTIVE, fg="black",
                                  relief="flat", cursor="hand2",
                                  command=lambda: self._switch_datagen(2))
        self.btn_dg2.pack(side="left")
        if not self.datagen2:
            self.btn_dg2.config(state="disabled", bg="#CCCCCC")

        self.var_dg_status = tk.StringVar(value="未接続")
        self.lbl_dg_status = ttk.Label(dg_switch, textvariable=self.var_dg_status,
                                        font=("Arial", 9), foreground="gray")
        self.lbl_dg_status.pack(side="left", padx=(15, 0))

        # === コネクタ設定（表形式）===
        conn_frame = ttk.LabelFrame(inner, text="コネクタ設定", padding=3)
        conn_frame.grid(row=1, column=0, padx=5, pady=(5, 2), sticky="ew")

        conn_table = ttk.Frame(conn_frame)
        conn_table.pack(fill="x")

        ttk.Button(conn_table, text="問合せ", command=self._query_connector_settings, width=10).grid(row=0, column=0, sticky="w")
        ttk.Label(conn_table, text="対象", font=("Arial", 9, "bold"), width=10).grid(row=0, column=1, sticky="w")
        ttk.Label(conn_table, text="FUNC", font=("Arial", 9, "bold"), width=10).grid(row=0, column=2, sticky="w")
        ttk.Label(conn_table, text="INV", font=("Arial", 9, "bold"), width=12).grid(row=0, column=3, sticky="w")

        ttk.Label(conn_table, text="CI(DEF)", font=("Arial", 9), width=10).grid(row=1, column=0, sticky="w")
        self.var_cmode_ci = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_cmode_ci, font=("Arial", 9, "bold"), width=10).grid(row=1, column=1, sticky="w")
        self.var_func_ci = tk.StringVar(value="---")
        self.lbl_func_ci = tk.Label(conn_table, textvariable=self.var_func_ci, font=("Arial", 9, "bold"), width=10, anchor="w")
        self.lbl_func_ci.grid(row=1, column=2, sticky="w")
        self.var_inv_ci = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_inv_ci, font=("Arial", 9, "bold"), width=12).grid(row=1, column=3, sticky="w")

        ttk.Label(conn_table, text="CII(INV/LBC)", font=("Arial", 9), width=10).grid(row=2, column=0, sticky="w")
        self.var_cmode_cii = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_cmode_cii, font=("Arial", 9, "bold"), width=10).grid(row=2, column=1, sticky="w")
        self.var_func_cii = tk.StringVar(value="---")
        self.lbl_func_cii = tk.Label(conn_table, textvariable=self.var_func_cii, font=("Arial", 9, "bold"), width=10, anchor="w")
        self.lbl_func_cii.grid(row=2, column=2, sticky="w")
        self.var_inv_cii = tk.StringVar(value="---")
        ttk.Label(conn_table, textvariable=self.var_inv_cii, font=("Arial", 9, "bold"), width=12).grid(row=2, column=3, sticky="w")

        # 設定行
        conn_set = ttk.Frame(conn_frame)
        conn_set.pack(fill="x", pady=(5, 0))

        ttk.Label(conn_set, text="コネクタ:").pack(side="left")
        self.var_conn_target = tk.StringVar(value="CI")
        ttk.Combobox(conn_set, textvariable=self.var_conn_target,
                     values=["CI", "CII"], state="readonly", width=4).pack(side="left", padx=(2, 10))

        ttk.Label(conn_set, text="対象:").pack(side="left")
        self.var_cmode_set = tk.StringVar(value="397PN")
        ttk.Combobox(conn_set, textvariable=self.var_cmode_set,
                     values=["397", "397PN", "397LBC", "398", "OS"], state="readonly", width=8).pack(side="left", padx=(2, 5))
        ttk.Button(conn_set, text="設定", command=self._send_cmode, width=5).pack(side="left", padx=(0, 15))

        ttk.Label(conn_set, text="INV:").pack(side="left")
        self.var_inv_set = tk.StringVar(value="OFF")
        ttk.Combobox(conn_set, textvariable=self.var_inv_set,
                     values=["ON", "OFF"], state="readonly", width=5).pack(side="left", padx=(2, 5))
        ttk.Button(conn_set, text="設定", command=self._send_inv, width=5).pack(side="left")

        # === FUNC切替用変数（排他制御）===
        self.var_func_alt = tk.BooleanVar(value=True)
        self.var_func_rndm = tk.BooleanVar(value=False)
        self.var_func_rmp = tk.BooleanVar(value=False)

        # === 2値Pattern設定 ===
        pattern_label = ttk.Frame(inner)
        ttk.Checkbutton(pattern_label, text="", variable=self.var_func_alt,
                        command=lambda: self._on_func_change("ALT")).pack(side="left")
        tk.Label(pattern_label, text="2値Pattern設定", fg="#0066CC", font=("Arial", 9, "bold")).pack(side="left")

        pattern_frame = ttk.LabelFrame(inner, labelwidget=pattern_label, padding=3)
        pattern_frame.grid(row=2, column=0, padx=5, pady=(5, 0), sticky="ew")

        # モード
        self.var_mode = tk.StringVar(value="Position")
        mode_frame = ttk.Frame(pattern_frame)
        mode_frame.grid(row=0, column=0, columnspan=7, sticky="w", padx=2, pady=(2, 5))
        ttk.Label(mode_frame, text="モード").pack(side="left")
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
        glitch_btn = ttk.Frame(pattern_frame)
        glitch_btn.grid(row=2, column=4, columnspan=3, sticky="w", padx=(10, 0))
        self.btn_glitch_pause = ttk.Button(glitch_btn, text="中断", width=6, state="disabled", command=self._glitch_pause)
        self.btn_glitch_resume = ttk.Button(glitch_btn, text="再開", width=6, state="disabled", command=self._glitch_resume)
        self.btn_glitch_stop = ttk.Button(glitch_btn, text="終了", width=6, state="disabled", command=self._glitch_stop)
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

        # === ランダム設定 ===
        rndm_label = ttk.Frame(inner)
        ttk.Checkbutton(rndm_label, text="", variable=self.var_func_rndm,
                        command=lambda: self._on_func_change("RNDM")).pack(side="left")
        tk.Label(rndm_label, text="ランダム設定", fg="#CC0000", font=("Arial", 9, "bold")).pack(side="left")

        rndm_frame = ttk.LabelFrame(inner, labelwidget=rndm_label, padding=3)
        rndm_frame.grid(row=3, column=0, padx=5, pady=0, sticky="ew")

        rndm_set = ttk.Frame(rndm_frame)
        rndm_set.pack(fill="x", pady=(0, 2))

        ttk.Label(rndm_set, text="ビット数:").pack(side="left")
        self.var_rndm_bit = tk.StringVar(value="20")
        ttk.Combobox(rndm_set, textvariable=self.var_rndm_bit,
                     values=["20", "16", "14", "14H", "14L"], state="readonly", width=5).pack(side="left", padx=(5, 10))

        ttk.Label(rndm_set, text="コネクタ:").pack(side="left")
        self.var_rndm_conn = tk.StringVar(value="CI")
        ttk.Combobox(rndm_set, textvariable=self.var_rndm_conn,
                     values=["CI", "CII"], state="readonly", width=4).pack(side="left", padx=(5, 10))
        ttk.Button(rndm_set, text="設定", command=self._send_rndm).pack(side="left", padx=(10, 0))

        rndm_disp = ttk.Frame(rndm_frame)
        rndm_disp.pack(fill="x")
        ttk.Button(rndm_disp, text="問合せ", command=self._query_rndm).pack(side="left")
        ttk.Label(rndm_disp, text="CI(DEF):").pack(side="left", padx=(10, 0))
        self.var_rndm_ci = tk.StringVar(value="---")
        ttk.Label(rndm_disp, textvariable=self.var_rndm_ci, font=("Arial", 9, "bold"), width=14).pack(side="left", padx=(2, 5))
        ttk.Label(rndm_disp, text="CII(INV/LBC):").pack(side="left")
        self.var_rndm_cii = tk.StringVar(value="---")
        ttk.Label(rndm_disp, textvariable=self.var_rndm_cii, font=("Arial", 9, "bold"), width=14).pack(side="left", padx=(2, 0))

        # === ランプ設定 ===
        ramp_label = ttk.Frame(inner)
        ttk.Checkbutton(ramp_label, text="", variable=self.var_func_rmp,
                        command=lambda: self._on_func_change("RMP")).pack(side="left")
        tk.Label(ramp_label, text="ランプ設定", fg="#009900", font=("Arial", 9, "bold")).pack(side="left")

        ramp_frame = ttk.LabelFrame(inner, labelwidget=ramp_label, padding=3)
        ramp_frame.grid(row=4, column=0, padx=5, pady=0, sticky="ew")

        ramp_set = ttk.Frame(ramp_frame)
        ramp_set.pack(fill="x", pady=(0, 5))
        ttk.Label(ramp_set, text="コネクタ:").pack(side="left")
        self.var_rmp_conn = tk.StringVar(value="CI")
        ttk.Combobox(ramp_set, textvariable=self.var_rmp_conn,
                     values=["CI", "CII"], state="readonly", width=4).pack(side="left", padx=(2, 10))
        ttk.Label(ramp_set, text="BGN:").pack(side="left")
        self.var_rmp_set_bgn = tk.StringVar(value="00000")
        ttk.Entry(ramp_set, textvariable=self.var_rmp_set_bgn, width=7).pack(side="left", padx=(2, 10))
        ttk.Label(ramp_set, text="END:").pack(side="left")
        self.var_rmp_set_end = tk.StringVar(value="FFFFF")
        ttk.Entry(ramp_set, textvariable=self.var_rmp_set_end, width=7).pack(side="left", padx=(2, 10))
        ttk.Label(ramp_set, text="STEP:").pack(side="left")
        self.var_rmp_set_step = tk.StringVar(value="00001")
        ttk.Entry(ramp_set, textvariable=self.var_rmp_set_step, width=7).pack(side="left", padx=(2, 10))
        ttk.Button(ramp_set, text="設定", command=self._send_rmp).pack(side="left")

        # ランプ表示（表形式）
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

        # === 手動コマンド・レート設定 ===
        manual_frame = ttk.LabelFrame(inner, text="手動コマンド・レート設定", padding=3)
        manual_frame.grid(row=5, column=0, padx=5, pady=(2, 5), sticky="ew")

        cmd_row = ttk.Frame(manual_frame)
        cmd_row.pack(fill="x", pady=(0, 3))
        ttk.Label(cmd_row, text="CMD:").pack(side="left")
        self.var_manual_cmd = tk.StringVar()
        ent_cmd = ttk.Entry(cmd_row, textvariable=self.var_manual_cmd, width=18)
        ent_cmd.pack(side="left", padx=(2, 2))
        ent_cmd.bind("<Return>", lambda e: self._send_manual_command())
        ttk.Button(cmd_row, text="送信", command=self._send_manual_command, width=4).pack(side="left")
        ttk.Button(cmd_row, text="レスポンス窓", command=self._show_response_window).pack(side="left", padx=(10, 0))
        self.var_show_recv = tk.BooleanVar(value=False)
        ttk.Checkbutton(cmd_row, text="RECV表示", variable=self.var_show_recv).pack(side="left", padx=(5, 0))

        rate_row = ttk.Frame(manual_frame)
        rate_row.pack(fill="x")
        ttk.Label(rate_row, text="Rate:").pack(side="left")
        self.var_rate_value = tk.StringVar(value="10")
        ttk.Entry(rate_row, textvariable=self.var_rate_value, width=8).pack(side="left", padx=(2, 2))
        self.var_rate_unit = tk.StringVar(value="msec")
        ttk.Combobox(rate_row, textvariable=self.var_rate_unit,
                     values=["msec", "\u03bcsec", "nsec"], state="readonly", width=5).pack(side="left", padx=(0, 2))
        ttk.Button(rate_row, text="設定", command=self._set_rate, width=4).pack(side="left")
        ttk.Button(rate_row, text="問合せ", command=self._query_rate, width=6).pack(side="left", padx=(3, 0))
        ttk.Label(rate_row, text="現在:").pack(side="left", padx=(8, 0))
        self.var_rate_display = tk.StringVar(value="---")
        ttk.Label(rate_row, textvariable=self.var_rate_display, font=("Arial", 9, "bold")).pack(side="left", padx=(2, 0))

        # 起動時のデフォルト状態を適用
        self._on_amp_change()

    # ==================== イベントハンドラ ====================
    def _on_func_change(self, func_type):
        if func_type == "ALT":
            self.var_func_alt.set(True)
            self.var_func_rndm.set(False)
            self.var_func_rmp.set(False)
        elif func_type == "RNDM":
            self.var_func_alt.set(False)
            self.var_func_rndm.set(True)
            self.var_func_rmp.set(False)
        elif func_type == "RMP":
            self.var_func_alt.set(False)
            self.var_func_rndm.set(False)
            self.var_func_rmp.set(True)

        if not self.datagen or not self.datagen.is_connected():
            return

        self._append_log("")
        self._append_log(f"\u3010FUNC\u5207\u66ff: {func_type}\u3011")
        self._send_and_log("gen stop")
        self._send_and_log(f"func {func_type.lower()}")
        self._send_and_log("gen start")
        self.after(100, self._query_connector_settings)

        if func_type == "ALT":
            self.var_amp.set("\u9759\u6b62")
            self._on_amp_change()
            self.after(200, self._send_pattern)
        elif func_type == "RNDM":
            self.after(200, self._query_rndm)
        elif func_type == "RMP":
            self.after(200, self._query_rmp)

    def _on_mode_change(self):
        mode = self.var_mode.get()
        if mode == "LBC":
            self.amp_combo['values'] = ["FS", "LBC\u30b0\u30ea\u30c3\u30c1", "\u9759\u6b62"]
            if self.var_amp.get() not in ["FS", "LBC\u30b0\u30ea\u30c3\u30c1", "\u9759\u6b62"]:
                self.var_amp.set("FS")
            self.center_combo.set("")
            self.center_combo.config(state="disabled")
        else:
            self.amp_combo['values'] = ["1/32FS", "FS", "MajorCarry", "\u9759\u6b62", "\u30b0\u30ea\u30c3\u30c1"]
            if self.var_amp.get() not in ["1/32FS", "FS", "MajorCarry", "\u9759\u6b62", "\u30b0\u30ea\u30c3\u30c1"]:
                self.var_amp.set("1/32FS")
            self.center_combo.config(state="readonly")
            if not self.var_center.get():
                self.var_center.set("0V")
        self._on_amp_change()

    def _on_amp_change(self, event=None):
        amp = self.var_amp.get()
        mode = self.var_mode.get()

        self.ent_glitch.config(state="disabled")
        self.btn_glitch_pause.config(state="disabled")
        self.btn_glitch_resume.config(state="disabled")
        self.btn_glitch_stop.config(state="disabled")
        self.var_glitch_status.set("")

        if mode == "Position":
            self.center_combo.config(state="readonly")
        else:
            self.center_combo.set("")
            self.center_combo.config(state="disabled")

        if amp == "FS":
            if mode == "Position":
                self.center_combo['values'] = ["\u00b1160V"]
                self.center_combo.set("\u00b1160V")
            self.dir_combo.config(state="readonly")
            self.pol_combo.config(state="readonly")
        elif amp == "1/32FS":
            if mode == "Position":
                self.center_combo['values'] = ["0V", "+160V", "-160V"]
                if self.var_center.get() not in ["0V", "+160V", "-160V"]:
                    self.var_center.set("0V")
            self.dir_combo.config(state="readonly")
            self.pol_combo.config(state="readonly")
        elif amp == "MajorCarry":
            self.center_combo.set("")
            self.center_combo.config(state="disabled")
            self.dir_combo.config(state="readonly")
            self.pol_combo.config(state="readonly")
        elif amp == "\u9759\u6b62":
            self.center_combo.set("")
            self.center_combo.config(state="disabled")
            self.var_dir.set("")
            self.dir_combo.set("")
            self.dir_combo.config(state="disabled")
            self.var_pol.set("")
            self.pol_combo.set("")
            self.pol_combo.config(state="disabled")
        elif amp == "\u30b0\u30ea\u30c3\u30c1":
            self.center_combo.set("")
            self.center_combo.config(state="disabled")
            self.dir_combo.config(state="readonly")
            self.pol_combo.config(state="readonly")
            self.ent_glitch.config(state="normal")
        elif amp == "LBC\u30b0\u30ea\u30c3\u30c1":
            self.center_combo.set("")
            self.center_combo.config(state="disabled")
            self.dir_combo.set("")
            self.dir_combo.config(state="disabled")
            self.pol_combo.set("")
            self.pol_combo.config(state="disabled")

    # ==================== DataGen切り替え ====================
    def _switch_datagen(self, dg_num):
        if dg_num == self.current_dg:
            return
        if dg_num == 2 and not self.datagen2:
            return

        self.current_dg = dg_num
        if dg_num == 1:
            self.datagen = self.datagen1
            self.btn_dg1.config(bg=self.COLOR_DG1_ACTIVE, fg="white")
            self.btn_dg2.config(bg=self.COLOR_DG_INACTIVE, fg="black")
            border_color = self.COLOR_DG1_ACTIVE
        else:
            self.datagen = self.datagen2
            self.btn_dg1.config(bg=self.COLOR_DG_INACTIVE, fg="black")
            self.btn_dg2.config(bg=self.COLOR_DG2_ACTIVE, fg="white")
            border_color = self.COLOR_DG2_ACTIVE

        self.border_top.config(bg=border_color)
        self.border_left.config(bg=border_color)
        self.border_right.config(bg=border_color)
        self.border_bottom.config(bg=border_color)

        self._update_dg_status()
        self._append_log("")
        self._append_log(f"\u3010DataGen{dg_num} \u306b\u5207\u308a\u66ff\u3048\u3011")

        if self.datagen and self.datagen.is_connected():
            self.after(100, self._query_connector_settings)

    def _update_dg_status(self):
        if self.datagen and self.datagen.is_connected():
            port = getattr(self.datagen.ser, 'port', '') if self.datagen.ser else ''
            self.var_dg_status.set(f"{port} \u63a5\u7d9a\u4e2d")
            self.lbl_dg_status.config(foreground="green")
        else:
            self.var_dg_status.set("\u672a\u63a5\u7d9a")
            self.lbl_dg_status.config(foreground="gray")

    # ==================== レスポンスウィンドウ ====================
    def _show_response_window(self):
        dg = self.current_dg
        if self.response_windows[dg] and self.response_windows[dg].winfo_exists():
            self.response_windows[dg].lift()
            self.response_windows[dg].focus_force()
            return

        border_color = self.COLOR_DG1_ACTIVE if dg == 1 else self.COLOR_DG2_ACTIVE
        win = tk.Toplevel(self)
        win.title(f"DataGen{dg} \u30ec\u30b9\u30dd\u30f3\u30b9\u8868\u793a")

        main_win = self.winfo_toplevel()
        main_win.update_idletasks()
        main_x = main_win.winfo_x()
        main_y = main_win.winfo_y()
        main_width = main_win.winfo_width()
        main_height = main_win.winfo_height()
        x_offset = main_x + main_width + 10 + (410 if dg == 2 else 0)
        win.geometry(f"400x{main_height}+{x_offset}+{main_y}")
        win.resizable(True, True)

        bc = tk.Frame(win)
        bc.pack(fill="both", expand=True)
        bc.grid_rowconfigure(1, weight=1)
        bc.grid_columnconfigure(1, weight=1)
        tk.Frame(bc, bg=border_color, height=4).grid(row=0, column=0, columnspan=3, sticky="ew")
        tk.Frame(bc, bg=border_color, width=4).grid(row=1, column=0, sticky="ns")
        tk.Frame(bc, bg=border_color, width=4).grid(row=1, column=2, sticky="ns")
        tk.Frame(bc, bg=border_color, height=4).grid(row=2, column=0, columnspan=3, sticky="ew")

        inner = ttk.Frame(bc)
        inner.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
        response_area = ScrolledText(inner, font=("Consolas", 10), wrap="none")
        response_area.pack(fill="both", expand=True)

        btn_f = ttk.Frame(inner)
        btn_f.pack(fill="x", pady=(5, 0))
        ttk.Button(btn_f, text="\u30af\u30ea\u30a2",
                   command=lambda ra=response_area: ra.delete(1.0, tk.END)).pack(side="right")

        self.response_windows[dg] = win
        self.response_areas[dg] = response_area

        def on_close(d=dg):
            self.response_windows[d] = None
            self.response_areas[d] = None
            win.destroy()
        win.protocol("WM_DELETE_WINDOW", on_close)

    # ==================== ログ・通信 ====================
    def _append_log(self, text):
        response_area = self.response_areas.get(self.current_dg)
        if response_area:
            response_area.insert(tk.END, text.rstrip() + "\n")
            response_area.see(tk.END)

    def _send_and_log(self, cmd, sleep_sec=0.05, add_blank=False):
        if add_blank:
            self._append_log("")
        self._append_log(f"SEND: {cmd}")
        if self.var_show_recv.get():
            response = self.datagen.send_command_with_response(cmd)
            if response:
                self._append_log(f"RECV:\n{response}")
        else:
            self.datagen.send_command(cmd)
            time.sleep(sleep_sec)

    def _send_and_log_thread(self, cmd, sleep_sec=0.05):
        show_recv = self.var_show_recv.get()
        if show_recv:
            response = self.datagen.send_command_with_response(cmd)
        else:
            self.datagen.send_command(cmd)
            time.sleep(sleep_sec)
            response = None
        try:
            self.after(0, lambda: self._append_log(f"SEND: {cmd}"))
            if show_recv and response:
                self.after(0, lambda r=response: self._append_log(f"RECV:\n{r}"))
        except Exception:
            pass

    # ==================== コマンド送信 ====================
    def _send_manual_command(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        cmd = self.var_manual_cmd.get().strip()
        if not cmd:
            return
        self._append_log("")
        self._append_log(f"SEND: {cmd}")
        response = self.datagen.send_command_with_response(cmd, wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
            if cmd.lower().startswith("rate"):
                self._update_rate_display(response)
        self.var_manual_cmd.set("")

    def _set_rate(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        try:
            value = float(self.var_rate_value.get())
        except ValueError:
            messagebox.showerror("\u30a8\u30e9\u30fc", "\u6570\u5024\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044")
            return

        unit = self.var_rate_unit.get()
        if unit == "msec":
            ns_val = value * 1000000
        elif unit == "\u03bcsec":
            ns_val = value * 1000
        else:
            ns_val = value

        rate_val = int(ns_val / 10)
        rate_val = max(3, min(rate_val, 1048575))

        cmd = f"rate {rate_val}"
        self._append_log("")
        self._append_log(f"SEND: {cmd}")
        response = self.datagen.send_command_with_response(cmd, wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_rate)

    def _send_rndm(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        cmd = f"rndm {self.var_rndm_bit.get()} {self.var_rndm_conn.get()}"
        self._append_log("")
        self._append_log(f"SEND: {cmd}")
        response = self.datagen.send_command_with_response(cmd, wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_rndm)

    def _send_cmode(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        cmd = f"cmode {self.var_cmode_set.get().lower()} {self.var_conn_target.get().lower()}"
        self._append_log("")
        self._append_log(f"SEND: {cmd}")
        response = self.datagen.send_command_with_response(cmd, wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_connector_settings)

    def _send_inv(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        cmd = f"inv {self.var_inv_set.get().lower()} {self.var_conn_target.get().lower()}"
        self._append_log("")
        self._append_log(f"SEND: {cmd}")
        response = self.datagen.send_command_with_response(cmd, wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_connector_settings)

    def _send_rmp(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        conn = self.var_rmp_conn.get().lower()
        self._append_log("")
        for param, var in [("bgn", self.var_rmp_set_bgn), ("end", self.var_rmp_set_end), ("step", self.var_rmp_set_step)]:
            cmd = f"rmp {param} {var.get().strip()} {conn}"
            self._append_log(f"SEND: {cmd}")
            response = self.datagen.send_command_with_response(cmd, wait_sec=0.05, read_timeout=0.05)
            if response:
                self._append_log(f"RECV:\n{response}")
        self.after(100, self._query_rmp)

    # ==================== 問い合わせ ====================
    def _query_connector_settings(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        self._append_log("")
        for qcmd, updater in [("cmode", self._update_cmode_display),
                               ("inv", self._update_inv_display),
                               ("func", self._update_func_display)]:
            self._append_log(f"SEND: {qcmd}")
            response = self.datagen.send_command_with_response(qcmd, wait_sec=0.05, read_timeout=0.05)
            if response:
                self._append_log(f"RECV:\n{response}")
                updater(response)

    def _query_rndm(self):
        if not self.datagen or not self.datagen.is_connected():
            return
        self._append_log("")
        self._append_log("SEND: rndm")
        response = self.datagen.send_command_with_response("rndm", wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
            self._update_rndm_display(response)

    def _query_rmp(self):
        if not self.datagen or not self.datagen.is_connected():
            return
        self._append_log("")
        self._append_log("SEND: rmp")
        response = self.datagen.send_command_with_response("rmp", wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
            self._update_rmp_display(response)

    def _query_rate(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        self._append_log("")
        self._append_log("SEND: rate")
        response = self.datagen.send_command_with_response("rate", wait_sec=0.05, read_timeout=0.05)
        if response:
            self._append_log(f"RECV:\n{response}")
            self._update_rate_display(response)

    # ==================== 表示更新 ====================
    def _update_cmode_display(self, response):
        ci, cii = "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector I "):
                m = re.search(r':\s*(.+)', line)
                if m:
                    ci = m.group(1).strip()
            elif line.startswith("connector II "):
                m = re.search(r':\s*(.+)', line)
                if m:
                    cii = m.group(1).strip()
        self.var_cmode_ci.set(ci)
        self.var_cmode_cii.set(cii)

    def _update_inv_display(self, response):
        ci, cii = "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector I "):
                m = re.search(r':\s*(.+)', line)
                if m:
                    ci = m.group(1).strip()
            elif line.startswith("connector II "):
                m = re.search(r':\s*(.+)', line)
                if m:
                    cii = m.group(1).strip()
        self.var_inv_ci.set(ci)
        self.var_inv_cii.set(cii)

    def _convert_func_name(self, func_type):
        if "square" in func_type.lower():
            return "2\u5024\u632f\u308a"
        return func_type

    def _get_func_color(self, func_type):
        ft = func_type.lower()
        if "square" in ft:
            return "#0066CC"
        elif "ramp" in ft:
            return "#009900"
        elif "random" in ft:
            return "#CC0000"
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
        if ci_p != "---" and ci_n != "---" and ci_p != ci_n:
            warnings.append(f"CI: pos={ci_p}, neg={ci_n}")
        if cii_p != "---" and cii_n != "---" and cii_p != cii_n:
            warnings.append(f"CII: pos={cii_p}, neg={cii_n}")
        if warnings:
            messagebox.showwarning("FUNC\u8a2d\u5b9a\u8b66\u544a", "P/N\u8a2d\u5b9a\u304c\u7570\u306a\u3063\u3066\u3044\u307e\u3059:\n" + "\n".join(warnings))

        self.var_func_ci.set(self._convert_func_name(ci_p))
        self.var_func_cii.set(self._convert_func_name(cii_p))
        self.lbl_func_ci.config(fg=self._get_func_color(ci_p))
        self.lbl_func_cii.config(fg=self._get_func_color(cii_p))

    def _update_rndm_display(self, response):
        ci, cii = "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector  I "):
                m = re.search(r':\s*(.+)', line)
                if m:
                    ci = m.group(1).replace(" ", "")
                    break
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector II "):
                m = re.search(r':\s*(.+)', line)
                if m:
                    cii = m.group(1).replace(" ", "")
                    break
        self.var_rndm_ci.set(ci)
        self.var_rndm_cii.set(cii)

    def _update_rmp_display(self, response):
        ci_bgn, ci_end, ci_step = "---", "---", "---"
        cii_bgn, cii_end, cii_step = "---", "---", "---"
        for line in response.split('\n'):
            line = line.strip()
            if line.startswith("connector  I pos"):
                m = re.search(r':\s*(\w+)\s+(\w+)\s+(\w+)', line)
                if m:
                    ci_bgn, ci_end, ci_step = m.group(1), m.group(2), m.group(3)
            elif line.startswith("connector II pos"):
                m = re.search(r':\s*(\w+)\s+(\w+)\s+(\w+)', line)
                if m:
                    cii_bgn, cii_end, cii_step = m.group(1), m.group(2), m.group(3)
        self.var_rmp_ci_bgn.set(ci_bgn)
        self.var_rmp_ci_end.set(ci_end)
        self.var_rmp_ci_step.set(ci_step)
        self.var_rmp_cii_bgn.set(cii_bgn)
        self.var_rmp_cii_end.set(cii_end)
        self.var_rmp_cii_step.set(cii_step)

    def _update_rate_display(self, response):
        m = re.search(r'data rate\s*:\s*(\d+)\s*\[ns\]\s*\((\d+)\s*\[Hz\]\)', response)
        if m:
            ns_val = int(m.group(1))
            hz_val = int(m.group(2))
            if ns_val >= 1000000:
                t_str = f"{ns_val / 1000000:.2f}msec"
            elif ns_val >= 1000:
                t_str = f"{ns_val / 1000:.1f}\u03bcsec"
            else:
                t_str = f"{ns_val}nsec"
            if hz_val >= 1000000:
                f_str = f"{hz_val / 1000000:.2f}MHz"
            elif hz_val >= 1000:
                f_str = f"{hz_val / 1000:.1f}kHz"
            else:
                f_str = f"{hz_val}Hz"
            self.var_rate_display.set(f"{t_str} ({f_str})")
            return
        if "period <=" in response:
            self.after(100, self._query_rate)

    # ==================== Init/Pattern ====================
    def _send_init(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        init_commands = [
            "gen stop",
            "alt s sab ci", "alt s sab cii",
            "func alt ci", "func alt cii",
            "inv off ci", "inv off cii",
            "trg mon",
            "rate 1000000",
            "cmode 397pn ci", "cmode 397lbc cii",
            "stb 397pn ts 5", "stb 397pn tw 4", "stb 397pn th 5", "stb 397pn tpn 10",
            "stb 397lbc ts 8", "stb 397lbc tw 8", "stb 397lbc th 90", "stb 397lbc tpn 0",
            "stb 397 ts 10", "stb 397 tw 8", "stb 397 th 10", "stb 397 tpn 0",
            "gen start",
        ]
        self._append_log("")
        self._append_log(f"\u3010Init\u9001\u4fe1\u958b\u59cb (DataGen{self.current_dg})\u3011")
        for cmd in init_commands:
            self._send_and_log(cmd, sleep_sec=0.05)
        if self.current_dg == 1:
            self.initialized = True
        else:
            self.initialized2 = True
        self._append_log("\u3010Init\u9001\u4fe1\u5b8c\u4e86\u3011")
        self.after(100, self._query_rate)

    def _send_hold_a(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        self._send_and_log("alt s sa", add_blank=True)

    def _start_alternating(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return
        self._send_and_log("alt s sab", add_blank=True)

    def _send_pattern(self):
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("\u30a8\u30e9\u30fc", "DataGen\u304c\u672a\u63a5\u7d9a\u3067\u3059")
            return

        is_initialized = self.initialized if self.current_dg == 1 else self.initialized2
        if not is_initialized:
            self._send_init()

        amp = self.var_amp.get()
        center = self.var_center.get()
        direction = self.var_dir.get()
        polarity = self.var_pol.get()
        mode = self.var_mode.get()

        if mode == "Position" and amp == "\u30b0\u30ea\u30c3\u30c1":
            try:
                sec = float(self.var_glitch_sec.get())
                if sec <= 0:
                    raise ValueError
            except Exception:
                messagebox.showerror("\u30a8\u30e9\u30fc", "\u30b0\u30ea\u30c3\u30c1\u9077\u79fb\u6642\u9593(sec)\u306b\u306f\u6b63\u306e\u6570\u5024\u3092\u5165\u529b\u3057\u3066\u304f\u3060\u3055\u3044\u3002")
                return
            self._start_glitch_sequence(direction, polarity, sec)
            return

        pattern_ci, pattern_cii = self._get_patterns(mode, amp, center, direction, polarity)
        self._append_log("")
        if pattern_ci:
            for cmd in pattern_ci:
                self._send_and_log(cmd, sleep_sec=0.05)
        if pattern_cii:
            for cmd in pattern_cii:
                self._send_and_log(cmd, sleep_sec=0.05)
        self._send_and_log("alt s sab")

    # ==================== パターン生成 ====================
    def _get_patterns(self, mode, amp, center, direction, polarity):
        cii_static = [
            "alt a 80000 cii p", "alt b 80000 cii p",
            "alt a 80000 cii n", "alt b 80000 cii n",
        ]
        ci_static = [
            "alt a 80000 ci p", "alt b 80000 ci p",
            "alt a 80000 ci n", "alt b 80000 ci n",
        ]

        if amp == "\u9759\u6b62":
            return ci_static, cii_static

        if mode == "LBC":
            if amp == "LBC\u30b0\u30ea\u30c3\u30c1":
                cii = [
                    "alt a 80030 cii p", "alt b 7ffff cii p",
                    "alt a 80030 cii n", "alt b 7ffff cii n",
                ]
                return ci_static, cii
            elif amp == "FS":
                if direction == "up":
                    cii = [
                        "alt a fffff cii p", "alt b 00000 cii p",
                        "alt a 00000 cii n", "alt b fffff cii n"
                    ]
                else:
                    cii = [
                        "alt a 00000 cii p", "alt b fffff cii p",
                        "alt a fffff cii n", "alt b 00000 cii n"
                    ]
                if polarity == "neg":
                    cii = self._flip_polarity(cii)
                return ci_static, cii

        if amp == "FS":
            return self._build_position_fs(direction, polarity), cii_static
        elif amp == "1/32FS":
            return self._build_position_132(center, direction, polarity), cii_static
        elif amp == "MajorCarry":
            return self._build_position_majorcarry(direction, polarity), cii_static

        return ci_static, cii_static

    def _build_position_fs(self, direction, polarity):
        d = (direction or "").lower()
        p = (polarity or "").lower()
        a_val, b_val = ("fffff", "00000") if d == "up" else ("00000", "fffff")
        if p == "neg":
            a_val, b_val = b_val, a_val
        return [f"alt a {a_val} ci p", f"alt b {b_val} ci p",
                f"alt a {a_val} ci n", f"alt b {b_val} ci n"]

    def _build_position_132(self, center, direction, polarity):
        c = (center or "").upper()
        d = (direction or "").lower()
        p = (polarity or "").lower()

        if p == "neg":
            if c in ("+160V", "+FULL"):
                c_eff = "-160V"
            elif c in ("-160V", "-FULL"):
                c_eff = "+160V"
            else:
                c_eff = "0V"
        else:
            c_eff = c

        if c_eff in ("+160V", "+FULL"):
            left, right = "f7fff", "fffff"
        elif c_eff in ("-160V", "-FULL"):
            left, right = "00000", "08000"
        else:
            left, right = "7bfff", "84000"

        a_val, b_val = left, right
        if d == "up":
            a_val, b_val = b_val, a_val
        if p == "neg":
            a_val, b_val = b_val, a_val

        return [f"alt a {a_val} ci p", f"alt b {b_val} ci p",
                f"alt a {a_val} ci n", f"alt b {b_val} ci n"]

    def _build_position_majorcarry(self, direction, polarity):
        d = (direction or "").lower()
        p = (polarity or "").lower()
        a_val, b_val = "7ffff", "80000"
        if d == "up":
            a_val, b_val = b_val, a_val
        if p == "neg":
            a_val, b_val = b_val, a_val
        return [f"alt a {a_val} ci p", f"alt b {b_val} ci p",
                f"alt a {a_val} ci n", f"alt b {b_val} ci n"]

    def _flip_polarity(self, cmds):
        result = []
        for cmd in cmds:
            cmd = cmd.replace(" ci p", " TMP1").replace(" ci n", " ci p").replace(" TMP1", " ci n")
            cmd = cmd.replace(" cii p", " TMP2").replace(" cii n", " cii p").replace(" TMP2", " cii n")
            result.append(cmd)
        return result

    # ==================== グリッチ制御 ====================
    def _start_glitch_sequence(self, direction, polarity, interval_sec):
        self._glitch_stop(silent=True)
        self._glitch_total = 7
        self._glitch_index = 0
        self._glitch_remaining = interval_sec
        self._glitch_phase = "sending"
        self._glitch_running = True
        self._glitch_paused = False
        self._update_glitch_buttons()
        self._glitch_thread = threading.Thread(
            target=self._glitch_worker, args=(direction, polarity, interval_sec), daemon=True
        )
        self._glitch_thread.start()
        self._append_log("[GLITCH] \u958b\u59cb")

    def _glitch_worker(self, direction, polarity, interval_sec):
        glitch_patterns = self._get_glitch_patterns()
        total = len(glitch_patterns)

        for idx in range(total):
            if not self._glitch_running:
                break

            self._glitch_index = idx
            self._glitch_remaining = interval_sec
            self._glitch_phase = "sending"
            self._refresh_glitch_status()
            self._update_glitch_buttons_async()

            pattern = self._apply_direction_polarity(glitch_patterns[idx], direction, polarity)
            try:
                for cmd in pattern:
                    if not self._glitch_running:
                        break
                    self._send_and_log_thread(cmd, sleep_sec=0.03)
                if self._glitch_running and self.datagen.is_connected():
                    self._send_and_log_thread("alt s sab", sleep_sec=0.03)
            except Exception as e:
                self._append_log(f"[GLITCH][ERROR] {e}")
                self._glitch_running = False
                break

            if not self._glitch_running:
                break

            self._glitch_phase = "countdown"
            self._refresh_glitch_status()
            self._update_glitch_buttons_async()

            elapsed = 0.0
            step = 0.1
            while self._glitch_running and elapsed < interval_sec:
                while self._glitch_running and self._glitch_paused:
                    self._glitch_remaining = max(0.0, interval_sec - elapsed)
                    self._refresh_glitch_status()
                    time.sleep(0.1)
                if not self._glitch_running:
                    break
                time.sleep(step)
                elapsed += step
                self._glitch_remaining = max(0.0, interval_sec - elapsed)
                self._refresh_glitch_status()

        self._glitch_phase = "idle"
        self._glitch_running = False
        self._glitch_paused = False
        self.var_glitch_status.set("GLITCH \u5b8c\u4e86")
        self._update_glitch_buttons_async()
        self._append_log("[GLITCH] \u5b8c\u4e86")

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
            swapped = []
            for cmd in result:
                tmp = cmd.replace("alt a", "TMPX").replace("alt b", "alt a").replace("TMPX", "alt b")
                swapped.append(tmp)
            result = swapped
        if (polarity or "").lower() == "neg":
            result = self._flip_polarity(result)
            result = self._invert_ci_codes(result)
        return result

    def _invert_ci_codes(self, cmds):
        out = []
        MASK = 0xFFFFF
        pat = re.compile(r'^(alt\s+[ab]\s+)([0-9a-fA-F]{1,5})(\s+ci\s+[pn])$', re.IGNORECASE)
        for cmd in cmds:
            m = pat.match(cmd.strip())
            if not m:
                out.append(cmd)
                continue
            pre, hex_str, suf = m.groups()
            val = int(hex_str, 16)
            inv = MASK - val
            inv_hex = f"{inv:05x}" if inv not in (0, MASK) else ("0" if inv == 0 else "fffff")
            out.append(f"{pre}{inv_hex}{suf}")
        return out

    def _glitch_pause(self):
        if self._glitch_running and not self._glitch_paused and self._glitch_phase == "countdown":
            self._glitch_paused = True
            self._refresh_glitch_status()
            self._update_glitch_buttons()
            self._append_log("[GLITCH] \u4e2d\u65ad")

    def _glitch_resume(self):
        if self._glitch_running and self._glitch_paused:
            self._glitch_paused = False
            self._refresh_glitch_status()
            self._update_glitch_buttons()
            self._append_log("[GLITCH] \u518d\u958b")

    def _glitch_stop(self, silent=False):
        if self._glitch_running:
            self._glitch_running = False
        th = self._glitch_thread
        self._glitch_thread = None
        if th and th.is_alive():
            try:
                th.join(timeout=0.2)
            except Exception:
                pass
        self._glitch_phase = "idle"
        self._glitch_paused = False
        self._glitch_remaining = 0.0
        self._update_glitch_buttons()
        if not silent:
            self.var_glitch_status.set("GLITCH \u505c\u6b62")
            self._append_log("[GLITCH] \u505c\u6b62")

    def _refresh_glitch_status(self):
        idx = getattr(self, "_glitch_index", 0)
        total = getattr(self, "_glitch_total", 7)
        remain = max(0.0, float(getattr(self, "_glitch_remaining", 0.0)))
        paused = bool(getattr(self, "_glitch_paused", False))
        text = f"GLITCH {idx+1}/{total}  \u4e2d\u65ad\u4e2d" if paused else f"GLITCH {idx+1}/{total}  \u6b21\u307e\u3067{remain:.1f}s"
        try:
            self.after(0, self.var_glitch_status.set, text)
        except Exception:
            self.var_glitch_status.set(text)

    def _update_glitch_buttons(self):
        amp = self.var_amp.get()
        mode = self.var_mode.get()
        is_glitch = (mode == "Position" and amp == "\u30b0\u30ea\u30c3\u30c1")
        running = bool(getattr(self, "_glitch_running", False))
        paused = bool(getattr(self, "_glitch_paused", False))
        phase = getattr(self, "_glitch_phase", "idle")

        self.ent_glitch.config(state=("normal" if is_glitch else "disabled"))
        if not is_glitch:
            self.btn_glitch_pause.config(state="disabled")
            self.btn_glitch_resume.config(state="disabled")
            self.btn_glitch_stop.config(state="disabled")
            self.var_glitch_status.set("")
            return

        self.btn_glitch_pause.config(state=("normal" if running and not paused and phase == "countdown" else "disabled"))
        self.btn_glitch_resume.config(state=("normal" if running and paused else "disabled"))
        self.btn_glitch_stop.config(state=("normal" if running else "disabled"))

    def _update_glitch_buttons_async(self):
        try:
            self.after(0, self._update_glitch_buttons)
        except Exception:
            self._update_glitch_buttons()
