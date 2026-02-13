import tkinter as tk
import time
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox

# DACプリセット値
DAC_PRESETS = {
    "full":   {"P": "FFFFF", "L": "FFFF"},
    "center": {"P": "80000", "L": "8000"},
    "zero":   {"P": "00000", "L": "0000"},
}


class DACTab(ttk.Frame):
    def __init__(self, parent, gpib_controller, serial_manager):
        super().__init__(parent)
        self.gpib = gpib_controller
        self.serial_mgr = serial_manager  # 通信1用SerialManager
        self.show_response = True  # レスポンス表示ON/OFF

        self.create_widgets()

    def create_widgets(self):
        # メインコンテナ
        main_container = ttk.Frame(self)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # ===== 左側:操作パネル(3列構成) =====
        left_frame = ttk.Frame(main_container)
        left_frame.pack(side="left", anchor="n", fill="both", expand=True, padx=(0, 10))

        # 色バー付き見出し(通常フォント)
        header_frame = tk.Frame(left_frame, bd=0, highlightthickness=0)
        header_frame.pack(anchor="w", pady=(0, 6))

        color_bar = tk.Frame(header_frame, bg="#81D4FA", width=8, height=20, bd=0, highlightthickness=0)
        color_bar.pack(side="left", padx=(0, 6))
        color_bar.pack_propagate(False)

        ttk.Label(header_frame, text="DAC操作(通信1)").pack(side="left")

        # 3列のフレームを作成
        columns_frame = ttk.Frame(left_frame)
        columns_frame.pack(fill="both", expand=True)

        # 第1列(左)
        col1_frame = ttk.Frame(columns_frame)
        col1_frame.pack(side="left", anchor="n", padx=(0, 10))

        # 第2列(中央)
        col2_frame = ttk.Frame(columns_frame)
        col2_frame.pack(side="left", anchor="n", padx=(0, 10))

        # 第3列(右) - レスポンス表示
        col3_frame = ttk.Frame(columns_frame)
        col3_frame.pack(side="left", anchor="n", fill="both", expand=True)

        # ===== 第1列: DEF選択とコマンド群 =====
        # ---- DEF選択 (横2列レイアウト) ----
        def_box = ttk.LabelFrame(col1_frame, text="DEF選択", padding=8)
        def_box.pack(anchor="w", pady=(0, 8), fill="x")

        self.def_vars = [tk.BooleanVar(value=(i == 0)) for i in range(6)]  # 既定はDEF0のみON
        for i in range(6):
            r, c = divmod(i, 2)
            ttk.Checkbutton(def_box, text=f"DEF{i}", variable=self.def_vars[i]).grid(
                row=r, column=c, padx=4, pady=1, sticky="w")

        # ---- コマンド群 ----
        cmd_box = ttk.LabelFrame(col1_frame, text="コマンド", padding=8)
        cmd_box.pack(anchor="w", pady=(0, 8), fill="x")
        self._grid_buttons(cmd_box, [
            ("test", lambda: self._send("test")),
            ("cal",  lambda: self._send("cal")),
            ("adc",  lambda: self._send("adc")),
        ], cols=2)

        # ---- 中止 ----
        stop_box = ttk.LabelFrame(col1_frame, text="中止", padding=8)
        stop_box.pack(anchor="w", pady=(0, 8), fill="x")
        self._grid_buttons(stop_box, [
            ("test q", lambda: self._send("test q")),
            ("cal q",  lambda: self._send("cal q")),
        ], cols=2)

        # ---- ステータス ----
        stat_box = ttk.LabelFrame(col1_frame, text="ステータス", padding=8)
        stat_box.pack(anchor="w", pady=(0, 8), fill="x")
        self._grid_buttons(stat_box, [
            ("test s", lambda: self._send("test s")),
            ("cal s",  lambda: self._send("cal s")),
        ], cols=2)

        # ---- レポート ----
        rep_box = ttk.LabelFrame(col1_frame, text="レポート", padding=8)
        rep_box.pack(anchor="w", pady=(0, 8), fill="x")
        self._grid_buttons(rep_box, [
            ("test r", lambda: self._send("test r")),
            ("cal r",  lambda: self._send("cal r")),
        ], cols=2)

        # ---- 手動コマンド ----
        manual_box = ttk.LabelFrame(col1_frame, text="手動コマンド", padding=8)
        manual_box.pack(anchor="w", pady=(0, 8), fill="x")

        manual_row = ttk.Frame(manual_box)
        manual_row.pack(fill="x")
        self.manual_cmd_var = tk.StringVar()
        manual_entry = ttk.Entry(manual_row, textvariable=self.manual_cmd_var, width=16)
        manual_entry.pack(side="left", padx=(0, 4))
        manual_entry.bind("<Return>", lambda e: self._send_manual())
        ttk.Button(manual_row, text="送信", width=4, command=self._send_manual).pack(side="left")

        # ===== 第2列: データセット & 出荷チェック =====
        # ---- データセット ----
        dataset_box = ttk.LabelFrame(col2_frame, text="データセット", padding=8)
        dataset_box.pack(anchor="w", pady=(0, 8), fill="x")

        # remote/local
        rl_frame = ttk.Frame(dataset_box)
        rl_frame.pack(anchor="w", pady=(0, 8))
        ttk.Button(rl_frame, text="remote", command=lambda: self._send("remote")).pack(side="left", padx=2)
        ttk.Button(rl_frame, text="local", command=lambda: self._send("local")).pack(side="left", padx=2)

        # 種別選択(Position/LBC)
        type_frame = ttk.Frame(dataset_box)
        type_frame.pack(anchor="w", pady=(0, 8))
        ttk.Label(type_frame, text="種別:").pack(side="left", padx=(0, 4))
        self.dac_type_var = tk.StringVar(value="P")
        ttk.Radiobutton(type_frame, text="Position", variable=self.dac_type_var, value="P").pack(side="left", padx=4)
        ttk.Radiobutton(type_frame, text="LBC", variable=self.dac_type_var, value="L").pack(side="left", padx=4)

        # フリー入力
        free_frame = ttk.Frame(dataset_box)
        free_frame.pack(anchor="w", pady=(0, 8), fill="x")
        ttk.Label(free_frame, text="値(HEX):").pack(side="left", padx=(0, 4))
        self.dac_value_entry = ttk.Entry(free_frame, width=10)
        self.dac_value_entry.pack(side="left", padx=(0, 4))
        ttk.Button(free_frame, text="送信", command=self._send_dac_free).pack(side="left")

        # 設定値読み込みボタン
        ttk.Button(dataset_box, text="設定値読み込み", command=self._read_dac_value).pack(fill="x", pady=(0, 4))

        # プリセットボタン(横並び)
        preset_frame = ttk.Frame(dataset_box)
        preset_frame.pack(anchor="w", pady=(4, 0))
        ttk.Label(preset_frame, text="プリセット:").pack(side="left", padx=(0, 2))
        ttk.Button(preset_frame, text="+Full", width=5,
                   command=lambda: self._send_dac_preset("full")).pack(side="left", padx=(0, 2))
        ttk.Button(preset_frame, text="Center", width=5,
                   command=lambda: self._send_dac_preset("center")).pack(side="left", padx=(0, 2))
        ttk.Button(preset_frame, text="-Full", width=5,
                   command=lambda: self._send_dac_preset("zero")).pack(side="left")

        # ===== 第3列:レスポンス表示 =====
        # 見出し
        ttk.Label(col3_frame, text="レスポンス表示(通信1)").pack(anchor="w", pady=(0, 4))

        # テキストエリア
        self.response_area = ScrolledText(
            col3_frame, height=32, width=40, font=("Consolas", 10), wrap="none"
        )
        self.response_area.pack(fill="both", expand=True, pady=(0, 6))

        # レスポンス下部コントロール
        resp_ctrl = ttk.Frame(col3_frame)
        resp_ctrl.pack(fill="x")

        self.var_show_recv = tk.BooleanVar(value=True)
        ttk.Checkbutton(resp_ctrl, text="レスポンス表示",
                        variable=self.var_show_recv,
                        command=self._on_show_recv_change).pack(side="left")
        ttk.Button(resp_ctrl, text="表示クリア", command=self._clear_response).pack(side="right")

    # ---------- ボタン配置ヘルパー ----------
    def _grid_buttons(self, parent, buttons, cols=2):
        """
        buttons: [(text, command), ...]
        cols   : 1行あたりの列数(既定2)
        """
        for i, (txt, cmd) in enumerate(buttons):
            r, c = divmod(i, cols)
            ttk.Button(parent, text=txt, command=cmd).grid(row=r, column=c, padx=4, pady=2, sticky="ew")
        # 横方向に均等配置
        for c in range(cols):
            parent.grid_columnconfigure(c, weight=1)

    def _get_selected_defs(self) -> list:
        """選択されたDEF番号リストを取得"""
        return [i for i, var in enumerate(self.def_vars) if var.get()]

    def _on_show_recv_change(self):
        """レスポンス表示ON/OFF切り替え"""
        self.show_response = self.var_show_recv.get()

    # ---------- コマンド送信 ----------
    def _send(self, base_command: str):
        """選択されたDEFに対してコマンド送信"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "通信1: シリアルポート未接続")
            return

        for def_num in self._get_selected_defs():
            cmd = f"DEF {def_num} {base_command}"
            self._append_text(f"[SEND] {cmd}")
            try:
                self.serial_mgr.flush_input()
                self.serial_mgr.write((cmd + "\r").encode("utf-8"))
                if self.show_response:
                    self._read_response()
                else:
                    time.sleep(0.05)
            except Exception as e:
                self._append_text(f"[ERROR] write failed: {e}")

    def _send_manual(self):
        """手動コマンド送信(DEFプレフィックスなし、生コマンド)"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "通信1: シリアルポート未接続")
            return

        cmd = self.manual_cmd_var.get().strip()
        if not cmd:
            return
        self.manual_cmd_var.set("")

        self._append_text(f"[SEND] {cmd}")
        try:
            self.serial_mgr.flush_input()
            self.serial_mgr.write((cmd + "\r").encode("utf-8"))
            self._read_response()
        except Exception as e:
            self._append_text(f"[ERROR] write failed: {e}")

    # ---------- DACデータセット送信 ----------
    def _send_dac_free(self):
        """フリー入力値でDACコマンド送信"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "通信1: シリアルポート未接続")
            return

        dac_type = self.dac_type_var.get()  # "P" or "L"
        hex_value = self.dac_value_entry.get().strip().upper()

        if not hex_value:
            messagebox.showwarning("入力エラー", "HEX値を入力してください")
            return

        try:
            int(hex_value, 16)
        except ValueError:
            messagebox.showwarning("入力エラー", "HEX値が不正です")
            return

        expected_len = 5 if dac_type == "P" else 4
        if len(hex_value) != expected_len:
            messagebox.showwarning("入力エラー", f"{dac_type}は{expected_len}桁のHEX値を入力してください")
            return

        self._send_dac_command(dac_type, hex_value)

    def _send_dac_preset(self, preset_type: str):
        """プリセット値でDACコマンド送信"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "通信1: シリアルポート未接続")
            return

        dac_type = self.dac_type_var.get()
        hex_value = DAC_PRESETS[preset_type][dac_type]
        self._send_dac_command(dac_type, hex_value)

    def _send_dac_command(self, dac_type: str, hex_value: str):
        """選択されたDEFに対してDACコマンド送信"""
        for def_num in self._get_selected_defs():
            cmd = f"DEF {def_num} DAC {dac_type} {hex_value}"
            self._append_text(f"[SEND] {cmd}")
            try:
                self.serial_mgr.flush_input()
                self.serial_mgr.write((cmd + "\r").encode("utf-8"))
                if self.show_response:
                    self._read_response()
                else:
                    time.sleep(0.05)
            except Exception as e:
                self._append_text(f"[ERROR] write failed: {e}")

    # ---------- レスポンス読み取り ----------
    def _read_dac_value(self):
        """選択されたDEFのDAC設定値を読み込み"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "通信1: シリアルポート未接続")
            return

        dac_type = self.dac_type_var.get()

        for def_num in self._get_selected_defs():
            cmd = f"DEF {def_num} DAC {dac_type}"
            self._append_text(f"[SEND] {cmd}")
            try:
                self.serial_mgr.flush_input()
                self.serial_mgr.write((cmd + "\r").encode("utf-8"))
                self._read_response_with_status()
            except Exception as e:
                self._append_text(f"[ERROR] write failed: {e}")

    def _read_response(self):
        """
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
                    self.response_area.insert(tk.END, line_buffer + "\n")
                    self.response_area.see(tk.END)
                    line_buffer = ""
            elif ch == ">":
                self.response_area.insert(tk.END, "> \n")
                self.response_area.see(tk.END)
                break
            else:
                line_buffer += ch

    def _read_response_with_status(self):
        """
        設定値読み込み用のレスポンス読み取り。'>'の後の*local* *relay_OFF*も読み取る。
        """
        line_buffer = ""
        prompt_found = False
        timeout = time.time() + 3
        no_data_count = 0

        while time.time() < timeout:
            ch = self.serial_mgr.read()

            if not ch:
                if prompt_found:
                    no_data_count += 1
                    if no_data_count > 20:  # 約0.2秒データが来なければ終了
                        break
                time.sleep(0.01)
                continue

            no_data_count = 0

            if ch in ("\r", "\n"):
                if line_buffer.strip():
                    self.response_area.insert(tk.END, line_buffer + "\n")
                    self.response_area.see(tk.END)
                    line_buffer = ""
            elif ch == ">":
                if line_buffer.strip():
                    self.response_area.insert(tk.END, line_buffer + "> \n")
                else:
                    self.response_area.insert(tk.END, "> \n")
                self.response_area.see(tk.END)
                line_buffer = ""
                prompt_found = True
            else:
                line_buffer += ch

        # 最後にバッファに残っているデータがあれば表示
        if line_buffer.strip():
            self.response_area.insert(tk.END, line_buffer + "\n")
            self.response_area.see(tk.END)

    # ---------- ユーティリティ ----------
    def _append_text(self, message: str):
        """レスポンスエリアにテキスト追加"""
        self.response_area.insert(tk.END, message + "\n")
        self.response_area.see(tk.END)

    def _clear_response(self):
        """レスポンス表示をクリア"""
        self.response_area.delete(1.0, tk.END)
