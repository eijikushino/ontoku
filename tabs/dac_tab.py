import tkinter as tk
import time
import threading
import queue
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
        self.serial_mgr = serial_manager  # DEFシリアル用SerialManager
        self.show_response = True  # レスポンス表示ON/OFF

        # 常駐リーダー用
        self._text_queue = queue.Queue()
        self._reader_running = False
        self._need_recv_header = False

        # 複数 DEF 順次送信用 (DEF試験 dac_control_tab から移植)
        # reader が「>」プロンプトを見た時刻 / データを受けた時刻を記録し、
        # _wait_for_prompt_idle で「プロンプト後 idle_sec 秒無音 = 応答完了」と判定
        self._last_prompt_time = 0.0
        self._last_data_time = 0.0
        self._prompt_event = threading.Event()

        self.create_widgets()

        # 常駐リーダースレッド起動
        self._reader_running = True
        threading.Thread(target=self._reader_loop, daemon=True).start()
        self._poll_text_queue()

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

        ttk.Label(header_frame, text="DEF操作").pack(side="left")

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
        ttk.Label(col3_frame, text="レスポンス表示(DEFシリアル)").pack(anchor="w", pady=(0, 4))

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
        """選択されたDEFに対してコマンド送信

        - 単一 DEF: 即時 write (従来動作)
        - 複数 DEF: バックグラウンドスレッドで順次 write + 応答完了待ち
          (DEF試験 dac_control_tab.py から移植。バス競合回避)
        """
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "DEFシリアル: ポート未接続")
            return

        defs = self._get_selected_defs()
        if not defs:
            return

        if len(defs) == 1:
            # 単一 DEF: 従来どおり即時 write
            def_num = defs[0]
            cmd = f"DEF {def_num} {base_command}"
            raw = (cmd + "\r").encode("utf-8")
            self._append_text(f"[SEND] {cmd}")
            try:
                self._need_recv_header = True
                self.serial_mgr.flush_input()
                self.serial_mgr.write(raw)
            except Exception as e:
                self._append_text(f"[ERROR] write failed: {e}")
            return

        # 複数 DEF: bg スレッドで順次送信 + 応答完了待ち
        is_long_cmd = base_command in ("cal", "test")
        idle_sec = 5.0 if is_long_cmd else 0.5
        max_wait = 60.0 if is_long_cmd else 10.0

        def _bg():
            for def_num in defs:
                cmd = f"DEF {def_num} {base_command}"
                raw = (cmd + "\r").encode("utf-8")
                # 応答完了検出用の時刻をリセット
                self._last_prompt_time = 0.0
                self._prompt_event.clear()
                self._text_queue.put(f"[SEND] {cmd}")
                try:
                    self._need_recv_header = True
                    self.serial_mgr.flush_input()
                    self.serial_mgr.write(raw)
                except Exception as e:
                    self._text_queue.put(f"[ERROR] write failed: {e}")
                    continue
                # 応答完了 (プロンプト後 idle_sec 秒無音) を待つ
                self._wait_for_prompt_idle(
                    idle_sec=idle_sec, max_wait=max_wait)

        threading.Thread(target=_bg, daemon=True).start()

    def _send_manual(self):
        """手動コマンド送信(DEFプレフィックスなし、生コマンド)"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "DEFシリアル: ポート未接続")
            return

        cmd = self.manual_cmd_var.get().strip()
        if not cmd:
            return
        self.manual_cmd_var.set("")

        self._append_text(f"[SEND] {cmd}")
        try:
            self._need_recv_header = True
            self.serial_mgr.write((cmd + "\r").encode("utf-8"))
        except Exception as e:
            self._append_text(f"[ERROR] write failed: {e}")

    # ---------- DACデータセット送信 ----------
    def _send_dac_free(self):
        """フリー入力値でDACコマンド送信"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "DEFシリアル: ポート未接続")
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
            messagebox.showwarning("通信エラー", "DEFシリアル: ポート未接続")
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
                self._need_recv_header = True
                self.serial_mgr.write((cmd + "\r").encode("utf-8"))
            except Exception as e:
                self._append_text(f"[ERROR] write failed: {e}")

    # ---------- レスポンス読み取り ----------
    def _read_dac_value(self):
        """選択されたDEFのDAC設定値を読み込み"""
        if not self.serial_mgr.is_connected():
            messagebox.showwarning("通信エラー", "DEFシリアル: ポート未接続")
            return

        dac_type = self.dac_type_var.get()

        for def_num in self._get_selected_defs():
            cmd = f"DEF {def_num} DAC {dac_type}"
            self._append_text(f"[SEND] {cmd}")
            try:
                self._need_recv_header = True
                self.serial_mgr.write((cmd + "\r").encode("utf-8"))
            except Exception as e:
                self._append_text(f"[ERROR] write failed: {e}")

    # ---------- 常駐リーダー ----------
    def _reader_loop(self):
        """常駐リーダースレッド: シリアルデータを読み続けて表示"""
        line_buffer = ""
        skip_space = False
        while self._reader_running:
            if not self.serial_mgr.is_connected():
                time.sleep(0.1)
                continue
            chunk = self._read_serial_chunk()
            if not chunk:
                time.sleep(0.01)
                continue
            for ch in chunk:
                if ch in ("\r", "\n"):
                    if line_buffer.strip():
                        if self.show_response:
                            if self._need_recv_header:
                                self._text_queue.put("[RECV]")
                                self._need_recv_header = False
                            self._text_queue.put(line_buffer)
                        self._last_data_time = time.time()
                    line_buffer = ""
                    skip_space = False
                elif ch == ">" and not line_buffer.strip():
                    # プロンプト検出: 応答完了検出用に時刻を記録
                    self._last_prompt_time = time.time()
                    self._prompt_event.set()
                    line_buffer = ""
                    skip_space = True
                elif skip_space and ch == " ":
                    skip_space = False
                else:
                    skip_space = False
                    line_buffer += ch
                    if line_buffer.strip():
                        self._last_data_time = time.time()

    def _wait_for_prompt_idle(self, idle_sec: float = 0.5,
                              max_wait: float = 10.0) -> bool:
        """「>」受信後 idle_sec 秒間データ無しになるまで待つ。

        複数 DEF 順次送信時に「1つの DEF の応答完了」を検出するために使用。
        max_wait 秒以内に idle 条件が満たされなければタイムアウトして False。
        戻り値: True=idle 成立, False=タイムアウト
        """
        start = time.time()
        while time.time() - start < max_wait:
            pt = self._last_prompt_time
            dt = self._last_data_time
            now = time.time()
            # 送信開始後に「>」が来ていて、その後にデータが来ておらず
            # かつ idle_sec 経過していれば応答完了
            if pt > start and pt >= dt and now - pt > idle_sec:
                return True
            time.sleep(0.05)
        return False

    def _read_serial_chunk(self):
        """受信バッファの全データを一括読み取り"""
        with self.serial_mgr.lock:
            if self.serial_mgr.is_connected():
                n = self.serial_mgr.ser.in_waiting
                if n > 0:
                    try:
                        return self.serial_mgr.ser.read(n).decode("utf-8", errors="ignore")
                    except Exception:
                        return ""
        return ""

    def _poll_text_queue(self):
        """メインスレッドでキューからテキストを表示"""
        try:
            for _ in range(200):
                message = self._text_queue.get_nowait()
                self.response_area.insert(tk.END, message + "\n")
                self.response_area.see(tk.END)
        except queue.Empty:
            pass
        self.after(16, self._poll_text_queue)

    # ---------- ユーティリティ ----------
    def _append_text(self, message: str):
        """スレッドセーフにレスポンスエリアへテキスト挿入"""
        if threading.current_thread() is threading.main_thread():
            self.response_area.insert(tk.END, message + "\n")
            self.response_area.see(tk.END)
        else:
            self._text_queue.put(message)

    def _clear_response(self):
        """レスポンス表示をクリア"""
        self.response_area.delete(1.0, tk.END)
