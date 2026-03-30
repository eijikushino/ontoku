import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import queue
import time
import os
import json
import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from dc_char_definitions import (
    POSITION_TEST_POINTS, POSITION_EXPECTED_STRINGS, POSITION_DISPLAY_ORDER,
    LBC_TEST_POINTS, LBC_EXPECTED_STRINGS,
    MONI_TEST_POINTS, MONI_EXPECTED_STRINGS, MONI_DISPLAY_ORDER,
    DMM_RANGE_POSITION, DMM_RANGE_LBC, DMM_RANGE_MONI,
)


class DCCharTab(ttk.Frame):
    """DC特性タブ: POSTION/LBC/moni の自動計測・XLSX保存・PNG出力"""

    SETTINGS_KEY = "dc_char"

    def __init__(self, parent, gpib_3458a, gpib_3499b, datagen_manager, test_tab):
        super().__init__(parent)
        self.gpib_dmm = gpib_3458a
        self.gpib_scanner = gpib_3499b
        self.datagen = datagen_manager
        self.test_tab = test_tab
        self.scanner_slot = "1"

        # Threading
        self.is_running = False
        self._stop_event = threading.Event()
        self._worker_thread = None
        self._update_queue = queue.Queue()

        # Settings
        self.save_dir = tk.StringVar(value='dc_char_data')
        self.settle_time_var = tk.DoubleVar(value=0.3)
        self.switch_delay_sec = tk.DoubleVar(value=1.0)  # Pattern Testと共有
        self.test_type = tk.StringVar(value='Position')  # Position / LBC / moni

        # Results
        self._results = {}  # {def_index: {'position': [...], 'lbc': [...], 'moni': [...]}}

        self._load_settings()
        self._create_widgets()

        # Auto-save settings on change
        for var in [self.save_dir, self.settle_time_var]:
            var.trace_add("write", lambda *_: self._save_settings())

    # ==================== Settings ====================
    def _load_settings(self):
        try:
            with open('app_settings.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            s = config.get(self.SETTINGS_KEY, {})
            if 'save_dir' in s:
                self.save_dir.set(s['save_dir'])
            if 'settle_time' in s:
                self.settle_time_var.set(s['settle_time'])
            # スキャナ切替時間はPattern Testと共有（measurement_window.switch_delay_sec）
            sw = config.get("measurement_window", {}).get("switch_delay_sec", 1.0)
            self.switch_delay_sec.set(max(0.4, sw))
        except Exception:
            pass

    def _save_settings(self):
        try:
            try:
                with open('app_settings.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
            except Exception:
                config = {}
            config[self.SETTINGS_KEY] = {
                'save_dir': self.save_dir.get(),
                'settle_time': self.settle_time_var.get(),
            }
            with open('app_settings.json', 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    # ==================== UI ====================
    def _create_widgets(self):
        # 左右分割
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=5, pady=5)

        left = ttk.Frame(paned, width=340)
        right = ttk.Frame(paned)
        paned.add(left, weight=0)
        paned.add(right, weight=1)

        # === 左パネル ===
        # 試験種別選択（排他）
        type_frame = ttk.LabelFrame(left, text="試験種別", padding=5)
        type_frame.pack(fill="x", padx=5, pady=(5, 2))

        type_row = ttk.Frame(type_frame)
        type_row.pack(fill="x")
        for t in ["Position", "LBC", "moni"]:
            ttk.Radiobutton(type_row, text=t, variable=self.test_type,
                            value=t).pack(side="left", padx=(0, 15))

        # 試験設定
        settings_frame = ttk.LabelFrame(left, text="試験設定", padding=5)
        settings_frame.pack(fill="x", padx=5, pady=2)

        row = ttk.Frame(settings_frame)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="DAC安定待ち:").pack(side="left")
        ttk.Entry(row, textvariable=self.settle_time_var, width=6).pack(side="left", padx=(5, 2))
        ttk.Label(row, text="sec").pack(side="left")

        row2 = ttk.Frame(settings_frame)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="スキャナ切替:").pack(side="left")
        ttk.Label(row2, textvariable=self.switch_delay_sec,
                  font=("Arial", 9, "bold")).pack(side="left", padx=(5, 2))
        ttk.Label(row2, text="sec (Pattern Test共有)").pack(side="left")

        # DEF選択 & スキャナCH (test_tabの変数を共有)
        def_frame = ttk.LabelFrame(left, text="DEF選択 & スキャナCH", padding=5)
        def_frame.pack(fill="x", padx=5, pady=2)

        if hasattr(self.test_tab, 'def_check_vars'):
            ch_values = ["ー"] + [f"CH{i:02d}" for i in range(1, 21)]
            for i in range(len(self.test_tab.def_check_vars)):
                row = ttk.Frame(def_frame)
                row.pack(fill="x", pady=1)
                ttk.Checkbutton(row, text=f"DEF{i}",
                                variable=self.test_tab.def_check_vars[i]).pack(side="left")
                ttk.Label(row, text="POS:").pack(side="left", padx=(10, 0))
                ttk.Combobox(row, textvariable=self.test_tab.scanner_channels_pos[i],
                             values=ch_values, width=5, state="readonly").pack(side="left", padx=2)
                ttk.Label(row, text="NEG:").pack(side="left")
                ttk.Combobox(row, textvariable=self.test_tab.scanner_channels_neg[i],
                             values=ch_values, width=5, state="readonly").pack(side="left", padx=2)

        # === 右パネル ===
        # 実行制御（右上に配置、保存先パスを広く表示）
        ctrl_frame = ttk.LabelFrame(right, text="実行制御", padding=5)
        ctrl_frame.pack(fill="x", padx=5, pady=(5, 2))

        btn_row = ttk.Frame(ctrl_frame)
        btn_row.pack(fill="x", pady=2)
        self.btn_start = ttk.Button(btn_row, text="開始", command=self._start_measurement)
        self.btn_start.pack(side="left", padx=2)
        self.btn_stop = ttk.Button(btn_row, text="停止", command=self._stop_measurement, state="disabled")
        self.btn_stop.pack(side="left", padx=2)
        self.var_progress_label = tk.StringVar(value="待機中")
        ttk.Label(btn_row, textvariable=self.var_progress_label,
                  font=("Arial", 9)).pack(side="left", padx=(15, 0))

        save_row = ttk.Frame(ctrl_frame)
        save_row.pack(fill="x", pady=2)
        ttk.Label(save_row, text="保存先:").pack(side="left")
        ttk.Entry(save_row, textvariable=self.save_dir).pack(side="left", fill="x", expand=True, padx=2)
        ttk.Button(save_row, text="参照", width=4,
                   command=self._browse_save_dir).pack(side="left")

        # 結果サマリー（測定電圧表示）
        result_frame = ttk.LabelFrame(right, text="結果サマリー", padding=5)
        result_frame.pack(fill="both", expand=True, padx=5, pady=2)

        columns = ('def_name', 'part', 'code', 'voltage', 'expected', 'error', 'judge')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=18)
        self.tree.heading('def_name', text='DEF')
        self.tree.heading('part', text='部')
        self.tree.heading('code', text='ｺｰﾄﾞ')
        self.tree.heading('voltage', text='計測値(V)')
        self.tree.heading('expected', text='期待値')
        self.tree.heading('error', text='誤差(V)')
        self.tree.heading('judge', text='結果')
        self.tree.column('def_name', width=55, anchor='center')
        self.tree.column('part', width=40, anchor='center')
        self.tree.column('code', width=65, anchor='center')
        self.tree.column('voltage', width=120, anchor='e')
        self.tree.column('expected', width=120, anchor='center')
        self.tree.column('error', width=110, anchor='e')
        self.tree.column('judge', width=40, anchor='center')
        self.tree.tag_configure('ok', background='#d5f5e3')
        self.tree.tag_configure('ng', background='#f5b7b1')

        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

    def _browse_save_dir(self):
        d = filedialog.askdirectory(title="保存先ディレクトリ選択")
        if d:
            self.save_dir.set(d)

    # ==================== 計測制御 ====================
    def _start_measurement(self):
        selected = self._get_selected_defs()
        if not selected:
            messagebox.showerror("エラー", "DEFを1つ以上選択してください")
            return

        # 接続確認
        if not self.gpib_dmm or not self.gpib_dmm.connected:
            messagebox.showerror("エラー", "DMM (3458A) が未接続です")
            return
        if not self.gpib_scanner or not self.gpib_scanner.connected:
            messagebox.showerror("エラー", "スキャナー (3499B) が未接続です")
            return
        if not self.datagen or not self.datagen.is_connected():
            messagebox.showerror("エラー", "DataGenが未接続です")
            return

        # 保存先ディレクトリ作成
        save_dir = self.save_dir.get()
        os.makedirs(save_dir, exist_ok=True)

        # 初期化
        self.is_running = True
        self._stop_event.clear()
        self._results = {}
        self.tree.delete(*self.tree.get_children())
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")

        self._worker_thread = threading.Thread(
            target=self._measurement_worker, args=(selected,), daemon=True
        )
        self._worker_thread.start()
        self._poll_updates()

    def _stop_measurement(self):
        self._stop_event.set()

    def _poll_updates(self):
        try:
            for _ in range(50):
                msg_type, data = self._update_queue.get_nowait()
                if msg_type == 'progress':
                    self.var_progress_label.set(data)
                elif msg_type == 'row':
                    # (def_name, part, code, voltage_str, expected_str, error_str, judge)
                    tag = 'ok' if data[-1] == 'OK' else 'ng'
                    self.tree.insert('', 'end', values=data, tags=(tag,))
                    self.tree.see(self.tree.get_children()[-1])
                elif msg_type == 'done':
                    self.var_progress_label.set("完了")
                    self.is_running = False
                    self.btn_start.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    if data:
                        messagebox.showinfo("完了", data)
                    return
                elif msg_type == 'error':
                    self.var_progress_label.set("エラー")
                    self.is_running = False
                    self.btn_start.config(state="normal")
                    self.btn_stop.config(state="disabled")
                    messagebox.showerror("エラー", str(data))
                    return
        except queue.Empty:
            pass
        if self.is_running:
            self.after(100, self._poll_updates)

    # ==================== DEF選択 ====================
    def _get_selected_defs(self):
        selected = []
        if hasattr(self.test_tab, 'def_check_vars'):
            for i, var in enumerate(self.test_tab.def_check_vars):
                if var.get():
                    pos_ch = self.test_tab.scanner_channels_pos[i].get() \
                        if i < len(self.test_tab.scanner_channels_pos) else "ー"
                    neg_ch = self.test_tab.scanner_channels_neg[i].get() \
                        if i < len(self.test_tab.scanner_channels_neg) else "ー"
                    if (pos_ch == "ー" or not pos_ch) and (neg_ch == "ー" or not neg_ch):
                        continue
                    selected.append({
                        'index': i, 'name': f"DEF{i}",
                        'pos_channel': pos_ch, 'neg_channel': neg_ch
                    })
        return selected

    def _get_serial_number(self, def_index):
        try:
            with open('app_settings.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            sn = config.get('comm_profiles', {}).get('1', {}).get('serial_numbers', {})
            return sn.get(f'DEF{def_index}_sn', f'DEF{def_index}')
        except Exception:
            return f'DEF{def_index}'

    # ==================== ハードウェア操作 ====================
    def _scanner_cpon(self):
        """スキャナー全チャンネルOPEN (cpon)"""
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000
        try:
            self.gpib_scanner.write(f":system:cpon {self.scanner_slot}")
        finally:
            self.gpib_scanner.instrument.timeout = orig_timeout

    def _switch_scanner(self, channel_addr, switch_delay):
        """スキャナCH切替（Pattern Testと同じcpon方式）
        cpon → sleep(delay/2) → *OPC? → CLOSE → sleep(delay/2)
        """
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000
        try:
            # cponで全チャンネルOPEN
            self.gpib_scanner.write(f":system:cpon {self.scanner_slot}")
            time.sleep(switch_delay / 2)
            # *OPC?でcpon完了を確認
            self.gpib_scanner.query("*OPC?")
            # 対象CHをCLOSE
            self.gpib_scanner.write(f"CLOSE ({channel_addr})")
            time.sleep(switch_delay / 2)
        finally:
            self.gpib_scanner.instrument.timeout = orig_timeout

    def _measure_voltage(self):
        orig_timeout = self.gpib_dmm.instrument.timeout
        self.gpib_dmm.instrument.timeout = 5000
        try:
            success, response = self.gpib_dmm.query("TRIG SGL")
            if success:
                return float(response.strip())
        except Exception:
            pass
        finally:
            self.gpib_dmm.instrument.timeout = orig_timeout
        return None

    def _datagen_send(self, cmd):
        self.datagen.send_command(cmd)
        time.sleep(0.05)

    def _ch_addr(self, ch_str):
        """CH番号文字列からスキャナアドレスを生成 (例: 'CH01' → '@101')"""
        num = ch_str.replace("CH", "")
        return f"@{self.scanner_slot}{num}"

    # ==================== 計測ワーカー ====================
    def _measurement_worker(self, selected_defs):
        try:
            settle = self.settle_time_var.get()
            switch_delay = self.switch_delay_sec.get()
            test_type = self.test_type.get()

            # 初期化
            self._scanner_cpon()
            self._datagen_send("gen stop")
            self._datagen_send("func alt")
            self._datagen_send("alt s sa")

            # DMM設定: NPLC 10
            self.gpib_dmm.write("NPLC 10")
            time.sleep(0.1)

            # DMMレンジ設定
            if test_type == "Position":
                self.gpib_dmm.write("DCV 1000")
            else:
                self.gpib_dmm.write("DCV AUTO")
            time.sleep(0.1)

            self._datagen_send("gen start")

            for def_info in selected_defs:
                if self._stop_event.is_set():
                    break

                def_name = def_info['name']
                def_idx = def_info['index']
                self._results[def_idx] = {'position': [], 'lbc': [], 'moni': []}

                if test_type == "Position":
                    self._run_position_test(def_info, settle, switch_delay)
                elif test_type == "LBC":
                    self._run_lbc_test(def_info, settle, switch_delay)
                else:  # moni
                    self._run_moni_test(def_info, settle, switch_delay)

            # 後片付け
            self._scanner_cpon()
            self._datagen_send("alt s sa")

            if self._stop_event.is_set():
                self._update_queue.put(('done', "計測を中断しました"))
                return

            # 保存
            save_dir = self.save_dir.get()
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            saved_files = []

            for def_idx, results in self._results.items():
                sn = self._get_serial_number(def_idx)

                # XLSX保存
                xlsx_path = self._save_xlsx(results, save_dir, sn, timestamp)
                saved_files.append(xlsx_path)

                # PNG保存（計測順→Excel表示順にリオーダー）
                sheet_name = test_type.lower()
                if sheet_name == 'position':
                    display_data = self._reorder_for_display(results['position'], POSITION_DISPLAY_ORDER)
                elif sheet_name == 'moni':
                    display_data = self._reorder_for_display(results['moni'], MONI_DISPLAY_ORDER)
                else:
                    display_data = results['lbc']
                png_path = self._generate_table_png(sheet_name, display_data, save_dir, sn, timestamp)
                saved_files.append(png_path)

            msg = f"保存完了:\n" + "\n".join(saved_files)
            self._update_queue.put(('done', msg))

        except Exception as e:
            self._update_queue.put(('error', str(e)))

    def _run_position_test(self, def_info, settle, switch_delay):
        """POSTION計測"""
        def_name = def_info['name']
        def_idx = def_info['index']
        current_pole = None
        for pi, tp in enumerate(POSITION_TEST_POINTS):
            if self._stop_event.is_set():
                break
            self._update_queue.put(('progress', f"{def_name} / POSTION / {tp.part} {tp.display_code}"))

            pole = "p" if tp.part == "POS" else "n"
            self._datagen_send(f"alt a {tp.address_code} ci {pole}")
            time.sleep(settle)

            if tp.part != current_pole:
                ch = def_info['pos_channel'] if tp.part == "POS" else def_info['neg_channel']
                self._switch_scanner(self._ch_addr(ch), switch_delay)
                current_pole = tp.part

            voltage = self._measure_voltage()

            if voltage is not None:
                error = voltage - tp.expected
                error_pct = (error / tp.expected * 100) if tp.expected != 0 else 0
                judge = "OK" if abs(error) <= tp.tolerance else "NG"
            else:
                error, error_pct, judge = None, None, "NG"

            self._results[def_idx]['position'].append({
                'part': tp.part, 'code': tp.display_code,
                'voltage': voltage, 'expected_str': POSITION_EXPECTED_STRINGS[pi],
                'error': error, 'error_pct': error_pct, 'judge': judge,
            })

            v_str = f"{voltage:.6f}" if voltage is not None else "---"
            e_str = f"{error:.6f}" if error is not None else "---"
            self._update_queue.put(('row', (
                def_name, tp.part, tp.display_code, v_str,
                POSITION_EXPECTED_STRINGS[pi], e_str, judge
            )))

    def _run_lbc_test(self, def_info, settle, switch_delay):
        """LBC計測"""
        def_name = def_info['name']
        def_idx = def_info['index']
        for li, tp in enumerate(LBC_TEST_POINTS):
            if self._stop_event.is_set():
                break
            self._update_queue.put(('progress', f"{def_name} / LBC / ATT{tp.att} {tp.display_code}"))

            self._datagen_send(f"alt a {tp.address_code} cii p")
            time.sleep(settle)

            # POS計測
            self._switch_scanner(self._ch_addr(def_info['pos_channel']), switch_delay)
            voltage_pos = self._measure_voltage()

            # NEG計測
            self._switch_scanner(self._ch_addr(def_info['neg_channel']), switch_delay)
            voltage_neg = self._measure_voltage()

            if voltage_pos is not None and voltage_neg is not None:
                error_pos = voltage_pos - tp.expected_pos
                error_neg = voltage_neg - tp.expected_neg
                judge = "OK" if abs(error_pos) <= tp.tolerance and abs(error_neg) <= tp.tolerance else "NG"
            else:
                error_pos, error_neg, judge = None, None, "NG"

            self._results[def_idx]['lbc'].append({
                'att': tp.att, 'code': tp.display_code,
                'voltage_pos': voltage_pos, 'voltage_neg': voltage_neg,
                'expected_str': LBC_EXPECTED_STRINGS[li],
                'error_pos': error_pos, 'error_neg': error_neg, 'judge': judge,
            })

            vp = f"{voltage_pos:.6f}" if voltage_pos is not None else "---"
            vn = f"{voltage_neg:.6f}" if voltage_neg is not None else "---"
            ep = f"{error_pos:.6f}" if error_pos is not None else "---"
            self._update_queue.put(('row', (
                def_name, f"ATT{tp.att}", tp.display_code,
                f"P:{vp} N:{vn}", LBC_EXPECTED_STRINGS[li], f"P:{ep}", judge
            )))

    def _run_moni_test(self, def_info, settle, switch_delay):
        """moni計測"""
        def_name = def_info['name']
        def_idx = def_info['index']
        current_pole = None
        for mi, tp in enumerate(MONI_TEST_POINTS):
            if self._stop_event.is_set():
                break
            self._update_queue.put(('progress', f"{def_name} / moni / {tp.part} {tp.display_code}"))

            pole = "p" if tp.part == "POS" else "n"
            self._datagen_send(f"alt a {tp.address_code} ci {pole}")
            time.sleep(settle)

            if tp.part != current_pole:
                ch = def_info['pos_channel'] if tp.part == "POS" else def_info['neg_channel']
                self._switch_scanner(self._ch_addr(ch), switch_delay)
                current_pole = tp.part

            voltage = self._measure_voltage()

            if voltage is not None:
                error = voltage - tp.expected
                error_pct = (error / tp.expected * 100) if tp.expected != 0 else 0
                judge = "OK" if abs(error) <= tp.tolerance else "NG"
            else:
                error, error_pct, judge = None, None, "NG"

            self._results[def_idx]['moni'].append({
                'part': tp.part, 'code': tp.display_code,
                'voltage': voltage, 'expected_str': MONI_EXPECTED_STRINGS[mi],
                'error': error, 'error_pct': error_pct, 'judge': judge,
            })

            v_str = f"{voltage:.6f}" if voltage is not None else "---"
            e_str = f"{error:.6f}" if error is not None else "---"
            self._update_queue.put(('row', (
                def_name, tp.part, tp.display_code, v_str,
                MONI_EXPECTED_STRINGS[mi], e_str, judge
            )))

    # ==================== 表示順リオーダー ====================
    def _reorder_for_display(self, data, order):
        """計測順の結果リストをExcel表示順に並べ替え"""
        return [data[i] for i in order]

    # ==================== XLSX保存 ====================
    def _save_xlsx(self, results, save_dir, serial_no, timestamp):
        wb = openpyxl.Workbook()

        # 計測順→Excel表示順にリオーダー
        pos_display = self._reorder_for_display(results['position'], POSITION_DISPLAY_ORDER)
        moni_display = self._reorder_for_display(results['moni'], MONI_DISPLAY_ORDER)

        # POSTION sheet
        ws = wb.active
        ws.title = "POSTION"
        self._write_position_sheet(ws, pos_display, serial_no)

        # LBC sheet (計測順=表示順なのでそのまま)
        ws_lbc = wb.create_sheet("LBC")
        self._write_lbc_sheet(ws_lbc, results['lbc'], serial_no)

        # moni sheet
        ws_moni = wb.create_sheet("moni")
        self._write_moni_sheet(ws_moni, moni_display, serial_no)

        filepath = os.path.join(save_dir, f"{serial_no}_DC特性_{timestamp}.xlsx")
        wb.save(filepath)
        return filepath

    def _write_position_sheet(self, ws, data, serial_no):
        thin = Side(style='thin')
        border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
        hdr_font = Font(name='MS Pゴシック', size=9, bold=True)
        dat_font = Font(name='MS Pゴシック', size=9)
        center = Alignment(horizontal='center', vertical='center')
        hdr_fill = PatternFill(start_color='D4E6F1', end_color='D4E6F1', fill_type='solid')
        ok_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
        ng_fill = PatternFill(start_color='F5B7B1', end_color='F5B7B1', fill_type='solid')

        headers = ['', 'ﾕﾆｯﾄNo.', '部', '入力ｺｰﾄﾞ', '出力電圧 (V)', '期待値±誤差 (V)', '誤差 (V)', '誤差(%)', '結果']
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = hdr_font
            cell.alignment = center
            cell.border = border_all
            cell.fill = hdr_fill

        for i, d in enumerate(data):
            r = i + 2
            ws.cell(row=r, column=1, value='').border = border_all
            ws.cell(row=r, column=2, value=serial_no if i == 0 else '').font = dat_font
            ws.cell(row=r, column=2).border = border_all
            ws.cell(row=r, column=2).alignment = center
            ws.cell(row=r, column=3, value=d['part']).font = dat_font
            ws.cell(row=r, column=3).border = border_all
            ws.cell(row=r, column=3).alignment = center
            ws.cell(row=r, column=4, value=d['code']).font = dat_font
            ws.cell(row=r, column=4).border = border_all
            ws.cell(row=r, column=4).alignment = center
            ws.cell(row=r, column=5, value=d['voltage']).font = dat_font
            ws.cell(row=r, column=5).border = border_all
            ws.cell(row=r, column=5).number_format = '0.000000'
            ws.cell(row=r, column=6, value=d['expected_str']).font = dat_font
            ws.cell(row=r, column=6).border = border_all
            ws.cell(row=r, column=6).alignment = center
            ws.cell(row=r, column=7, value=d['error']).font = dat_font
            ws.cell(row=r, column=7).border = border_all
            ws.cell(row=r, column=7).number_format = '0.000000'
            pct = d.get('error_pct')
            ws.cell(row=r, column=8, value=pct if pct else '').font = dat_font
            ws.cell(row=r, column=8).border = border_all
            ws.cell(row=r, column=8).number_format = '0.000000'
            jcell = ws.cell(row=r, column=9, value=d['judge'])
            jcell.font = dat_font
            jcell.border = border_all
            jcell.alignment = center
            jcell.fill = ok_fill if d['judge'] == 'OK' else ng_fill

        for c, w in {1: 4, 2: 14, 3: 6, 4: 12, 5: 18, 6: 18, 7: 18, 8: 14, 9: 8}.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    def _write_lbc_sheet(self, ws, data, serial_no):
        thin = Side(style='thin')
        border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
        hdr_font = Font(name='MS Pゴシック', size=9, bold=True)
        dat_font = Font(name='MS Pゴシック', size=9)
        center = Alignment(horizontal='center', vertical='center')
        hdr_fill = PatternFill(start_color='D4E6F1', end_color='D4E6F1', fill_type='solid')
        ok_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
        ng_fill = PatternFill(start_color='F5B7B1', end_color='F5B7B1', fill_type='solid')

        headers = ['', 'ﾕﾆｯﾄNo.', 'LBC ATT', '入力ｺｰﾄﾞ', 'POS出力電圧(V)', 'NEG出力電圧(V)',
                   '期待値(V)', 'POS 誤差(V)', 'NEG 誤差(V)', '結果']
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = hdr_font
            cell.alignment = center
            cell.border = border_all
            cell.fill = hdr_fill

        for i, d in enumerate(data):
            r = i + 2
            ws.cell(row=r, column=1, value='').border = border_all
            ws.cell(row=r, column=2, value=serial_no if i == 0 else '').font = dat_font
            ws.cell(row=r, column=2).border = border_all
            ws.cell(row=r, column=2).alignment = center
            ws.cell(row=r, column=3, value=d['att']).font = dat_font
            ws.cell(row=r, column=3).border = border_all
            ws.cell(row=r, column=3).alignment = center
            ws.cell(row=r, column=4, value=d['code']).font = dat_font
            ws.cell(row=r, column=4).border = border_all
            ws.cell(row=r, column=4).alignment = center
            ws.cell(row=r, column=5, value=d['voltage_pos']).font = dat_font
            ws.cell(row=r, column=5).border = border_all
            ws.cell(row=r, column=5).number_format = '0.000000'
            ws.cell(row=r, column=6, value=d['voltage_neg']).font = dat_font
            ws.cell(row=r, column=6).border = border_all
            ws.cell(row=r, column=6).number_format = '0.000000'
            ws.cell(row=r, column=7, value=d['expected_str']).font = dat_font
            ws.cell(row=r, column=7).border = border_all
            ws.cell(row=r, column=7).alignment = center
            ws.cell(row=r, column=8, value=d['error_pos']).font = dat_font
            ws.cell(row=r, column=8).border = border_all
            ws.cell(row=r, column=8).number_format = '0.000000'
            ws.cell(row=r, column=9, value=d['error_neg']).font = dat_font
            ws.cell(row=r, column=9).border = border_all
            ws.cell(row=r, column=9).number_format = '0.000000'
            jcell = ws.cell(row=r, column=10, value=d['judge'])
            jcell.font = dat_font
            jcell.border = border_all
            jcell.alignment = center
            jcell.fill = ok_fill if d['judge'] == 'OK' else ng_fill

        for c, w in {1: 4, 2: 14, 3: 8, 4: 12, 5: 18, 6: 18, 7: 18, 8: 18, 9: 18, 10: 8}.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    def _write_moni_sheet(self, ws, data, serial_no):
        thin = Side(style='thin')
        border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
        hdr_font = Font(name='MS Pゴシック', size=9, bold=True)
        dat_font = Font(name='MS Pゴシック', size=9)
        center = Alignment(horizontal='center', vertical='center')
        hdr_fill = PatternFill(start_color='D4E6F1', end_color='D4E6F1', fill_type='solid')
        ok_fill = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
        ng_fill = PatternFill(start_color='F5B7B1', end_color='F5B7B1', fill_type='solid')

        headers = ['', 'ﾕﾆｯﾄNo.', '部', '入力ｺｰﾄﾞ', '出力電圧(V)', '期待値±誤差(V)(<1%)',
                   '誤差(V)', '誤差(%)', '結果']
        for c, h in enumerate(headers, 1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = hdr_font
            cell.alignment = center
            cell.border = border_all
            cell.fill = hdr_fill

        for i, d in enumerate(data):
            r = i + 2
            ws.cell(row=r, column=1, value='').border = border_all
            ws.cell(row=r, column=2, value=serial_no if i == 0 else '').font = dat_font
            ws.cell(row=r, column=2).border = border_all
            ws.cell(row=r, column=2).alignment = center
            ws.cell(row=r, column=3, value=d['part']).font = dat_font
            ws.cell(row=r, column=3).border = border_all
            ws.cell(row=r, column=3).alignment = center
            ws.cell(row=r, column=4, value=d['code']).font = dat_font
            ws.cell(row=r, column=4).border = border_all
            ws.cell(row=r, column=4).alignment = center
            ws.cell(row=r, column=5, value=d['voltage']).font = dat_font
            ws.cell(row=r, column=5).border = border_all
            ws.cell(row=r, column=5).number_format = '0.000000'
            ws.cell(row=r, column=6, value=d['expected_str']).font = dat_font
            ws.cell(row=r, column=6).border = border_all
            ws.cell(row=r, column=6).alignment = center
            ws.cell(row=r, column=7, value=d['error']).font = dat_font
            ws.cell(row=r, column=7).border = border_all
            ws.cell(row=r, column=7).number_format = '0.000000'
            pct = d.get('error_pct')
            ws.cell(row=r, column=8, value=pct if pct else '').font = dat_font
            ws.cell(row=r, column=8).border = border_all
            ws.cell(row=r, column=8).number_format = '0.000000'
            jcell = ws.cell(row=r, column=9, value=d['judge'])
            jcell.font = dat_font
            jcell.border = border_all
            jcell.alignment = center
            jcell.fill = ok_fill if d['judge'] == 'OK' else ng_fill

        for c, w in {1: 4, 2: 14, 3: 6, 4: 12, 5: 18, 6: 18, 7: 18, 8: 14, 9: 8}.items():
            ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = w

    # ==================== PNG生成 ====================
    def _generate_table_png(self, sheet_name, data, save_dir, serial_no, timestamp):
        # 日本語フォント設定
        plt.rcParams['font.family'] = 'MS Gothic'

        if sheet_name == 'position':
            title = "DC特性 - POSTION"
            headers = ['部', '入力ｺｰﾄﾞ', '出力電圧(V)', '期待値±誤差(V)', '誤差(V)', '誤差(%)', '結果']
            rows = []
            for d in data:
                v = f"{d['voltage']:.6f}" if d['voltage'] is not None else ""
                e = f"{d['error']:.6f}" if d['error'] is not None else ""
                ep = f"{d['error_pct']:.6f}" if d.get('error_pct') is not None else ""
                rows.append([d['part'], d['code'], v, d['expected_str'], e, ep, d['judge']])
        elif sheet_name == 'lbc':
            title = "DC特性 - LBC"
            headers = ['ATT', '入力ｺｰﾄﾞ', 'POS電圧(V)', 'NEG電圧(V)', '期待値(V)', 'POS誤差(V)', 'NEG誤差(V)', '結果']
            rows = []
            for d in data:
                vp = f"{d['voltage_pos']:.6f}" if d['voltage_pos'] is not None else ""
                vn = f"{d['voltage_neg']:.6f}" if d['voltage_neg'] is not None else ""
                ep = f"{d['error_pos']:.6f}" if d['error_pos'] is not None else ""
                en = f"{d['error_neg']:.6f}" if d['error_neg'] is not None else ""
                rows.append([d['att'], d['code'], vp, vn, d['expected_str'], ep, en, d['judge']])
        else:  # moni
            title = "DC特性 - moni"
            headers = ['部', '入力ｺｰﾄﾞ', '出力電圧(V)', '期待値±誤差(V)', '誤差(V)', '誤差(%)', '結果']
            rows = []
            for d in data:
                v = f"{d['voltage']:.6f}" if d['voltage'] is not None else ""
                e = f"{d['error']:.6f}" if d['error'] is not None else ""
                ep = f"{d['error_pct']:.6f}" if d.get('error_pct') is not None else ""
                rows.append([d['part'], d['code'], v, d['expected_str'], e, ep, d['judge']])

        # セル色
        cell_colors = []
        for row in rows:
            row_colors = []
            for cell in row:
                if cell == 'NG':
                    row_colors.append('#f5b7b1')
                elif cell == 'OK':
                    row_colors.append('#d5f5e3')
                else:
                    row_colors.append('white')
            cell_colors.append(row_colors)

        fig_height = max(2, len(rows) * 0.35 + 1.5)
        fig, ax = plt.subplots(figsize=(12, fig_height))
        ax.axis('off')
        ax.set_title(f"{title}  ({serial_no})", fontsize=12, fontweight='bold', pad=20)

        table = ax.table(
            cellText=rows,
            colLabels=headers,
            cellColours=cell_colors,
            colColours=['#d4e6f1'] * len(headers),
            loc='center',
            cellLoc='center',
        )
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        table.scale(1.0, 1.4)

        # ヘッダー行を太字に
        for j in range(len(headers)):
            table[0, j].set_text_props(fontweight='bold')

        fig.tight_layout()
        sheet_upper = sheet_name.upper() if sheet_name != 'moni' else 'moni'
        filepath = os.path.join(save_dir, f"{serial_no}_DC特性_{sheet_upper}_{timestamp}.png")
        fig.savefig(filepath, dpi=150, bbox_inches='tight', pad_inches=0.1)
        plt.close(fig)
        return filepath
