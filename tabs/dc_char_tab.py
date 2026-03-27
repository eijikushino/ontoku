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
    POSITION_TEST_POINTS, POSITION_EXPECTED_STRINGS,
    LBC_TEST_POINTS, LBC_EXPECTED_STRINGS,
    MONI_TEST_POINTS, MONI_EXPECTED_STRINGS,
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
        # 試験設定
        settings_frame = ttk.LabelFrame(left, text="試験設定", padding=5)
        settings_frame.pack(fill="x", padx=5, pady=(5, 2))

        row = ttk.Frame(settings_frame)
        row.pack(fill="x", pady=2)
        ttk.Label(row, text="DAC安定待ち:").pack(side="left")
        ttk.Entry(row, textvariable=self.settle_time_var, width=6).pack(side="left", padx=(5, 2))
        ttk.Label(row, text="sec").pack(side="left")

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

        # 実行制御
        ctrl_frame = ttk.LabelFrame(left, text="実行制御", padding=5)
        ctrl_frame.pack(fill="x", padx=5, pady=2)

        btn_row = ttk.Frame(ctrl_frame)
        btn_row.pack(fill="x", pady=2)
        self.btn_start = ttk.Button(btn_row, text="開始", command=self._start_measurement)
        self.btn_start.pack(side="left", padx=2)
        self.btn_stop = ttk.Button(btn_row, text="停止", command=self._stop_measurement, state="disabled")
        self.btn_stop.pack(side="left", padx=2)

        save_row = ttk.Frame(ctrl_frame)
        save_row.pack(fill="x", pady=2)
        ttk.Label(save_row, text="保存先:").pack(side="left")
        ttk.Entry(save_row, textvariable=self.save_dir, width=20).pack(side="left", padx=2)
        ttk.Button(save_row, text="参照", width=4,
                   command=self._browse_save_dir).pack(side="left")

        # === 右パネル ===
        # 進捗
        progress_frame = ttk.LabelFrame(right, text="進捗", padding=5)
        progress_frame.pack(fill="x", padx=5, pady=(5, 2))

        self.var_progress_label = tk.StringVar(value="待機中")
        ttk.Label(progress_frame, textvariable=self.var_progress_label,
                  font=("Arial", 10)).pack(anchor="w")

        self.var_progress_value = tk.StringVar(value="")
        ttk.Label(progress_frame, textvariable=self.var_progress_value,
                  font=("Consolas", 10, "bold")).pack(anchor="w")

        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill="x", pady=(2, 0))

        # 結果サマリー
        result_frame = ttk.LabelFrame(right, text="結果サマリー", padding=5)
        result_frame.pack(fill="both", expand=True, padx=5, pady=2)

        columns = ('def_name', 'test', 'result')
        self.tree = ttk.Treeview(result_frame, columns=columns, show='headings', height=15)
        self.tree.heading('def_name', text='DEF')
        self.tree.heading('test', text='試験')
        self.tree.heading('result', text='結果')
        self.tree.column('def_name', width=80)
        self.tree.column('test', width=100)
        self.tree.column('result', width=60)
        self.tree.tag_configure('ok', background='#d5f5e3')
        self.tree.tag_configure('ng', background='#f5b7b1')
        self.tree.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(result_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

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
        if not self.gpib_dmm or not self.gpib_dmm.is_connected:
            messagebox.showerror("エラー", "DMM (3458A) が未接続です")
            return
        if not self.gpib_scanner or not self.gpib_scanner.is_connected:
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
                    label, value, current, total = data
                    self.var_progress_label.set(label)
                    self.var_progress_value.set(value)
                    self.progress_bar['maximum'] = total
                    self.progress_bar['value'] = current
                elif msg_type == 'result':
                    def_name, test_name, judge = data
                    tag = 'ok' if judge == 'OK' else 'ng'
                    self.tree.insert('', 'end', values=(def_name, test_name, judge), tags=(tag,))
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
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000
        try:
            self.gpib_scanner.write(f":system:cpon {self.scanner_slot}")
        finally:
            self.gpib_scanner.instrument.timeout = orig_timeout

    def _switch_scanner(self, channel_addr, switch_delay):
        orig_timeout = self.gpib_scanner.instrument.timeout
        self.gpib_scanner.instrument.timeout = 5000
        try:
            self.gpib_scanner.write(f":system:cpon {self.scanner_slot}")
            time.sleep(switch_delay / 2)
            self.gpib_scanner.query("*OPC?")
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

    def _set_dmm_range(self, range_str):
        self.gpib_dmm.write(f"DCV {range_str}")
        time.sleep(0.1)

    # ==================== 計測ワーカー ====================
    def _measurement_worker(self, selected_defs):
        try:
            settle = self.settle_time_var.get()
            total_steps = len(selected_defs) * (len(POSITION_TEST_POINTS) + len(LBC_TEST_POINTS) + len(MONI_TEST_POINTS))
            current_step = 0

            # 初期化
            self._scanner_cpon()
            self._datagen_send("gen stop")
            self._datagen_send("func alt")
            self._datagen_send("alt s sa")
            self._datagen_send("gen start")

            for def_info in selected_defs:
                if self._stop_event.is_set():
                    break

                def_name = def_info['name']
                def_idx = def_info['index']
                self._results[def_idx] = {'position': [], 'lbc': [], 'moni': []}

                # --- POSTION計測 ---
                self._set_dmm_range(DMM_RANGE_POSITION)
                pos_all_ok = True
                for pi, tp in enumerate(POSITION_TEST_POINTS):
                    if self._stop_event.is_set():
                        break
                    current_step += 1
                    self._update_queue.put(('progress', (
                        f"{def_name} / POSTION / {tp.part} {tp.display_code}",
                        "", current_step, total_steps
                    )))

                    # DAC設定
                    pole = "p" if tp.part == "POS" else "n"
                    self._datagen_send(f"alt a {tp.address_code} ci {pole}")
                    time.sleep(settle)

                    # スキャナ切替 & 計測
                    ch = def_info['pos_channel'] if tp.part == "POS" else def_info['neg_channel']
                    ch_addr = f"{self.scanner_slot}{ch}"
                    self._switch_scanner(ch_addr, settle)
                    voltage = self._measure_voltage()

                    # 判定
                    if voltage is not None:
                        error = voltage - tp.expected
                        error_pct = (error / tp.expected * 100) if tp.expected != 0 else 0
                        judge = "OK" if abs(error) <= tp.tolerance else "NG"
                    else:
                        error, error_pct, judge = None, None, "NG"

                    if judge != "OK":
                        pos_all_ok = False

                    self._results[def_idx]['position'].append({
                        'part': tp.part, 'code': tp.display_code,
                        'voltage': voltage, 'expected_str': POSITION_EXPECTED_STRINGS[pi],
                        'error': error, 'error_pct': error_pct, 'judge': judge,
                    })

                    self._update_queue.put(('progress', (
                        f"{def_name} / POSTION / {tp.part} {tp.display_code}",
                        f"{voltage:.6f} V" if voltage else "計測失敗",
                        current_step, total_steps
                    )))

                self._update_queue.put(('result', (def_name, 'POSTION', 'OK' if pos_all_ok else 'NG')))

                # --- LBC計測 ---
                if self._stop_event.is_set():
                    break
                self._set_dmm_range(DMM_RANGE_LBC)
                lbc_all_ok = True
                for li, tp in enumerate(LBC_TEST_POINTS):
                    if self._stop_event.is_set():
                        break
                    current_step += 1
                    self._update_queue.put(('progress', (
                        f"{def_name} / LBC / ATT{tp.att} {tp.display_code}",
                        "", current_step, total_steps
                    )))

                    # ATT設定 + DAC設定
                    # ATT切替はcmodeコマンドで実施（LBC ATT設定）
                    self._datagen_send(f"alt a {tp.address_code} cii p")
                    time.sleep(settle)

                    # POS計測
                    ch_pos = f"{self.scanner_slot}{def_info['pos_channel']}"
                    self._switch_scanner(ch_pos, settle)
                    voltage_pos = self._measure_voltage()

                    # NEG計測
                    ch_neg = f"{self.scanner_slot}{def_info['neg_channel']}"
                    self._switch_scanner(ch_neg, settle)
                    voltage_neg = self._measure_voltage()

                    # 判定
                    if voltage_pos is not None and voltage_neg is not None:
                        error_pos = voltage_pos - tp.expected_pos
                        error_neg = voltage_neg - tp.expected_neg
                        judge = "OK" if abs(error_pos) <= tp.tolerance and abs(error_neg) <= tp.tolerance else "NG"
                    else:
                        error_pos, error_neg, judge = None, None, "NG"

                    if judge != "OK":
                        lbc_all_ok = False

                    self._results[def_idx]['lbc'].append({
                        'att': tp.att, 'code': tp.display_code,
                        'voltage_pos': voltage_pos, 'voltage_neg': voltage_neg,
                        'expected_str': LBC_EXPECTED_STRINGS[li],
                        'error_pos': error_pos, 'error_neg': error_neg, 'judge': judge,
                    })

                    disp = f"POS:{voltage_pos:.6f} NEG:{voltage_neg:.6f}" if voltage_pos and voltage_neg else "計測失敗"
                    self._update_queue.put(('progress', (
                        f"{def_name} / LBC / ATT{tp.att} {tp.display_code}",
                        disp, current_step, total_steps
                    )))

                self._update_queue.put(('result', (def_name, 'LBC', 'OK' if lbc_all_ok else 'NG')))

                # --- moni計測 ---
                if self._stop_event.is_set():
                    break
                self._set_dmm_range(DMM_RANGE_MONI)
                moni_all_ok = True
                for mi, tp in enumerate(MONI_TEST_POINTS):
                    if self._stop_event.is_set():
                        break
                    current_step += 1
                    self._update_queue.put(('progress', (
                        f"{def_name} / moni / {tp.part} {tp.display_code}",
                        "", current_step, total_steps
                    )))

                    # DAC設定
                    pole = "p" if tp.part == "POS" else "n"
                    self._datagen_send(f"alt a {tp.address_code} ci {pole}")
                    time.sleep(settle)

                    # スキャナ切替 & 計測
                    ch = def_info['pos_channel'] if tp.part == "POS" else def_info['neg_channel']
                    ch_addr = f"{self.scanner_slot}{ch}"
                    self._switch_scanner(ch_addr, settle)
                    voltage = self._measure_voltage()

                    # 判定
                    if voltage is not None:
                        error = voltage - tp.expected
                        error_pct = (error / tp.expected * 100) if tp.expected != 0 else 0
                        judge = "OK" if abs(error) <= tp.tolerance else "NG"
                    else:
                        error, error_pct, judge = None, None, "NG"

                    if judge != "OK":
                        moni_all_ok = False

                    self._results[def_idx]['moni'].append({
                        'part': tp.part, 'code': tp.display_code,
                        'voltage': voltage, 'expected_str': MONI_EXPECTED_STRINGS[mi],
                        'error': error, 'error_pct': error_pct, 'judge': judge,
                    })

                    self._update_queue.put(('progress', (
                        f"{def_name} / moni / {tp.part} {tp.display_code}",
                        f"{voltage:.6f} V" if voltage else "計測失敗",
                        current_step, total_steps
                    )))

                self._update_queue.put(('result', (def_name, 'moni', 'OK' if moni_all_ok else 'NG')))

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

                # PNG保存
                for sheet_name in ['position', 'lbc', 'moni']:
                    png_path = self._generate_table_png(
                        sheet_name, results[sheet_name],
                        save_dir, sn, timestamp
                    )
                    saved_files.append(png_path)

            msg = f"保存完了:\n" + "\n".join(saved_files)
            self._update_queue.put(('done', msg))

        except Exception as e:
            self._update_queue.put(('error', str(e)))

    # ==================== XLSX保存 ====================
    def _save_xlsx(self, results, save_dir, serial_no, timestamp):
        wb = openpyxl.Workbook()

        # POSTION sheet
        ws = wb.active
        ws.title = "POSTION"
        self._write_position_sheet(ws, results['position'], serial_no)

        # LBC sheet
        ws_lbc = wb.create_sheet("LBC")
        self._write_lbc_sheet(ws_lbc, results['lbc'], serial_no)

        # moni sheet
        ws_moni = wb.create_sheet("moni")
        self._write_moni_sheet(ws_moni, results['moni'], serial_no)

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
