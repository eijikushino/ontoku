import tkinter as tk
from tkinter import ttk, messagebox
import matplotlib.pyplot as plt
import matplotlib
import numpy as np
from datetime import datetime
import re

# 日本語フォント設定（Windows）
matplotlib.rcParams['font.family'] = ['MS Gothic', 'Yu Gothic', 'Meiryo', 'sans-serif']


class LSBGraphPlotter:
    """
    CSVデータをLSB換算してグラフ化するクラス（Matplotlib最適化版）

    機能:
    - 電圧データをLSB値に変換
    - Codeごとに色分けしてプロット
    - plt.show()で別ウィンドウ表示（高速）
    - 横軸は時間表示（min, hour）
    - 縦軸のLSB/div設定可能
    - 大量データ対応（マーカー無効化、目盛り数制限）
    """

    def __init__(self, bit_precision, pos_full_voltage, neg_full_voltage, lsb_per_div=10, ref_mode="ideal",
                 yaxis_mode="auto", yaxis_min=None, yaxis_max=None,
                 skip_after_change=0, skip_first_data=False, skip_before_change=False):
        """
        Args:
            bit_precision: bit精度 (例: 24)
            pos_full_voltage: +Full基準電圧 (V)
            neg_full_voltage: -Full基準電圧 (V)
            lsb_per_div: 縦軸1divあたりのLSB数 (デフォルト: 10)
            ref_mode: 基準電圧計算モード
                      "ideal" - 理想値（コードから理論計算）
                      "all_avg" - 全平均（全データの平均）
                      "section_avg" - 区間別平均（各区間ごとの平均）
                      "first_avg" - 初回平均（最初の区間の平均）
            yaxis_mode: Y軸範囲モード ("auto" or "manual")
            yaxis_min: Y軸最小値 (LSB) - manualモード時のみ使用
            yaxis_max: Y軸最大値 (LSB) - manualモード時のみ使用
            skip_after_change: コード切替後にスキップする行数 (デフォルト: 0)
            skip_first_data: パターン開始最初のデータをスキップするか (デフォルト: False)
            skip_before_change: 切替わり直前のデータをスキップするか (デフォルト: False)
        """
        self.bit_precision = bit_precision
        self.pos_full_voltage = pos_full_voltage
        self.neg_full_voltage = neg_full_voltage
        self.lsb_per_div = lsb_per_div
        self.ref_mode = ref_mode
        self.yaxis_mode = yaxis_mode
        self.yaxis_min = yaxis_min
        self.yaxis_max = yaxis_max
        self.skip_after_change = skip_after_change
        self.skip_first_data = skip_first_data
        self.skip_before_change = skip_before_change

        # 1LSB電圧値を計算
        self.lsb_voltage = (pos_full_voltage - neg_full_voltage) / (2 ** bit_precision - 1)

    def extract_hex_from_code(self, code_str, dataset):
        """
        コード文字列からHEX値を抽出して10進数に変換

        Args:
            code_str: コード文字列 (例: "0x80000(C)", "Manual(80000)")
            dataset: "Position" or "LBC"

        Returns:
            10進数のコード値、抽出できない場合はNone
        """
        if not code_str or code_str == "---":
            return None

        # 括弧内の文字を抽出
        bracket_content = self.extract_bracket_content(code_str)

        # Manual の場合
        if "Manual" in code_str:
            hex_value = bracket_content.replace("Manual", "").strip().strip("()")
            if hex_value:
                try:
                    return int(hex_value, 16)
                except ValueError:
                    return None
            return None

        # HEX値として解釈を試みる
        try:
            return int(bracket_content, 16)
        except ValueError:
            # プリセット値のマッピング
            hex_values = {
                "Position": {"+": "FFFFF", "C": "80000", "-": "00000", "H": "FFFFF"},
                "LBC": {"+": "FFFF", "C": "8000", "-": "0000", "H": "FFFF"}
            }
            hex_val = hex_values.get(dataset, {}).get(bracket_content)
            if hex_val:
                return int(hex_val, 16)
            return None

    def calculate_ref_voltage(self, code_value, dataset, pole):
        """
        入力コードから基準電圧を計算

        Args:
            code_value: 10進数のコード値
            dataset: "Position" or "LBC"
            pole: "POS" or "NEG"

        Returns:
            基準電圧 (V)
        """
        # データセットに応じたbit精度
        bit_precision = 20 if dataset == "Position" else 16
        max_code = 2 ** bit_precision - 1
        voltage_range = self.pos_full_voltage - self.neg_full_voltage

        if pole == "POS":
            # POS: 基準電圧 = (-Full) + (code / max) × range
            ref_voltage = self.neg_full_voltage + (code_value / max_code) * voltage_range
        else:
            # NEG: 基準電圧 = (+Full) - (code / max) × range
            ref_voltage = self.pos_full_voltage - (code_value / max_code) * voltage_range

        return ref_voltage

    def voltage_to_lsb(self, voltage, pole, code_str=None, dataset=None):
        """
        電圧値をLSB値に変換

        Args:
            voltage: 測定電圧 (V)
            pole: "POS" or "NEG"
            code_str: コード文字列 (例: "0x80000(C)")
            dataset: "Position" or "LBC"

        Returns:
            LSB値
        """
        # コード値から基準電圧を計算
        if code_str and dataset:
            code_value = self.extract_hex_from_code(code_str, dataset)
            if code_value is not None:
                ref_voltage = self.calculate_ref_voltage(code_value, dataset, pole)
            else:
                # コード抽出失敗時はフォールバック
                ref_voltage = self.pos_full_voltage if pole == "POS" else self.neg_full_voltage
        else:
            # 旧互換: コード指定なしの場合
            ref_voltage = self.pos_full_voltage if pole == "POS" else self.neg_full_voltage

        lsb_value = (voltage - ref_voltage) / self.lsb_voltage
        return lsb_value

    def parse_timestamp(self, timestamp_str):
        """
        タイムスタンプ文字列をdatetimeオブジェクトに変換

        Args:
            timestamp_str: "2024-01-01 12:34:56.789" 形式の文字列

        Returns:
            datetimeオブジェクト
        """
        try:
            return datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S.%f")
        except ValueError:
            # ミリ秒なしのフォーマットも試す
            return datetime.strptime(timestamp_str, "%Y-%m-%d %H:%M:%S")

    def calculate_elapsed_time(self, timestamp, base_timestamp):
        """
        基準時刻からの経過時間を計算（分単位）

        Args:
            timestamp: 現在のタイムスタンプ
            base_timestamp: 基準タイムスタンプ

        Returns:
            経過時間（分）
        """
        delta = timestamp - base_timestamp
        return delta.total_seconds() / 60.0

    def format_time_label(self, minutes):
        """
        分を「hour min」形式に変換

        Args:
            minutes: 経過時間（分）

        Returns:
            フォーマットされた時間文字列（英語）
        """
        hours = int(minutes // 60)
        mins = int(minutes % 60)

        if hours > 0:
            return f"{hours}h {mins}min"
        else:
            return f"{mins}min"

    def extract_bracket_content(self, code_str):
        """
        コード文字列から括弧内の文字を抽出

        Args:
            code_str: コード文字列 (例: "0x00(+)", "0x01(-)", "0xFF(H)")

        Returns:
            括弧内の文字列 (例: "+", "-", "H")、括弧がない場合は元の文字列
        """
        # 括弧内の文字を抽出
        match = re.search(r'\(([^)]+)\)', code_str)
        if match:
            return match.group(1)
        else:
            # 括弧がない場合は元の文字列をそのまま返す
            return code_str

    def extract_data_from_csv(self, csv_data, serial, pole, column_name):
        """
        CSVデータから指定した列のデータを抽出してLSB変換

        Args:
            csv_data: CSVデータ (DictReaderで読み込んだリスト)
            serial: シリアルNo.
            pole: "POS" or "NEG"
            column_name: 列名 (例: "DFH903_POS")

        Returns:
            (elapsed_times, lsb_values, codes, datasets): タプル
                elapsed_times: 経過時間（分）のリスト
                lsb_values: LSB値のリスト
                codes: Codeのリスト
                datasets: DataSetのリスト (Position/LBC)
        """
        elapsed_times = []
        voltages = []
        codes = []
        datasets = []
        base_timestamp = None

        # 第1パス: 生データを収集
        for row in csv_data:
            if column_name in row and row[column_name]:
                try:
                    # タイムスタンプを解析
                    timestamp_str = row.get('Timestamp', '')
                    if not timestamp_str:
                        continue

                    timestamp = self.parse_timestamp(timestamp_str)

                    # 基準時刻の設定（最初のデータポイント）
                    if base_timestamp is None:
                        base_timestamp = timestamp

                    # 経過時間を計算
                    elapsed_min = self.calculate_elapsed_time(timestamp, base_timestamp)

                    # コードとデータセットを取得
                    code_str = row.get('Code', '')
                    dataset = row.get('DataSet', '')

                    # 電圧値を取得
                    voltage = float(row[column_name])

                    elapsed_times.append(elapsed_min)
                    voltages.append(voltage)
                    codes.append(code_str)
                    datasets.append(dataset)

                except (ValueError, KeyError) as e:
                    continue

        # スキップ処理（切替後、開始時、切替前）
        if (self.skip_after_change > 0 or self.skip_first_data or self.skip_before_change) and len(codes) > 0:
            skip_indices = self._get_skip_indices(codes, datasets)
            # スキップ対象でないインデックスのみ残す
            elapsed_times = [elapsed_times[i] for i in range(len(elapsed_times)) if i not in skip_indices]
            voltages = [voltages[i] for i in range(len(voltages)) if i not in skip_indices]
            codes = [codes[i] for i in range(len(codes)) if i not in skip_indices]
            datasets = [datasets[i] for i in range(len(datasets)) if i not in skip_indices]

        # 基準電圧モードに応じて計算
        if self.ref_mode == "all_avg":
            # 全平均モード: 全データの平均
            ref_voltages = self._calculate_all_average(voltages, codes, datasets)
        elif self.ref_mode == "section_avg":
            # 区間別平均モード: 各区間ごとの平均
            ref_voltages = self._calculate_section_average(voltages, codes, datasets)
        elif self.ref_mode == "first_avg":
            # 初回平均モード: 最初の区間の平均
            ref_voltages = self._calculate_first_section_average(voltages, codes, datasets)
        else:
            # 理想値モード
            ref_voltages = None

        # 第2パス: LSB値に変換
        lsb_values = []
        for i in range(len(voltages)):
            voltage = voltages[i]
            code_str = codes[i]
            dataset = datasets[i]

            if ref_voltages is not None:
                # 平均モード: 計算済みの基準電圧を使用
                ref_voltage = ref_voltages[i]
                lsb_value = (voltage - ref_voltage) / self.lsb_voltage
            else:
                # 理想値モード: コードから計算した理論値を基準に
                lsb_value = self.voltage_to_lsb(voltage, pole, code_str, dataset)

            lsb_values.append(lsb_value)

        return elapsed_times, lsb_values, codes, datasets

    def _get_skip_indices(self, codes, datasets):
        """
        スキップするインデックスを取得（3種類のスキップ条件を処理）

        Args:
            codes: Codeのリスト
            datasets: DataSetのリスト

        Returns:
            set: スキップするインデックスのセット
        """
        skip_indices = set()
        n = len(codes)
        if n == 0:
            return skip_indices

        # コード切り替え位置を検出
        change_positions = []  # 切り替わりが発生したインデックス（新しいコードの最初）
        prev_key = (codes[0], datasets[0])
        for i in range(1, n):
            key = (codes[i], datasets[i])
            if key != prev_key:
                change_positions.append(i)
            prev_key = key

        # 1. パターン開始最初のデータをスキップ（skip_after_change行数分）
        if self.skip_first_data and self.skip_after_change > 0:
            for i in range(min(self.skip_after_change, n)):
                skip_indices.add(i)

        # 2. コード切替後スキップ
        if self.skip_after_change > 0:
            for pos in change_positions:
                for i in range(pos, min(pos + self.skip_after_change, n)):
                    skip_indices.add(i)

        # 3. 切替わり直前をスキップ（1行）
        if self.skip_before_change:
            for pos in change_positions:
                if pos > 0:
                    skip_indices.add(pos - 1)

        return skip_indices

    def _calculate_all_average(self, voltages, codes, datasets):
        """
        全平均モード: コードごとの全データ平均電圧を各データポイントの基準電圧として返す

        Args:
            voltages: 電圧値のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト

        Returns:
            list: 各データポイントに対応する基準電圧のリスト
        """
        # コードごとに電圧を集計
        voltage_sums = {}
        voltage_counts = {}

        for i in range(len(voltages)):
            key = (codes[i], datasets[i])
            if key not in voltage_sums:
                voltage_sums[key] = 0.0
                voltage_counts[key] = 0
            voltage_sums[key] += voltages[i]
            voltage_counts[key] += 1

        # 平均を計算
        avg_voltages = {}
        for key in voltage_sums:
            avg_voltages[key] = voltage_sums[key] / voltage_counts[key]

        # 各データポイントに対応する基準電圧を返す
        ref_voltages = []
        for i in range(len(voltages)):
            key = (codes[i], datasets[i])
            ref_voltages.append(avg_voltages.get(key, voltages[i]))

        return ref_voltages

    def _calculate_section_average(self, voltages, codes, datasets):
        """
        区間別平均モード: 各区間ごとの平均電圧を基準電圧として返す

        Args:
            voltages: 電圧値のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト

        Returns:
            list: 各データポイントに対応する基準電圧のリスト
        """
        # 区間を検出（コードが変わるたびに新しい区間）
        sections = self._detect_sections(codes, datasets)

        # 各区間の平均を計算
        section_averages = {}
        for section_id, indices in sections.items():
            section_sum = sum(voltages[i] for i in indices)
            section_averages[section_id] = section_sum / len(indices)

        # 各データポイントに対応する基準電圧を返す
        ref_voltages = [0.0] * len(voltages)
        for section_id, indices in sections.items():
            avg = section_averages[section_id]
            for i in indices:
                ref_voltages[i] = avg

        return ref_voltages

    def _calculate_first_section_average(self, voltages, codes, datasets):
        """
        初回平均モード: 各コードの最初の区間の平均電圧を全データの基準電圧として返す

        Args:
            voltages: 電圧値のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト

        Returns:
            list: 各データポイントに対応する基準電圧のリスト
        """
        # 区間を検出
        sections = self._detect_sections(codes, datasets)

        # 各コードの最初の区間を特定し、その平均を計算
        first_section_avg = {}  # {(code, dataset): average}
        code_first_section = {}  # {(code, dataset): section_id}

        for section_id, indices in sections.items():
            if not indices:
                continue
            # section_id = (code, dataset, section_number)
            code, dataset, section_num = section_id
            key = (code, dataset)

            # 最初の区間のみ記録
            if key not in code_first_section:
                code_first_section[key] = section_id
                section_sum = sum(voltages[i] for i in indices)
                first_section_avg[key] = section_sum / len(indices)

        # 各データポイントに対応する基準電圧を返す
        ref_voltages = []
        for i in range(len(voltages)):
            key = (codes[i], datasets[i])
            ref_voltages.append(first_section_avg.get(key, voltages[i]))

        return ref_voltages

    def _detect_sections(self, codes, datasets):
        """
        コードの区間を検出する

        Args:
            codes: Codeのリスト
            datasets: DataSetのリスト

        Returns:
            dict: {(code, dataset, section_number): [indices]}
        """
        sections = {}
        section_counts = {}  # {(code, dataset): count}
        current_key = None

        for i in range(len(codes)):
            key = (codes[i], datasets[i])

            if key != current_key:
                # 新しい区間開始
                if key not in section_counts:
                    section_counts[key] = 0
                section_counts[key] += 1
                current_key = key

            section_id = (codes[i], datasets[i], section_counts[key])
            if section_id not in sections:
                sections[section_id] = []
            sections[section_id].append(i)

        return sections

    def _get_ref_mode_label(self):
        """基準電圧モードの表示名を取得"""
        mode_labels = {
            "ideal": "理想値",
            "all_avg": "全平均",
            "section_avg": "区間別平均",
            "first_avg": "初回平均"
        }
        return mode_labels.get(self.ref_mode, self.ref_mode)

    def create_plot_window(self, parent, serial, pole, elapsed_times, lsb_values, codes, datasets):
        """
        個別のグラフウィンドウを作成（Matplotlib最適化版）

        Args:
            parent: 親ウィンドウ（使用しない）
            serial: シリアルNo.
            pole: "POS" or "NEG"
            elapsed_times: 経過時間（分）のリスト
            lsb_values: LSB値のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト (Position/LBC)

        Returns:
            None（Matplotlibの別ウィンドウが開く）
        """
        if not elapsed_times:
            messagebox.showwarning("Warning", f"No data for {serial} {pole}")
            return None

        # グラフを作成
        fig, ax = plt.subplots(figsize=(10, 6))

        # Codeごとに色分けしてプロット（線のみ、マーカーなし）
        self._plot_by_code_with_lines(ax, elapsed_times, lsb_values, codes, datasets)

        # 軸ラベルとタイトル
        ax.set_xlabel('Time')
        ax.set_ylabel('Deviation [LSB]')
        ref_mode_label = self._get_ref_mode_label()
        ax.set_title(f'{serial} - {pole}  [基準: {ref_mode_label}]')
        ax.legend(loc='best')
        ax.grid(True, alpha=0.3)

        # X軸のフォーマット（一定間隔で時間表示）
        self._format_time_axis(ax, max(elapsed_times))

        # Y軸のフォーマット（LSB/div設定に基づく、目盛り数制限付き）
        self._format_lsb_axis(ax, lsb_values)

        # Y軸のオフセット表示を無効化（左上の科学的記数法を非表示）
        ax.ticklabel_format(style='plain', axis='y', useOffset=False)

        # レイアウト調整（ラベルが切れないように）
        fig.tight_layout()

        # Matplotlibの別ウィンドウで表示（高速）
        plt.show(block=False)

        return True  # 成功を返す

    def _plot_by_code_with_lines(self, ax, elapsed_times, lsb_values, codes, datasets):
        """
        Codeごとに色分けしてプロット（連続区間のみ線で結ぶ）- 最適化版

        Args:
            ax: Matplotlibのaxisオブジェクト
            elapsed_times: 経過時間（分）のリスト
            lsb_values: LSB値のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト (Position/LBC)
        """
        # NumPy配列に変換（高速化）
        times_array = np.array(elapsed_times)
        values_array = np.array(lsb_values)
        codes_array = np.array(codes)
        datasets_array = np.array(datasets)

        # unique_keysを出現順で生成（順序を保証）
        seen = set()
        unique_keys = []
        for key in zip(codes, datasets):
            if key not in seen:
                seen.add(key)
                unique_keys.append(key)

        # tab10は10色のパレット、インデックス0-9で順番に取得
        colors = [plt.cm.tab10(i % 10) for i in range(len(unique_keys))]

        # ラベル重複回避
        used_labels = set()

        for (code, dataset), color in zip(unique_keys, colors):
            # 凡例ラベルを作成
            if code == "---" or dataset == "---":
                legend_label = "---:---"
            else:
                prefix = "Position" if dataset == "Position" else "LBC"

                # Manual の場合
                if "Manual" in code:
                    bracket_content = self.extract_bracket_content(code)
                    hex_value = bracket_content.replace("Manual", "").strip().strip("()")
                    if hex_value:
                        legend_label = f"{prefix}:{hex_value}"
                    else:
                        legend_label = f"{prefix}:----"
                else:
                    bracket_content = self.extract_bracket_content(code)

                    try:
                        int(bracket_content, 16)
                        legend_label = f"{prefix}:{bracket_content}"
                    except ValueError:
                        hex_values = {
                            "Position": {"+": "FFFFF", "C": "80000", "-": "00000", "H": "FFFFF"},
                            "LBC": {"+": "FFFF", "C": "8000", "-": "0000", "H": "FFFF"}
                        }
                        hex_val = hex_values.get(dataset, {}).get(bracket_content, bracket_content)
                        legend_label = f"{prefix}:{hex_val}"

            # 該当するインデックスを抽出
            mask = (codes_array == code) & (datasets_array == dataset)
            indices = np.where(mask)[0]

            if len(indices) == 0:
                continue

            # 連続区間を検出してプロット
            # インデックスが連続している区間ごとに分割
            segments = []
            segment_start = 0
            for i in range(1, len(indices)):
                # インデックスが連続していない場合（1より大きい差がある場合）
                if indices[i] - indices[i-1] > 1:
                    segments.append(indices[segment_start:i])
                    segment_start = i
            segments.append(indices[segment_start:])

            # 各セグメントをプロット
            first_segment = True
            for segment in segments:
                seg_times = times_array[segment]
                seg_values = values_array[segment]

                if first_segment and legend_label not in used_labels:
                    ax.plot(seg_times, seg_values,
                           color=color,
                           linestyle='-',
                           linewidth=1.5,
                           label=legend_label,
                           alpha=0.7)
                    used_labels.add(legend_label)
                    first_segment = False
                else:
                    ax.plot(seg_times, seg_values,
                           color=color,
                           linestyle='-',
                           linewidth=1.5,
                           alpha=0.7)

    def _format_time_axis(self, ax, max_minutes):
        """
        X軸を時間フォーマットで表示（横向き）

        Args:
            ax: Matplotlibのaxisオブジェクト
            max_minutes: 最大経過時間（分）
        """
        # 適切な目盛り間隔を決定
        if max_minutes < 10:
            tick_interval = 1
        elif max_minutes < 60:
            tick_interval = 5
        elif max_minutes < 300:
            tick_interval = 30
        else:
            tick_interval = 60

        # 目盛り位置を計算
        ticks = np.arange(0, max_minutes + tick_interval, tick_interval)

        # 目盛りラベルを「hour min」形式に変換
        labels = [self.format_time_label(t) for t in ticks]

        ax.set_xticks(ticks)
        ax.set_xticklabels(labels, rotation=0, ha='center')

    def _format_lsb_axis(self, ax, lsb_values):
        """
        Y軸をLSB/div設定に基づいてフォーマット（目盛り数制限付き）

        Args:
            ax: Matplotlibのaxisオブジェクト
            lsb_values: LSB値のリスト
        """
        if not lsb_values:
            return

        # マニュアルモードの場合は設定値をそのまま使用
        if self.yaxis_mode == "manual" and self.yaxis_min is not None and self.yaxis_max is not None:
            ax.set_ylim(self.yaxis_min, self.yaxis_max)
            # 目盛りはlsb_per_divに基づいて設定
            ticks = np.arange(
                np.ceil(self.yaxis_min / self.lsb_per_div) * self.lsb_per_div,
                self.yaxis_max + self.lsb_per_div,
                self.lsb_per_div
            )
            # 範囲内の目盛りのみ
            ticks = ticks[ticks <= self.yaxis_max]
            ax.set_yticks(ticks)
            return

        # オートモード: データの範囲を取得
        min_lsb = min(lsb_values)
        max_lsb = max(lsb_values)

        data_range = max_lsb - min_lsb

        # 目盛り数を最大10個に制限するため、適切な間隔を計算
        max_ticks = 10
        if data_range / self.lsb_per_div > max_ticks:
            # データ範囲が広い場合、自動で間隔を調整
            actual_div = data_range / max_ticks
            # きりの良い数値に丸める（1, 2, 5, 10, 20, 50, 100, ...）
            magnitude = 10 ** np.floor(np.log10(actual_div))
            normalized = actual_div / magnitude
            if normalized <= 1:
                actual_div = magnitude
            elif normalized <= 2:
                actual_div = 2 * magnitude
            elif normalized <= 5:
                actual_div = 5 * magnitude
            else:
                actual_div = 10 * magnitude
        else:
            actual_div = self.lsb_per_div

        # 目盛りを設定
        min_tick = np.floor(min_lsb / actual_div) * actual_div
        max_tick = np.ceil(max_lsb / actual_div) * actual_div

        # 目盛りを生成
        ticks = np.arange(min_tick, max_tick + actual_div, actual_div)

        ax.set_yticks(ticks)
        ax.set_ylim(min_tick - actual_div * 0.5, max_tick + actual_div * 0.5)

    def update_plot_window(self, window):
        """
        既存のウィンドウを更新（Matplotlib版では再描画）

        Args:
            window: ウィンドウオブジェクト（未使用）
        """
        # Matplotlib plt.show(block=False)方式では更新不要
        pass

    def plot_csv_data(self, parent, csv_data, serial, pole):
        """
        CSVデータをプロットする（便利メソッド）

        Args:
            parent: 親ウィンドウ（使用しない）
            csv_data: CSVデータ
            serial: シリアルNo.
            pole: "POS" or "NEG"

        Returns:
            None（Matplotlibの別ウィンドウが開く）
        """
        column_name = f"{serial}_{pole}"
        elapsed_times, lsb_values, codes, datasets = self.extract_data_from_csv(
            csv_data, serial, pole, column_name
        )

        return self.create_plot_window(
            parent, serial, pole, elapsed_times, lsb_values, codes, datasets
        )
