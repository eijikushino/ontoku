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

    def _get_code_color(self, code_str):
        """
        コード文字列から色を取得

        Args:
            code_str: コード文字列

        Returns:
            色（matplotlib形式）
        """
        # 括弧内のHEX値を抽出
        bracket_content = self.extract_bracket_content(code_str)
        hex_upper = bracket_content.upper()

        # コードに基づく色の割り当て
        if hex_upper in ['FFFFF', 'FFFF', '+']:
            return '#006400'  # 深緑 (darkgreen)
        elif hex_upper in ['00000', '0000', '-']:
            return '#FF0000'  # 明るい赤 (red)
        elif hex_upper in ['80000', '8000', 'C']:
            return '#00008B'  # 濃い青 (darkblue)
        else:
            # その他のManual値はグレー系
            return '#808080'  # グレー

    def _plot_by_code_with_lines(self, ax, elapsed_times, lsb_values, codes, datasets, temp_char_mode=False):
        """
        Codeごとに色分けしてプロット（連続区間のみ線で結ぶ）- 最適化版

        Args:
            ax: Matplotlibのaxisオブジェクト
            elapsed_times: 経過時間（分）のリスト
            lsb_values: LSB値のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト (Position/LBC)
            temp_char_mode: 温特グラフモード（凡例からPosition:を除去）
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

        # 色を取得（温特グラフモードは固定色、それ以外はtab10パレット）
        if temp_char_mode:
            colors = [self._get_code_color(code) for code, dataset in unique_keys]
        else:
            colors = [plt.cm.tab10(i % 10) for i in range(len(unique_keys))]

        # ラベル重複回避
        used_labels = set()

        for (code, dataset), color in zip(unique_keys, colors):
            # 凡例ラベルを作成
            if code == "---" or dataset == "---":
                legend_label = "---"
            else:
                prefix = "Position" if dataset == "Position" else "LBC"

                # Manual の場合
                if "Manual" in code:
                    bracket_content = self.extract_bracket_content(code)
                    hex_value = bracket_content.replace("Manual", "").strip().strip("()")
                    if hex_value:
                        hex_label = hex_value
                    else:
                        hex_label = "----"
                else:
                    bracket_content = self.extract_bracket_content(code)

                    try:
                        int(bracket_content, 16)
                        hex_label = bracket_content
                    except ValueError:
                        hex_values = {
                            "Position": {"+": "FFFFF", "C": "80000", "-": "00000", "H": "FFFFF"},
                            "LBC": {"+": "FFFF", "C": "8000", "-": "0000", "H": "FFFF"}
                        }
                        hex_label = hex_values.get(dataset, {}).get(bracket_content, bracket_content)

                # 温特グラフモードではPosition:を省略
                if temp_char_mode:
                    legend_label = hex_label
                else:
                    legend_label = f"{prefix}:{hex_label}"

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

    def extract_data_for_temp_characteristic(self, csv_data, serial, pole, column_name, no_abs=False):
        """
        温特グラフ用データ抽出（実測値からLSB電圧を計算）

        最初に登場するFFFFF（+Full）と00000（-Full）の平均電圧からLSB電圧を計算する。

        Args:
            csv_data: CSVデータのリスト
            serial: シリアルNo.
            pole: "POS" or "NEG"
            column_name: 列名
            no_abs: LSB電圧計算で絶対値を使用しない（デフォルト: False）

        Returns:
            (elapsed_times, lsb_values, codes, datasets, calc_info): タプル
                calc_info: 計算情報の辞書 {
                    'fffff_avg': +Full平均電圧,
                    '00000_avg': -Full平均電圧,
                    'lsb_voltage': 計算されたLSB電圧
                }
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
                    timestamp_str = row.get('Timestamp', '')
                    if not timestamp_str:
                        continue

                    timestamp = self.parse_timestamp(timestamp_str)
                    if base_timestamp is None:
                        base_timestamp = timestamp

                    elapsed_min = self.calculate_elapsed_time(timestamp, base_timestamp)
                    code_str = row.get('Code', '')
                    dataset = row.get('DataSet', '')
                    voltage = float(row[column_name])

                    elapsed_times.append(elapsed_min)
                    voltages.append(voltage)
                    codes.append(code_str)
                    datasets.append(dataset)

                except (ValueError, KeyError):
                    continue

        # スキップ処理
        if (self.skip_after_change > 0 or self.skip_first_data or self.skip_before_change) and len(codes) > 0:
            skip_indices = self._get_skip_indices(codes, datasets)
            elapsed_times = [elapsed_times[i] for i in range(len(elapsed_times)) if i not in skip_indices]
            voltages = [voltages[i] for i in range(len(voltages)) if i not in skip_indices]
            codes = [codes[i] for i in range(len(codes)) if i not in skip_indices]
            datasets = [datasets[i] for i in range(len(datasets)) if i not in skip_indices]

        # FFFFFと00000の最初の区間の平均電圧を計算
        # コード形式: "FFFFF", "+Full (FFFFF)", etc.
        fffff_voltages = []
        zero_voltages = []
        fffff_first_section_found = False
        zero_first_section_found = False
        prev_is_fffff = False
        prev_is_zero = False

        for i, code_str in enumerate(codes):
            code_upper = code_str.upper().strip()
            # FFFFFを含むかチェック（"FFFFF" or "+FULL (FFFFF)" etc.）
            is_fffff = 'FFFFF' in code_upper
            # 00000を含むかチェック（"00000" or "-FULL (00000)" etc.）
            is_zero = '00000' in code_upper

            # FFFFFの最初の区間
            if is_fffff:
                if not fffff_first_section_found:
                    if not prev_is_fffff:
                        # 新しい区間開始
                        if fffff_voltages:
                            # 既に収集済みなら最初の区間終了
                            fffff_first_section_found = True
                        else:
                            fffff_voltages.append(voltages[i])
                    else:
                        fffff_voltages.append(voltages[i])

            # 00000の最初の区間
            elif is_zero:
                if not zero_first_section_found:
                    if not prev_is_zero:
                        if zero_voltages:
                            zero_first_section_found = True
                        else:
                            zero_voltages.append(voltages[i])
                    else:
                        zero_voltages.append(voltages[i])

            prev_is_fffff = is_fffff
            prev_is_zero = is_zero

        # 平均電圧を計算
        fffff_avg = sum(fffff_voltages) / len(fffff_voltages) if fffff_voltages else None
        zero_avg = sum(zero_voltages) / len(zero_voltages) if zero_voltages else None

        # LSB電圧を計算
        if fffff_avg is not None and zero_avg is not None:
            if no_abs:
                # 絶対値を使用しない（NEGの場合、符号が負になる可能性あり）
                measured_lsb_voltage = (fffff_avg - zero_avg) / (2 ** self.bit_precision - 1)
            else:
                measured_lsb_voltage = abs(fffff_avg - zero_avg) / (2 ** self.bit_precision - 1)
        else:
            # 計算できない場合は理論値を使用
            measured_lsb_voltage = self.lsb_voltage

        calc_info = {
            'fffff_avg': fffff_avg,
            '00000_avg': zero_avg,
            'lsb_voltage': measured_lsb_voltage,
            'has_fffff': len(fffff_voltages) > 0,
            'has_zero': len(zero_voltages) > 0,
            'no_abs': no_abs
        }

        # 第2パス: LSB値に変換（初回平均モードで計算）
        # 基準電圧は各コードの最初の区間の平均を使用
        ref_voltages = self._calculate_first_section_average(voltages, codes, datasets)

        lsb_values = []
        for i in range(len(voltages)):
            voltage = voltages[i]
            ref_voltage = ref_voltages[i]
            # 実測LSB電圧を使用
            lsb_value = (voltage - ref_voltage) / measured_lsb_voltage
            lsb_values.append(lsb_value)

        return elapsed_times, lsb_values, codes, datasets, calc_info

    def plot_temperature_characteristic(self, csv_data, temp_csv_data, serial, pole,
                                         temp_yaxis_mode="manual", temp_yaxis_min=-8, temp_yaxis_max=8,
                                         no_abs=False, xaxis_full=False,
                                         temp_zone_divs=None, show_temp_arrows=True):
        """
        温特グラフを表示（2軸: LSB変動 + 温度差）

        Args:
            csv_data: 測定CSVデータ
            temp_csv_data: 温度CSVデータ
            serial: シリアルNo.
            pole: "POS" or "NEG"
            temp_yaxis_mode: Y軸モード ("auto" or "manual")
            temp_yaxis_min: Y軸(LSB)最小値（デフォルト: -8）
            temp_yaxis_max: Y軸(LSB)最大値（デフォルト: 8）
            no_abs: LSB電圧計算で絶対値を使用しない（デフォルト: False）
            xaxis_full: X軸全表示（デフォルト: False=25div固定）
            temp_zone_divs: 温度区間Div範囲（デフォルト: [(0,6), (6,13), (13,19)]）
            show_temp_arrows: 温度区間矢印を表示するか（デフォルト: True）

        Returns:
            True: 成功, None: 失敗
        """
        column_name = f"{serial}_{pole}"

        # 測定データからLSBデータを抽出（温特グラフ用：実測LSB電圧を使用）
        elapsed_times, lsb_values, codes, datasets, calc_info = self.extract_data_for_temp_characteristic(
            csv_data, serial, pole, column_name, no_abs=no_abs
        )

        if not elapsed_times:
            return None

        # 温度データを抽出・結合
        temp_times, temp_values = self._extract_temperature_data(temp_csv_data, csv_data)

        if not temp_times:
            return None

        # 2軸グラフを作成（縦:横 = 4.5:5.5）
        fig, ax1 = plt.subplots(figsize=(7.33, 6))

        # 左軸: LSB変動（コードごとに色分け、温特グラフ用凡例）
        self._plot_by_code_with_lines(ax1, elapsed_times, lsb_values, codes, datasets, temp_char_mode=True)
        ax1.set_xlabel('時間 10分/Div')
        ax1.set_ylabel('変動値 20bit@LSB/Div', color='black')
        ax1.tick_params(axis='y', labelcolor='black')

        # X軸フォーマット（デフォルト25div=250分、または全表示）
        max_minutes = max(elapsed_times) if elapsed_times else 60
        min_minutes = min(elapsed_times) if elapsed_times else 0
        self._format_time_axis_temp_char(ax1, min_minutes, max_minutes, xaxis_full)

        # 左軸: Y軸範囲設定
        if temp_yaxis_mode == "auto":
            # オートモード: データ範囲に合わせる
            self._format_lsb_axis(ax1, lsb_values)
            y_min, y_max = ax1.get_ylim()
        else:
            # マニュアルモード: 指定値を使用
            y_min, y_max = temp_yaxis_min, temp_yaxis_max
            ax1.set_ylim(y_min, y_max)
            ax1.set_yticks(np.arange(y_min, y_max + self.lsb_per_div, self.lsb_per_div))
        ax1.ticklabel_format(style='plain', axis='y', useOffset=False)
        ax1.grid(True, axis='y', alpha=0.5, linestyle='-', linewidth=0.5)

        # 右軸: 温度（常に±8℃固定）
        ax2 = ax1.twinx()
        ax2.plot(temp_times, temp_values, color='#8B4513', linestyle='-',
                 linewidth=1.5, label='温度', alpha=0.8)
        ax2.set_ylabel('温度(℃)', color='black', rotation=270, labelpad=15)
        ax2.tick_params(axis='y', labelcolor='black')
        ax2.set_ylim(-8, 8)  # 温度軸は±8℃固定
        ax2.set_yticks(np.arange(-8, 10, 2))  # 2℃ごとに目盛り

        # タイトル（serialがDFHで始まる場合は重複を避ける）
        if serial.upper().startswith('DFH'):
            title_prefix = '1PB397MK2'
        else:
            title_prefix = '1PB397MK2DFH_'
        fig.suptitle(f'{title_prefix}{serial} {pole} 温度特性試験結果')

        # 温度区間の矢印を描画
        if show_temp_arrows:
            # デフォルトのDiv範囲
            if temp_zone_divs is None:
                temp_zone_divs = [(0, 6), (6, 13), (13, 19)]

            # Div範囲から分単位の温度区間を作成（1Div = 10分）
            temp_zone_labels = ["23℃", "28℃", "18℃"]
            temp_zones = []
            for i, (start_div, end_div) in enumerate(temp_zone_divs):
                temp_zones.append({
                    'zone_num': i,
                    'start_time': start_div * 10,  # Div → 分
                    'end_time': end_div * 10
                })

            if temp_zones:
                # Y軸の上部に矢印を配置（Y軸範囲の上端より少し上）
                y_min_val, y_max_val = ax1.get_ylim()
                arrow_y = y_max_val + (y_max_val - y_min_val) * 0.02
                self._draw_temp_zone_arrows(ax1, temp_zones, temp_zone_labels, arrow_y)

        # 凡例を統合
        lines1, labels1 = ax1.get_legend_handles_labels()
        lines2, labels2 = ax2.get_legend_handles_labels()
        ax1.legend(lines1 + lines2, labels1 + labels2, loc='upper left')

        fig.tight_layout()
        plt.show(block=False)

        # 計算情報とfigureを返す（設定画面で表示・保存用）
        calc_info['figure'] = fig
        return calc_info

    def _extract_temperature_data(self, temp_csv_data, measurement_csv_data):
        """
        温度CSVデータから経過時間と温度を抽出

        Args:
            temp_csv_data: 温度CSVデータ（測定順と温度を含む）
            measurement_csv_data: 測定CSVデータ（タイムスタンプ取得用）

        Returns:
            (elapsed_times, temp_values): 経過時間（分）と温度のリスト
        """
        elapsed_times = []
        temp_values = []

        # 温度CSVのカラム名を自動検出
        temp_column = None
        index_column = None
        if temp_csv_data:
            first_row = temp_csv_data[0]
            for key in first_row.keys():
                # BOM付きの場合も対応
                clean_key = key.replace('\ufeff', '').strip()
                if 'temp' in clean_key.lower() or '温度' in clean_key:
                    temp_column = key
                if '測定順' in clean_key or 'index' in clean_key.lower() or 'no' in clean_key.lower():
                    index_column = key

        # カラム名が見つからない場合、最初の2列を使用
        if temp_column is None or index_column is None:
            if temp_csv_data:
                keys = list(temp_csv_data[0].keys())
                if len(keys) >= 2:
                    index_column = keys[0]
                    temp_column = keys[1]

        if temp_column is None:
            return [], []

        # 測定CSVから経過時間を計算（タイムスタンプベース）
        measurement_times = []
        base_timestamp = None
        for row in measurement_csv_data:
            timestamp_str = row.get('Timestamp', '')
            if timestamp_str:
                try:
                    timestamp = self.parse_timestamp(timestamp_str)
                    if base_timestamp is None:
                        base_timestamp = timestamp
                    elapsed_min = self.calculate_elapsed_time(timestamp, base_timestamp)
                    measurement_times.append(elapsed_min)
                except:
                    pass

        # 温度データを測定順に対応させる
        total_temp_points = len(temp_csv_data)

        if total_temp_points == 0:
            return [], []

        # 温度CSVの1列目がタイムスタンプか番号かを判定
        is_timestamp = self._is_timestamp_column(temp_csv_data, index_column)

        # 温度データを辞書に格納（測定順/タイムスタンプ → 温度）
        temp_dict = {}
        for row in temp_csv_data:
            try:
                idx_str = row.get(index_column, '').strip()
                temp_str = row.get(temp_column, '').strip()
                if idx_str and temp_str:
                    idx = int(idx_str)
                    temp_value = float(temp_str)
                    temp_dict[idx] = temp_value
            except (ValueError, KeyError):
                continue

        if is_timestamp:
            # タイムスタンプの場合、測定CSVの経過時間を使用
            total_measurement_points = len(measurement_times)
            if total_measurement_points > 0:
                # 測定点数と温度点数が同じ場合、1:1で対応
                if total_measurement_points == total_temp_points:
                    for i, elapsed_min in enumerate(measurement_times):
                        idx = i + 1  # 測定順は1から始まる
                        if idx in temp_dict:
                            elapsed_times.append(elapsed_min)
                            temp_values.append(temp_dict[idx])
                else:
                    # 点数が異なる場合も測定順で対応（存在する分だけ）
                    for i, elapsed_min in enumerate(measurement_times):
                        idx = i + 1
                        if idx in temp_dict:
                            elapsed_times.append(elapsed_min)
                            temp_values.append(temp_dict[idx])
        else:
            # 番号の場合、10分あたり100点 = 1点あたり0.1分として計算
            for idx in sorted(temp_dict.keys()):
                elapsed_min = (idx - 1) * 0.1  # 番号1が0分からスタート
                elapsed_times.append(elapsed_min)
                temp_values.append(temp_dict[idx])

        return elapsed_times, temp_values

    def _is_timestamp_column(self, temp_csv_data, index_column):
        """
        温度CSVの1列目がタイムスタンプかどうかを判定

        Args:
            temp_csv_data: 温度CSVデータ
            index_column: 1列目のカラム名

        Returns:
            bool: タイムスタンプならTrue、番号ならFalse
        """
        if not temp_csv_data or not index_column:
            return False

        # 最初の数行の値をチェック
        sample_values = []
        for i, row in enumerate(temp_csv_data[:10]):
            try:
                val_str = row.get(index_column, '').strip()
                if val_str:
                    sample_values.append(val_str)
            except:
                continue

        if not sample_values:
            return False

        # 判定1: 日時フォーマットを含むか（:や-が含まれていればタイムスタンプ）
        for val in sample_values:
            if ':' in val or '-' in val:
                return True

        # 判定2: 連続した小さい整数（1, 2, 3...）なら番号
        try:
            int_values = [int(v) for v in sample_values]
            # 最初の値が1〜10で、連続しているなら番号と判断
            if int_values[0] <= 10:
                is_sequential = all(int_values[i] == int_values[0] + i for i in range(len(int_values)))
                if is_sequential:
                    return False
        except ValueError:
            pass

        # デフォルトはタイムスタンプとして扱う
        return True

    def _format_time_axis_10min(self, ax, max_minutes):
        """
        X軸を10分/div、25divでフォーマット

        Args:
            ax: Matplotlibのaxisオブジェクト
            max_minutes: 最大経過時間（分）- 未使用、250分固定
        """
        tick_interval = 10  # 10分/div
        total_divs = 25     # 25div固定
        max_time = tick_interval * total_divs  # 250分

        # 目盛り位置を計算（0, 10, 20, ... 250）
        ticks = np.arange(0, max_time + tick_interval, tick_interval)

        # 目盛りラベルを「hour min」形式に変換
        labels = [self.format_time_label(t) for t in ticks]

        ax.set_xticks(ticks)
        ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
        ax.set_xlim(0, max_time)

    def _format_time_axis_auto(self, ax, min_minutes, max_minutes):
        """
        X軸をデータ範囲に合わせてフォーマット（余白なし）

        Args:
            ax: Matplotlibのaxisオブジェクト
            min_minutes: 最小経過時間（分）
            max_minutes: 最大経過時間（分）
        """
        data_range = max_minutes - min_minutes

        # 適切な目盛り間隔を決定
        if data_range < 30:
            tick_interval = 5
        elif data_range < 60:
            tick_interval = 10
        elif data_range < 180:
            tick_interval = 30
        elif data_range < 600:
            tick_interval = 60
        else:
            tick_interval = 120

        # 目盛り位置を計算（データ範囲に合わせる）
        start_tick = int(min_minutes / tick_interval) * tick_interval
        end_tick = int(np.ceil(max_minutes / tick_interval)) * tick_interval
        ticks = np.arange(start_tick, end_tick + tick_interval, tick_interval)

        # 目盛りラベルを「hour min」形式に変換
        labels = [self.format_time_label(t) for t in ticks]

        ax.set_xticks(ticks)
        ax.set_xticklabels(labels, rotation=45, ha='right', fontsize=8)
        # データ範囲に合わせてX軸を設定（余白なし）
        ax.set_xlim(min_minutes, max_minutes)

    def _format_time_axis_temp_char(self, ax, min_minutes, max_minutes, xaxis_full=False):
        """
        温特グラフ用X軸フォーマット（10分ごとの補助線、目盛りラベルなし）

        Args:
            ax: Matplotlibのaxisオブジェクト
            min_minutes: 最小経過時間（分）
            max_minutes: 最大経過時間（分）
            xaxis_full: 全表示モード（False=25div固定、True=データ範囲）
        """
        tick_interval = 10  # 10分/Div
        default_max = 250  # デフォルト25div = 250分

        if xaxis_full:
            # 全表示モード: データ範囲に合わせる
            display_max = max_minutes
        else:
            # デフォルト: 25div=250分固定
            display_max = default_max

        # 目盛り位置を計算（10分ごと）
        start_tick = 0
        end_tick = int(np.ceil(display_max / tick_interval)) * tick_interval
        ticks = np.arange(start_tick, end_tick + tick_interval, tick_interval)

        ax.set_xticks(ticks)
        ax.set_xticklabels(['' for _ in ticks])  # 目盛りラベルなし
        # X軸範囲を設定（0から開始）
        ax.set_xlim(0, display_max)

        # 10分ごとに補助線（グリッド）を追加
        ax.grid(True, axis='x', alpha=0.5, linestyle='-', linewidth=0.5)
        ax.grid(True, axis='y', alpha=0.3, linestyle='-', linewidth=0.5)

    def extract_section_averages(self, csv_data, serial, pole, column_name, last_minutes=10):
        """
        各区間の平均電圧を抽出（最後のN分から計算）

        温特パターン構成:
        1. 23℃: FFFFF → 00000 → 80000
        2. 28℃: FFFFF → 00000 → 80000
        3. 18℃: FFFFF → 00000 → 80000
        4. 23℃: FFFFF → 00000 → 80000

        Args:
            csv_data: CSVデータのリスト
            serial: シリアルNo.
            pole: "POS" or "NEG"
            column_name: 列名
            last_minutes: 平均計算に使用する最後のN分（デフォルト: 10分）

        Returns:
            list: 各区間の情報リスト
                [{
                    'section_num': 区間番号（1始まり）,
                    'code': コード名（FFFFF/00000/80000）,
                    'avg_voltage': 平均電圧,
                    'data_count': データ数,
                    'total_minutes': 区間の総時間（分）,
                    'used_minutes': 計算に使用した時間（分）
                }, ...]
        """
        # 生データを収集
        raw_data = []
        base_timestamp = None

        for row in csv_data:
            if column_name in row and row[column_name]:
                try:
                    timestamp_str = row.get('Timestamp', '')
                    if not timestamp_str:
                        continue

                    timestamp = self.parse_timestamp(timestamp_str)
                    if base_timestamp is None:
                        base_timestamp = timestamp

                    elapsed_min = self.calculate_elapsed_time(timestamp, base_timestamp)
                    code_str = row.get('Code', '')
                    dataset = row.get('DataSet', '')
                    voltage = float(row[column_name])

                    raw_data.append({
                        'elapsed_min': elapsed_min,
                        'voltage': voltage,
                        'code': code_str,
                        'dataset': dataset
                    })

                except (ValueError, KeyError):
                    continue

        if not raw_data:
            return []

        # スキップ処理
        codes = [d['code'] for d in raw_data]
        datasets = [d['dataset'] for d in raw_data]
        if (self.skip_after_change > 0 or self.skip_first_data or self.skip_before_change):
            skip_indices = self._get_skip_indices(codes, datasets)
            raw_data = [raw_data[i] for i in range(len(raw_data)) if i not in skip_indices]

        if not raw_data:
            return []

        # 区間を検出（連続する同じコードを1区間とする）
        sections = []
        current_section = None
        current_code = None

        for data in raw_data:
            code_upper = data['code'].upper().strip()
            # コード名を正規化
            if 'FFFFF' in code_upper:
                normalized_code = 'FFFFF'
            elif '00000' in code_upper:
                normalized_code = '00000'
            elif '80000' in code_upper:
                normalized_code = '80000'
            else:
                normalized_code = code_upper

            if normalized_code != current_code:
                # 新しい区間開始
                if current_section is not None:
                    sections.append(current_section)
                current_section = {
                    'code': normalized_code,
                    'data': [data]
                }
                current_code = normalized_code
            else:
                current_section['data'].append(data)

        # 最後の区間を追加
        if current_section is not None:
            sections.append(current_section)

        # 各区間の平均を計算
        # 温特パターン: FFFFF→00000→80000 を5回繰り返し（23℃, 28℃, 18℃, 23℃, 23℃戻し）
        # 28℃〜23℃(2)のFFFFFのみ最後10分を使用、それ以外は全データを使用
        # 23℃戻しは参考値なので全データを使用
        codes_per_temp = 3  # FFFFF, 00000, 80000
        first_temp_sections = codes_per_temp  # 最初の23℃は3区間（0,1,2）
        last_temp_start = codes_per_temp * 4  # 23℃戻しの開始区間（12）

        results = []
        for i, section in enumerate(sections):
            data_list = section['data']
            if not data_list:
                continue

            # 区間の時間範囲
            start_time = data_list[0]['elapsed_min']
            end_time = data_list[-1]['elapsed_min']
            total_minutes = end_time - start_time

            # 28℃〜23℃(2)のFFFFFかどうかを判定
            # 区間番号: 0,1,2=23℃(1), 3,4,5=28℃, 6,7,8=18℃, 9,10,11=23℃(2), 12,13,14=23℃戻し
            # FFFFFは各温度の最初（0, 3, 6, 9, 12）
            is_fffff = section['code'] == 'FFFFF'
            is_after_first_temp = i >= first_temp_sections  # 28℃以降
            is_last_temp = i >= last_temp_start  # 23℃戻し

            if is_fffff and is_after_first_temp and not is_last_temp:
                # 28℃〜23℃(2)のFFFFF: 最後10分のデータを使用
                cutoff_time = end_time - last_minutes
                calc_data = [d for d in data_list if d['elapsed_min'] >= cutoff_time]
                if not calc_data:
                    calc_data = data_list
                use_last_10min = True
            else:
                # それ以外（23℃(1)、00000/80000、23℃戻し）: 全データを使用
                calc_data = data_list
                use_last_10min = False

            # 平均電圧を計算
            voltages = [d['voltage'] for d in calc_data]
            avg_voltage = sum(voltages) / len(voltages)

            # 使用した時間範囲
            used_start = calc_data[0]['elapsed_min']
            used_end = calc_data[-1]['elapsed_min']
            used_minutes = used_end - used_start

            results.append({
                'section_num': i + 1,
                'code': section['code'],
                'avg_voltage': avg_voltage,
                'data_count': len(calc_data),
                'total_data_count': len(data_list),
                'total_minutes': total_minutes,
                'used_minutes': used_minutes,
                'start_time': start_time,
                'end_time': end_time,
                'use_last_10min': use_last_10min
            })

        return results

    def _detect_temp_zones(self, elapsed_times, codes, datasets):
        """
        温度区間を検出（コードパターンから判定）

        温特パターン: FFFFF→00000→80000 を4回繰り返し
        - 1回目: 23℃
        - 2回目: 28℃
        - 3回目: 18℃
        - 4回目: 23℃（戻し）

        Args:
            elapsed_times: 経過時間（分）のリスト
            codes: Codeのリスト
            datasets: DataSetのリスト

        Returns:
            list: 温度区間リスト
                [{
                    'zone_num': 温度区間番号（0始まり）,
                    'start_time': 開始時間（分）,
                    'end_time': 終了時間（分）
                }, ...]
        """
        if not elapsed_times or not codes:
            return []

        # コード区間を検出（連続する同じコードを1区間）
        code_sections = []
        current_code = None
        section_start_idx = 0

        for i, code_str in enumerate(codes):
            # コード名を正規化
            code_upper = code_str.upper().strip()
            if 'FFFFF' in code_upper:
                normalized = 'FFFFF'
            elif '00000' in code_upper:
                normalized = '00000'
            elif '80000' in code_upper:
                normalized = '80000'
            else:
                normalized = code_upper

            if normalized != current_code:
                if current_code is not None:
                    code_sections.append({
                        'code': current_code,
                        'start_idx': section_start_idx,
                        'end_idx': i - 1
                    })
                current_code = normalized
                section_start_idx = i

        # 最後の区間を追加
        if current_code is not None:
            code_sections.append({
                'code': current_code,
                'start_idx': section_start_idx,
                'end_idx': len(codes) - 1
            })

        # 温度区間を構成（3コード=1温度区間）
        temp_zones = []
        codes_per_temp = 3  # FFFFF, 00000, 80000

        for i in range(0, len(code_sections), codes_per_temp):
            if i + codes_per_temp - 1 < len(code_sections):
                # 温度区間の開始・終了インデックス
                start_idx = code_sections[i]['start_idx']
                end_idx = code_sections[i + codes_per_temp - 1]['end_idx']

                temp_zones.append({
                    'zone_num': len(temp_zones),
                    'start_time': elapsed_times[start_idx],
                    'end_time': elapsed_times[end_idx]
                })

        return temp_zones

    def _draw_temp_zone_arrows(self, ax, temp_zones, temp_labels, y_position):
        """
        温度区間の矢印とラベルを描画

        Args:
            ax: Matplotlibのaxisオブジェクト
            temp_zones: 温度区間リスト
            temp_labels: 温度ラベルリスト（例: ["23℃", "28℃", "18℃"]）
            y_position: Y軸位置（データ座標）
        """
        if not temp_zones or not temp_labels:
            return

        # 最大3つの温度区間のみ描画（4つ目の戻しは除外）
        num_zones = min(len(temp_zones), len(temp_labels), 3)

        for i in range(num_zones):
            zone = temp_zones[i]
            label = temp_labels[i] if i < len(temp_labels) else ""

            start_time = zone['start_time']
            end_time = zone['end_time']
            center_time = (start_time + end_time) / 2

            # 両端矢印を描画（←→）
            ax.annotate('', xy=(end_time, y_position), xytext=(start_time, y_position),
                       arrowprops=dict(arrowstyle='<->', color='#555555', lw=1.2),
                       annotation_clip=False)

            # 温度ラベルを中央に配置
            ax.text(center_time, y_position, label,
                   ha='center', va='bottom', fontsize=9, color='#333333',
                   clip_on=False)
