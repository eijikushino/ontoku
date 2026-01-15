import csv
import os
import threading
from datetime import datetime


class MeasurementCSVLogger:
    """
    パターン計測用CSV保存クラス
    
    スキャナ周回ごとにタイムスタンプと測定電圧を記録
    周回完了ごとにファイルに追記（メモリハング対策）
    シリアルNo.が異なる場合は新しい列に追加
    """
    
    def __init__(self, save_dir, filename, def_serial_numbers):
        """
        Args:
            save_dir: 保存先ディレクトリ
            filename: CSVファイル名（拡張子付き、例: "measurement.csv"）
            def_serial_numbers: {def_index: serial_number} の辞書
                               例: {0: "DFH903", 1: "SUBJ02", ...}
        """
        self.save_dir = save_dir
        self.filename = filename
        self.def_serial_numbers = def_serial_numbers  # {def_index: "DFH903", ...}
        
        # 保存用データ構造
        self.cycle_data = {}  # 現在の1周回分のデータ {def_index: {"POS": value, "NEG": value}}
        self.cycle_timestamp = None
        self.cycle_dataset = None    # ★★★ 追加 ★★★
        self.cycle_code = None        # ★★★ 追加 ★★★


        self.is_logging = False
        self.csv_filepath = None
        self.csv_file = None
        self.csv_writer = None
        self.existing_headers = []  # 既存のヘッダー情報
        self.current_headers = []   # 現在使用するヘッダー

        # スレッド化用ロック（CSV書き込みの排他制御）
        self.write_lock = threading.Lock()
        
    def start_logging(self):
        """ログ記録開始"""
        if self.is_logging:
            return False, "既にログ記録中です"
        
        # 保存先ディレクトリの作成
        if not os.path.exists(self.save_dir):
            try:
                os.makedirs(self.save_dir)
            except Exception as e:
                return False, f"ディレクトリ作成失敗: {str(e)}"
        
        # ファイルパス
        self.csv_filepath = os.path.join(self.save_dir, self.filename)
        
        try:
            # ファイルが既に存在するか確認
            file_exists = os.path.exists(self.csv_filepath)
            
            if file_exists:
                # 既存ファイルのヘッダーを読み込む
                with open(self.csv_filepath, 'r', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    self.existing_headers = next(reader, [])
                
                # 現在のヘッダーを生成
                self.current_headers = self._generate_headers()
                
                # ヘッダーが異なる場合、既存ファイルを更新
                if self.existing_headers != self.current_headers:
                    self._update_file_with_new_columns()
                
                # 追記モードで開く
                self.csv_file = open(self.csv_filepath, 'a', newline='', encoding='utf-8')
                self.csv_writer = csv.writer(self.csv_file)
                
            else:
                # 新規ファイル作成
                self.current_headers = self._generate_headers()
                self.csv_file = open(self.csv_filepath, 'w', newline='', encoding='utf-8')
                self.csv_writer = csv.writer(self.csv_file)
                
                # ヘッダー行を書き込み
                self.csv_writer.writerow(self.current_headers)
                self.csv_file.flush()
            
        except Exception as e:
            if self.csv_file:
                self.csv_file.close()
            return False, f"ファイル操作失敗: {str(e)}"
        
        self.is_logging = True
        self.cycle_data = {}
        self.cycle_timestamp = None
        self.cycle_dataset = None
        self.cycle_code = None

        mode_str = "追記" if file_exists else "新規作成"
        return True, f"ログ記録開始: {self.filename} ({mode_str})"
    
    def stop_logging(self):
        """ログ記録停止"""
        if not self.is_logging:
            return False, "ログ記録が開始されていません"

        self.is_logging = False

        # 最後の1周回分のデータを保存（未完了でも保存）
        if self.cycle_timestamp and self.cycle_data:
            self._write_cycle()
        
        # ファイルを閉じる
        try:
            if self.csv_file:
                self.csv_file.close()
                self.csv_file = None
                self.csv_writer = None
            
            return True, f"保存完了: {self.csv_filepath}"
        
        except Exception as e:
            return False, f"保存失敗: {str(e)}"
    
    def record_measurement(self, def_index, pole, value, is_cycle_start=False, dataset='', code=''):
        """
        測定値を記録（シンプル版：即時書き込み）

        Args:
            def_index: DEFインデックス (0-5)
            pole: "POS" or "NEG"
            value: 測定値（文字列）
            is_cycle_start: スキャナ周回の最初の測定かどうか
            dataset: DataSet値
            code: Code値
        """
        if not self.is_logging:
            return

        # 周回の最初の測定の場合
        if is_cycle_start:
            # 前の周回データがあれば書き込み
            if self.cycle_timestamp is not None and self.cycle_data:
                thread = threading.Thread(
                    target=self._write_cycle_async,
                    args=({
                        'timestamp': self.cycle_timestamp,
                        'dataset': self.cycle_dataset,
                        'code': self.cycle_code,
                        'data': dict(self.cycle_data)
                    },),
                    daemon=True
                )
                thread.start()

            # 新しい周回を開始
            self.cycle_timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
            self.cycle_dataset = dataset
            self.cycle_code = code
            self.cycle_data = {}

        # 測定値を記録
        if def_index not in self.cycle_data:
            self.cycle_data[def_index] = {}

        self.cycle_data[def_index][pole] = value

    def discard_current_cycle(self):
        """現在の周回データを破棄（コード変更時に呼び出し）"""
        self.cycle_data = {}
        self.cycle_timestamp = None
        self.cycle_dataset = None
        self.cycle_code = None

    def _generate_headers(self):
        """現在のシリアルNo.からヘッダーを生成"""
        headers = ["Timestamp", "DataSet", "Code"]  # ★★★ DataSet, Codeを追加 ★★★
        for def_idx in sorted(self.def_serial_numbers.keys()):
            serial = self.def_serial_numbers[def_idx]
            headers.append(f"{serial}_POS")
            headers.append(f"{serial}_NEG")
        return headers
    
    def _update_file_with_new_columns(self):
        """既存ファイルに新しい列を追加（既存列も保持）"""
        try:
            # 既存データを全て読み込む
            existing_data = []
            with open(self.csv_filepath, 'r', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    existing_data.append(row)
            
            # ★★★ 既存ヘッダーと新しいヘッダーをマージ ★★★
            # Timestamp, DataSet, Codeは最初に固定
            merged_headers = ["Timestamp", "DataSet", "Code"]  # ★★★ 修正 ★★★
            
            # 既存のヘッダー（固定列以外）を全て追加
            for header in self.existing_headers:
                if header not in ["Timestamp", "DataSet", "Code"] and header not in merged_headers:  # ★★★ 修正 ★★★
                    merged_headers.append(header)
            
            # 新しいヘッダー（固定列以外）で既存にないものを追加
            for header in self.current_headers:
                if header not in ["Timestamp", "DataSet", "Code"] and header not in merged_headers:  # ★★★ 修正 ★★★
                    merged_headers.append(header)
            
            # マージしたヘッダーを使用
            self.current_headers = merged_headers
            
            # 新しいヘッダーでファイルを書き換え
            with open(self.csv_filepath, 'w', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=self.current_headers)
                writer.writeheader()
                
                # 既存データを新しい列構造で書き込み
                for row in existing_data:
                    new_row = {}
                    for header in self.current_headers:
                        # 既存の列データがあればそれを使用、なければ空白
                        new_row[header] = row.get(header, '')
                    writer.writerow(new_row)
            
        except Exception as e:
            raise Exception(f"列追加処理失敗: {str(e)}")
    
    def _write_cycle_async(self, write_data):
        """1周回分のデータをCSVファイルに書き込み（別スレッドで実行）"""
        with self.write_lock:
            if not self.csv_writer:
                return

            try:
                # 1行分のデータを作成
                row = [write_data['timestamp'], write_data['dataset'] or '', write_data['code'] or '']

                # ヘッダーに従ってデータを配置
                for header in self.current_headers[3:]:
                    if '_POS' in header:
                        serial = header.replace('_POS', '')
                        pole = 'POS'
                    elif '_NEG' in header:
                        serial = header.replace('_NEG', '')
                        pole = 'NEG'
                    else:
                        row.append('')
                        continue

                    # 該当するdef_indexを検索
                    value = ''
                    for def_idx, sn in self.def_serial_numbers.items():
                        if sn == serial:
                            data = write_data['data'].get(def_idx, {})
                            value = data.get(pole, '')
                            break

                    row.append(value)

                # CSVファイルに書き込み
                self.csv_writer.writerow(row)
                self.csv_file.flush()
            except Exception:
                pass  # 別スレッドなのでエラーは無視

    def _write_cycle(self):
        """1周回分のデータをCSVファイルに書き込み（同期版 - stop_logging用）"""
        if self.cycle_timestamp is None or not self.csv_writer:
            return

        try:
            # 1行分のデータを作成
            row = [self.cycle_timestamp, self.cycle_dataset or '', self.cycle_code or '']  # ★★★ DataSet, Codeを追加 ★★★

            # ヘッダーに従ってデータを配置
            for header in self.current_headers[3:]:  # ★★★ Timestamp, DataSet, Codeを除く（[1:]→[3:]に変更）★★★
                # ヘッダーから SerialNo_POS/NEG を解析
                if '_POS' in header:
                    serial = header.replace('_POS', '')
                    pole = 'POS'
                elif '_NEG' in header:
                    serial = header.replace('_NEG', '')
                    pole = 'NEG'
                else:
                    row.append('')
                    continue

                # 該当するdef_indexを検索
                value = ''
                for def_idx, sn in self.def_serial_numbers.items():
                    if sn == serial:
                        data = self.cycle_data.get(def_idx, {})
                        value = data.get(pole, '')
                        break

                row.append(value)

            # CSVファイルに書き込み
            self.csv_writer.writerow(row)
            self.csv_file.flush()  # すぐにディスクに書き込む
            
        except Exception as e:
            print(f"CSV書き込みエラー: {e}")
        
        # 次の周回のためにクリア
        self.cycle_data = {}
        self.cycle_timestamp = None
        self.cycle_dataset = None   # ★★★ 追加 ★★★
        self.cycle_code = None       # ★★★ 追加 ★★★

