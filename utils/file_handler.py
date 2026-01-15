import csv
import json
from datetime import datetime
from tkinter import filedialog

class FileHandler:
    """ファイル操作用クラス"""
    
    @staticmethod
    def save_csv(data, headers=None, default_name=None):
        """CSVファイルに保存"""
        if default_name is None:
            default_name = f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not filename:
            return False, "保存がキャンセルされました"
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                if headers:
                    writer.writerow(headers)
                writer.writerows(data)
            return True, f"保存成功: {filename}"
        except Exception as e:
            return False, f"保存失敗: {str(e)}"
    
    @staticmethod
    def load_csv(default_dir=None):
        """CSVファイルを読み込み"""
        filename = filedialog.askopenfilename(
            initialdir=default_dir,
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if not filename:
            return False, "読み込みがキャンセルされました", None
        
        try:
            data = []
            with open(filename, 'r', encoding='utf-8') as f:
                reader = csv.reader(f)
                for row in reader:
                    data.append(row)
            return True, f"読み込み成功: {filename}", data
        except Exception as e:
            return False, f"読み込み失敗: {str(e)}", None
    
    @staticmethod
    def save_json(data, default_name=None):
        """JSONファイルに保存"""
        if default_name is None:
            default_name = f"data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=default_name,
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
        )
        
        if not filename:
            return False, "保存がキャンセルされました"
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            return True, f"保存成功: {filename}"
        except Exception as e:
            return False, f"保存失敗: {str(e)}"