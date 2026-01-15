import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime

class LoggerWidget:
    """ログ表示ウィジェット"""
    
    def __init__(self, parent, height=10):
        """
        初期化
        
        Parameters:
        -----------
        parent : tkinter.Frame
            親フレーム
        height : int
            テキストエリアの高さ
        """
        self.frame = tk.Frame(parent)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        # ボタンフレーム
        button_frame = tk.Frame(self.frame)
        button_frame.pack(fill=tk.X, pady=(0, 5))
        
        tk.Button(button_frame, text="ログをクリア", 
                  command=self.clear).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="全選択", 
                  command=self.select_all).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="コピー", 
                  command=self.copy_to_clipboard).pack(side=tk.LEFT, padx=5)
        
        # スクロール付きテキストエリア
        self.text_area = scrolledtext.ScrolledText(
            self.frame,
            height=height,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Courier", 9)
        )
        self.text_area.pack(fill=tk.BOTH, expand=True)
        
        # タグの設定（色分け用）
        self.text_area.tag_config("INFO", foreground="black")
        self.text_area.tag_config("SUCCESS", foreground="green")
        self.text_area.tag_config("ERROR", foreground="red")
        self.text_area.tag_config("WARNING", foreground="orange")
    
    def log(self, message, level="INFO"):
        """
        ログメッセージを追加
        
        Parameters:
        -----------
        message : str
            ログメッセージ
        level : str
            ログレベル（INFO, SUCCESS, ERROR, WARNING）
        """
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] [{level}] {message}\n"
        
        self.text_area.config(state=tk.NORMAL)
        self.text_area.insert(tk.END, log_message, level)
        self.text_area.see(tk.END)  # 最新行にスクロール
        self.text_area.config(state=tk.DISABLED)
    
    def clear(self):
        """ログをクリア"""
        self.text_area.config(state=tk.NORMAL)
        self.text_area.delete(1.0, tk.END)
        self.text_area.config(state=tk.DISABLED)
    
    def select_all(self):
        """全テキストを選択"""
        self.text_area.tag_add(tk.SEL, "1.0", tk.END)
        self.text_area.mark_set(tk.INSERT, "1.0")
        self.text_area.see(tk.INSERT)
    
    def copy_to_clipboard(self):
        """選択されたテキストをクリップボードにコピー"""
        try:
            # 選択範囲を取得
            selected_text = self.text_area.get(tk.SEL_FIRST, tk.SEL_LAST)
            # クリップボードにコピー
            self.text_area.clipboard_clear()
            self.text_area.clipboard_append(selected_text)
            self.log("クリップボードにコピーしました", "SUCCESS")
        except tk.TclError:
            # 選択範囲がない場合は全テキストをコピー
            all_text = self.text_area.get("1.0", tk.END)
            self.text_area.clipboard_clear()
            self.text_area.clipboard_append(all_text)
            self.log("全ログをクリップボードにコピーしました", "SUCCESS")