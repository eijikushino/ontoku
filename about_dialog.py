# about_dialog.py
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from version import __version__, __build_date__
import os
import sys

def show_about_dialog(parent):
    """バージョン情報ダイアログを表示"""
    dialog = tk.Toplevel(parent)
    dialog.title("バージョン情報")
    dialog.geometry("600x500")
    dialog.resizable(True, True)
    dialog.grab_set()

    # 中央寄せ
    dialog.transient(parent)
    x = parent.winfo_x() + (parent.winfo_width() - 600) // 2
    y = parent.winfo_y() + (parent.winfo_height() - 500) // 2
    dialog.geometry(f"600x500+{x}+{y}")

    # 上部：基本情報
    top_frame = ttk.Frame(dialog, padding=15)
    top_frame.pack(fill="x")

    ttk.Label(top_frame, text="DEF Command Set App",
              font=("Arial", 16, "bold")).pack(pady=(0, 8))

    info_frame = ttk.Frame(top_frame)
    info_frame.pack()

    ttk.Label(info_frame, text=f"バージョン: {__version__}",
              font=("Arial", 11)).grid(row=0, column=0, sticky="w", padx=5, pady=2)
    ttk.Label(info_frame, text=f"ビルド日: {__build_date__}",
              font=("Arial", 11)).grid(row=0, column=1, sticky="w", padx=5, pady=2)

    ttk.Separator(dialog, orient="horizontal").pack(fill="x", padx=15, pady=10)

    # 中部：変更履歴
    changelog_frame = ttk.LabelFrame(dialog, text="変更履歴", padding=10)
    changelog_frame.pack(fill="both", expand=True, padx=15, pady=(0, 10))

    # スクロール可能なテキスト
    text_frame = ttk.Frame(changelog_frame)
    text_frame.pack(fill="both", expand=True)

    changelog_text = ScrolledText(text_frame, wrap="word",
                                  font=("Consolas", 9), height=15)
    changelog_text.pack(fill="both", expand=True)

    # CHANGELOG.mdを読み込んで表示
    changelog_content = load_changelog()
    if changelog_content:
        changelog_text.insert("1.0", changelog_content)
    else:
        changelog_text.insert("1.0", "変更履歴ファイルが見つかりません。")

    changelog_text.config(state="disabled")

    # 下部：閉じるボタン
    bottom_frame = ttk.Frame(dialog, padding=10)
    bottom_frame.pack(fill="x")

    ttk.Label(bottom_frame, text="© 2025 All Rights Reserved",
              font=("Arial", 9)).pack(side="left")

    ttk.Button(bottom_frame, text="閉じる", command=dialog.destroy,
               width=10).pack(side="right")


def load_changelog():
    """CHANGELOG.mdを読み込む"""
    # 実行環境に応じてパスを調整
    changelog_paths = [
        "CHANGELOG.md",  # 通常のPython実行時
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "CHANGELOG.md"),
    ]

    # PyInstallerでexe化されている場合
    if getattr(sys, 'frozen', False):
        # exe実行時の一時フォルダ
        base_path = sys._MEIPASS
        changelog_paths.insert(0, os.path.join(base_path, "CHANGELOG.md"))
        # exeと同じディレクトリ
        exe_dir = os.path.dirname(sys.executable)
        changelog_paths.insert(0, os.path.join(exe_dir, "CHANGELOG.md"))

    # 各パスを順番に試す
    for path in changelog_paths:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return f.read()
            except Exception as e:
                print(f"CHANGELOG.md読み込みエラー: {e}")
                continue

    return None


# テスト用
if __name__ == "__main__":
    root = tk.Tk()
    root.geometry("400x300")

    ttk.Button(root, text="バージョン情報を表示",
               command=lambda: show_about_dialog(root)).pack(pady=50)

    root.mainloop()
