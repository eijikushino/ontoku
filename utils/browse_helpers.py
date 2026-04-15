"""ファイル/フォルダ選択ダイアログの共通ヘルパー

既設パスを initialdir / initialfile にセットして、
前回の選択を即座に再利用できるようにする。
"""
import os
from tkinter import filedialog


def _find_existing_dir(path: str) -> str:
    """指定パスが存在しない場合、存在する親フォルダまで遡る"""
    if not path:
        return ""
    abs_path = os.path.abspath(path)
    probe = abs_path
    while probe and not os.path.isdir(probe):
        parent = os.path.dirname(probe)
        if parent == probe:
            return ""
        probe = parent
    return probe


def pick_directory(current_path: str, title: str = "フォルダを選択") -> str:
    """既設パスを初期表示にしてフォルダ選択ダイアログを開く

    Args:
        current_path: 現在設定されているフォルダパス (空でも可)
        title: ダイアログタイトル

    Returns:
        選択されたフォルダパス (キャンセル時は空文字列)
    """
    initial = _find_existing_dir(current_path)
    return filedialog.askdirectory(title=title, initialdir=initial or None)


def pick_file(current_path: str,
              title: str = "ファイルを選択",
              filetypes=None) -> str:
    """既設ファイルパスを初期表示にしてファイル選択ダイアログを開く

    既設パスがファイルとして存在する場合は、
    そのファイル名まで initialfile にセットしてハイライト状態で開く。
    ファイルが存在しない場合でも、親フォルダがあればそこを開く。

    Args:
        current_path: 現在設定されているファイルパス (空でも可)
        title: ダイアログタイトル
        filetypes: filedialog 用のファイル種別リスト

    Returns:
        選択されたファイルパス (キャンセル時は空文字列)
    """
    initial_dir = ""
    initial_file = ""

    if current_path:
        abs_path = os.path.abspath(current_path)
        if os.path.isfile(abs_path):
            initial_dir = os.path.dirname(abs_path)
            initial_file = os.path.basename(abs_path)
        else:
            parent = os.path.dirname(abs_path)
            initial_dir = _find_existing_dir(parent)

    kwargs = {
        "title": title,
        "filetypes": filetypes or [("すべて", "*.*")],
    }
    if initial_dir:
        kwargs["initialdir"] = initial_dir
    if initial_file:
        kwargs["initialfile"] = initial_file

    return filedialog.askopenfilename(**kwargs)
