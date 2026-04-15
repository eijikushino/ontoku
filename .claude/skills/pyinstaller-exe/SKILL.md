---
name: pyinstaller-exe
description: "PyInstallerでPythonアプリをexe化するスキル。使用タイミング: (1) Pythonアプリを配布したい (2) exe化したい (3) 起動時間を短縮したい (4) ファイルサイズを削減したい"
---

# PyInstaller EXE化スキル

PythonアプリケーションをWindows実行ファイル（exe）に変換するノウハウ。

## 基本コマンド

### onefileモード（1ファイル配布）

```bash
pyinstaller --onefile --windowed --name "アプリ名" main.py
```

### onedirモード（フォルダ配布、起動が速い）

```bash
pyinstaller --onedir --windowed --name "アプリ名" main.py
```

## オプション一覧

| オプション | 説明 |
|-----------|------|
| `--onefile` | 単一exeファイルに圧縮（起動時に展開が必要） |
| `--onedir` | フォルダ配布（起動が速い） |
| `--windowed` | コンソール非表示（GUIアプリ向け） |
| `--console` | コンソール表示（CLIアプリ向け） |
| `--name "名前"` | 出力ファイル名 |
| `--add-data "src;dst"` | データファイル追加 |
| `--hidden-import "module"` | 隠れた依存モジュール追加 |
| `--exclude-module "module"` | 不要モジュール除外 |
| `--clean` | ビルド前にキャッシュクリア |

## サイズ削減テクニック

### 不要ライブラリの除外

使っていないライブラリを除外することでサイズを大幅削減できる。

```bash
pyinstaller --onedir --windowed --name "アプリ名" \
  --exclude-module torch \
  --exclude-module tensorflow \
  --exclude-module keras \
  --exclude-module IPython \
  --exclude-module jedi \
  --exclude-module sympy \
  --exclude-module scipy \
  --exclude-module numba \
  --exclude-module llvmlite \
  --exclude-module lxml \
  --exclude-module cryptography \
  --exclude-module pygments \
  main.py
```

**注意**: `openpyxl`、`win32com`、`PIL`（Pillow）はExcel/PNG保存機能で使用するため除外しないこと。

### 動的importモジュールの追加

関数内でimportされるライブラリはPyInstallerが自動検出できないため、`--hidden-import` で明示的に追加する：

```bash
--hidden-import openpyxl \
--hidden-import win32com.client \
--hidden-import pythoncom \
--hidden-import PIL \
--hidden-import PIL.Image \
--hidden-import PIL.ImageGrab \
--hidden-import PIL.ImageTk \
```

### 実績例

| 段階 | サイズ | 除外内容 |
|------|--------|----------|
| 最初 | 234MB | 除外なし |
| 1回目 | 121MB | torch, tensorflow, keras, IPython, jedi, sympy |
| 2回目 | 47MB | + scipy, numba, llvmlite, lxml, cryptography, pygments |
| onedir | 103MB(フォルダ) | 同上、ただし起動最速 |

## Matplotlib対応

### バックエンド設定（必須）

exe化するとグラフが表示されないことがある。main.pyの先頭に以下を追加：

```python
# main.py の先頭（importより前）
import matplotlib
matplotlib.use('TkAgg')

import tkinter as tk
# 以降の通常のimport...
```

### ビルド時にバックエンド指定

```bash
--hidden-import "matplotlib.backends.backend_tkagg"
```

## データファイルの扱い

### 方法1: 手動コピー（簡単）

ビルド後、exeと同じフォルダにデータファイルをコピー。

```
dist/
├── アプリ名.exe
├── config.json      ← 手動コピー
└── data.csv         ← 手動コピー
```

### 方法2: --add-data（exe内埋め込み）

```bash
--add-data "config.json;."
--add-data "data.csv;."
```

注意: この場合、コード側でPyInstaller対応のパス処理が必要：

```python
import sys
import os

def resource_path(relative_path):
    """PyInstaller対応のリソースパス取得"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(__file__), relative_path)

# 使用例
config_path = resource_path("config.json")
```

## 配布モード比較

| モード | メリット | デメリット |
|--------|----------|------------|
| onefile | 1ファイルで配布簡単 | 起動時に展開処理（遅い） |
| onedir | 起動が速い | フォルダ単位で配布 |

**推奨**: 起動時間が重要なら`--onedir`でフォルダをZIP圧縮して配布。

## 汎用 spec パターン (推奨)

バージョンが上がるたびに `.spec` を新規作成するとファイルが増え続けるため、`version.py` から `__version__` を読み取って exe 名を自動決定する汎用 `.spec` を 1 本だけ置く運用を推奨する。

### 温特アプリの汎用spec例 (`DEF_Command_Set_App.spec`)

```python
# -*- mode: python ; coding: utf-8 -*-
import os
import re

_spec_dir = os.path.dirname(os.path.abspath(SPEC))
with open(os.path.join(_spec_dir, 'version.py'), encoding='utf-8') as _f:
    _match = re.search(r'__version__\s*=\s*["\']([^"\']+)["\']', _f.read())
_version = _match.group(1) if _match else '0.00'
_app_name = f'DEF_Command_Set_App_v{_version}'


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['matplotlib.backends.backend_tkagg', 'openpyxl',
                   'win32com.client', 'pythoncom',
                   'PIL', 'PIL.Image', 'PIL.ImageGrab', 'PIL.ImageTk'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'tensorflow', 'keras', 'IPython', 'jedi', 'sympy',
              'scipy', 'numba', 'llvmlite', 'lxml', 'cryptography', 'pygments'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name=_app_name,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe, a.binaries, a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name=_app_name,
)
```

### ビルドコマンド

```bash
pyinstaller DEF_Command_Set_App.spec --clean --noconfirm
```

- `version.py` を更新してから実行するだけ
- `dist/DEF_Command_Set_App_v{バージョン}/` に自動で正しい名前で生成される
- **古いバージョンの .spec を増やさない** (バージョン管理は `version.py` と git 側に任せる)

### 注意点

- `pyinstaller` は COLLECT 時に `dist/{app_name}/` を一旦削除してから再作成するため、再ビルドのたびに `template/` や `*.json` などの手動コピーファイルも消える → 後述の「ビルド後の配布ファイルコピー」を必ず再実行すること

## 完全なビルドコマンド例

```bash
cd "プロジェクトフォルダ"

pyinstaller --onedir --windowed \
  --name "DEF評価アプリ_ver116" \
  --add-data "LIN_rules_index_expr_IJ_v6.csv;." \
  --add-data "config.json;." \
  --hidden-import "matplotlib.backends.backend_tkagg" \
  --hidden-import openpyxl \
  --hidden-import win32com.client \
  --hidden-import pythoncom \
  --hidden-import PIL \
  --hidden-import PIL.Image \
  --hidden-import PIL.ImageGrab \
  --hidden-import PIL.ImageTk \
  --exclude-module torch \
  --exclude-module tensorflow \
  --exclude-module keras \
  --exclude-module IPython \
  --exclude-module jedi \
  --exclude-module sympy \
  --exclude-module scipy \
  --exclude-module numba \
  --exclude-module llvmlite \
  --exclude-module lxml \
  --exclude-module cryptography \
  --exclude-module pygments \
  --clean \
  main.py
```

## ビルド後の配布ファイルコピー

ビルド完了後、exeと同じフォルダ（`dist/アプリ名/`）に以下の配布用ファイルをコピーすること：

| ファイル／フォルダ | 用途 |
|----------|------|
| `CHANGELOG.md` | バージョン情報表示用 |
| `app_settings.json` | アプリ起動時の設定値 |
| `config.json` | 起動設定 |
| `graph_settings.json` | グラフ描画設定 |
| `template/` | Excel 出力テンプレート一式 (dc_char / linearity_*.xlsx) |
| プロジェクト固有のデータファイル | 例: `pattern/`, `LIN_rules_index_expr_IJ_v6.csv` |

### 温特アプリの場合の完全コピーコマンド

```bash
# ビルド後に実行 (温特アプリ DEF_Command_Set_App)
cp CHANGELOG.md "dist/DEF_Command_Set_App_v{バージョン}/"
cp app_settings.json "dist/DEF_Command_Set_App_v{バージョン}/"
cp config.json "dist/DEF_Command_Set_App_v{バージョン}/"
cp graph_settings.json "dist/DEF_Command_Set_App_v{バージョン}/"
cp -r template "dist/DEF_Command_Set_App_v{バージョン}/"
# 必要に応じて pattern/ などもコピー
```

### 確認事項

コピー後は以下がdistフォルダ直下に揃っていることを確認:

```
dist/DEF_Command_Set_App_v{バージョン}/
├── DEF_Command_Set_App_v{バージョン}.exe
├── _internal/
├── template/              ← linearity/dc_char テンプレート
├── app_settings.json
├── config.json
├── graph_settings.json
└── CHANGELOG.md
```

**重要**: `--add-data` でexe内に埋め込んだファイルも、ユーザーが編集・参照できるようにdistフォルダにコピーする。
特に `template/*.xlsx` はユーザーが書式を調整する可能性があるため、外部ファイルとして配布すること。

## トラブルシューティング

### グラフが表示されない

1. main.pyの先頭に `matplotlib.use('TkAgg')` を追加
2. `--hidden-import "matplotlib.backends.backend_tkagg"` を指定

### 起動が遅い

1. `--onefile` → `--onedir` に変更
2. 不要ライブラリを `--exclude-module` で除外

### ファイルが読めない

1. データファイルをexeと同じフォルダにコピー
2. または `resource_path()` 関数でパス取得

### ウイルス対策ソフトが誤検知

- UPX圧縮を使うと誤検知されやすい
- `--noupx` オプションで回避可能

## 依存

- `pip install pyinstaller`
