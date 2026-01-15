# 温特アプリ (DEF Command Set App)

温度特性試験用のPythonデスクトップアプリケーション

## 概要

- **バージョン**: 1.0
- **GUI**: Tkinter
- **用途**: パターン試験・温度特性グラフ表示

## 対応ハードウェア

| 機器 | 用途 | 通信 |
|------|------|------|
| HP 3458A | デジタルマルチメーター (DMM) | GPIB |
| HP 3499B | スイッチ/スキャナー | GPIB |
| DEFデバイス | DAC制御 | シリアル (38400bps) |

## ファイル構成

```
温特/
├── main.py                  # エントリポイント (178行)
├── gpib_controller.py       # GPIB通信制御 - pyvisa (297行)
├── serial_manager.py        # シリアル通信 - pyserial (79行)
├── app_settings.json        # アプリ設定
├── graph_settings.json      # グラフ設定
├── tabs/
│   ├── communication_tab.py # 通信設定タブ (391行)
│   ├── test_tab.py          # パターンテストタブ (1135行) ※最大
│   ├── graph_tab.py         # グラフ描画タブ (307行)
│   ├── dac_tab.py           # DAC操作タブ (326行)
│   ├── file_tab.py          # ファイル保存タブ
│   ├── scanner_tab.py       # スキャナータブ
│   ├── dmm3458a_tab.py      # DMM3458A制御タブ
│   └── measurement_window.py # 計測ウィンドウ (776行)
├── utils/
│   ├── graph_plotter.py     # グラフ描画 - matplotlib (388行)
│   ├── csv_logger.py        # CSV保存機能
│   ├── logger.py            # ログ機能
│   └── validators.py        # バリデーション
└── pattern/
    └── pattern.csv          # パターンデータ
```

## 通信プロトコル

### シリアル通信 (DEFコマンド)

```
DEF {n} {command}\r          # 基本コマンド (n=0-5)
DEF {n} DAC {P|L} {HEX}\r    # DAC設定
  - P: Position (20bit, 5桁HEX: 00000-FFFFF)
  - L: LBC (16bit, 4桁HEX: 0000-FFFF)
```

プリセット値:
- +Full: Position=FFFFF, LBC=FFFF
- Center: Position=80000, LBC=8000
- -Full: Position=00000, LBC=0000

### GPIB通信

**3458A (DMM)**:
- `TRIG SGL` - 単発トリガ
- `*IDN?` - 機器識別

**3499B (Scanner)**:
- `OPEN (@1xx)` - チャンネルOPEN
- `CLOSE (@1xx)` - チャンネルCLOSE

## 主要クラス

| クラス | ファイル | 役割 |
|--------|----------|------|
| MainApplication | main.py | メインウィンドウ |
| GPIBController | gpib_controller.py | GPIB通信管理 |
| SerialManager | serial_manager.py | シリアル通信管理 |
| TestTab | tabs/test_tab.py | パターンテスト実行 |
| MeasurementWindow | tabs/measurement_window.py | 自動計測 |
| LSBGraphPlotter | utils/graph_plotter.py | グラフ描画 |

## 修正時の注意点

1. **test_tab.py** が最も複雑 (1135行) - パターン実行・ホールド・スキップ機能
2. **measurement_window.py** - CSV保存・自動計測ループ
3. **設定ファイル** は `app_settings.json` に集約
4. **Neg極性時** はHEX値をビット反転して送信

## 依存ライブラリ

- tkinter (標準)
- pyvisa (GPIB/VISA通信)
- pyserial (シリアル通信)
- matplotlib (グラフ描画)
- numpy (数値計算)

## 開発ルール

- コード修正後は必ずコミット確認を求めること
- 修正前に対象ファイルのみ読み込む（全体読み込み禁止）
- 変更内容を簡潔に説明してからコミットする
