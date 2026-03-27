"""DC特性試験の定数定義 - 期待値・許容誤差・アドレスコード"""
from dataclasses import dataclass
from typing import List


@dataclass(frozen=True)
class PositionTestPoint:
    """POSTION計測ポイント"""
    part: str           # "POS" or "NEG"
    address_code: str   # e.g. "FFFFF"
    display_code: str   # e.g. "FFFFF H"
    expected: float     # 期待値 (V)
    tolerance: float    # 許容誤差 (V)


# POSTION: 1デバイスあたり6ポイント（POS 3点 + NEG 3点）
POSITION_TEST_POINTS: List[PositionTestPoint] = [
    PositionTestPoint("POS", "FFFFF", "FFFFF H", +160.000, 0.160),
    PositionTestPoint("POS", "80000", "80000 H",    0.000, 0.100),
    PositionTestPoint("POS", "00000", "00000 H", -160.000, 0.160),
    PositionTestPoint("NEG", "00000", "00000 H", +160.000, 0.160),
    PositionTestPoint("NEG", "80000", "80000 H",    0.000, 0.100),
    PositionTestPoint("NEG", "FFFFF", "FFFFF H", -160.000, 0.160),
]

# 期待値文字列（テンプレート表示用）
POSITION_EXPECTED_STRINGS = [
    "160.000±0.160",
    "0.000±0.100",
    "-160.000±0.160",
    "160.000±0.160",
    "0.000±0.100",
    "-160.000±0.160",
]


@dataclass(frozen=True)
class LBCTestPoint:
    """LBC計測ポイント"""
    att: str            # "1/1", "1/2", "1/4"
    address_code: str   # "FFFF" or "0000"
    display_code: str   # "FFFF H" or "0000 H"
    expected_pos: float # POS期待値 (V)
    expected_neg: float # NEG期待値 (V)
    tolerance: float    # 許容誤差 (V)


# LBC: 1デバイスあたり6ポイント（ATT 3種 × コード 2種）
LBC_TEST_POINTS: List[LBCTestPoint] = [
    LBCTestPoint("1/1", "FFFF", "FFFF H", +6.180, -6.180, 0.062),
    LBCTestPoint("1/1", "0000", "0000 H", -6.180, +6.180, 0.062),
    LBCTestPoint("1/2", "FFFF", "FFFF H", +3.090, -3.090, 0.031),
    LBCTestPoint("1/2", "0000", "0000 H", -3.090, +3.090, 0.031),
    LBCTestPoint("1/4", "FFFF", "FFFF H", +1.545, -1.545, 0.015),
    LBCTestPoint("1/4", "0000", "0000 H", -1.545, +1.545, 0.015),
]

# 期待値文字列（テンプレート表示用、FFFFのみ表示）
LBC_EXPECTED_STRINGS = [
    "±6.180±0.062",
    "",
    "±3.090±0.031",
    "",
    "±1.545±0.015",
    "",
]


@dataclass(frozen=True)
class MoniTestPoint:
    """moni計測ポイント"""
    part: str           # "POS" or "NEG"
    address_code: str   # "FFFFF" or "00000"
    display_code: str   # "FFFFF H" or "00000 H"
    expected: float     # 期待値 (V)
    tolerance: float    # 許容誤差 (V)


# moni: 1デバイスあたり4ポイント（POS 2点 + NEG 2点）
MONI_TEST_POINTS: List[MoniTestPoint] = [
    MoniTestPoint("POS", "FFFFF", "FFFFF H", +5.00, 0.05),
    MoniTestPoint("POS", "00000", "00000 H", -5.00, 0.05),
    MoniTestPoint("NEG", "00000", "00000 H", +5.00, 0.05),
    MoniTestPoint("NEG", "FFFFF", "FFFFF H", -5.00, 0.05),
]

# 期待値文字列（テンプレート表示用）
MONI_EXPECTED_STRINGS = [
    "+5.00±0.05",
    "-5.00±0.05",
    "+5.00±0.05",
    "-5.00±0.05",
]

# DMM レンジ設定
DMM_RANGE_POSITION = "100"   # 160V計測に対応
DMM_RANGE_LBC = "10"         # 6.18V計測に対応
DMM_RANGE_MONI = "10"        # 5V計測に対応
