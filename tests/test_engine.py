
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from engine import (
    calc_middle_mile_rub,
    calc_volume_liters,
    calc_volumetric_weight_kg,
    load_tariffs,
)


def test_volume_liters_rounding():
    assert calc_volume_liters(11, 40, 84) == 37


def test_middle_mile_example():
    tariffs = load_tariffs(ROOT / "data" / "yandex_tariffs.xlsx")
    assert calc_middle_mile_rub(37, tariffs) == 862


def test_volumetric_weight():
    assert round(calc_volumetric_weight_kg(50, 40, 30), 2) == 12.0
