"""Pins calculate_metrics against the 12-month sample in the screenshot/PDF."""
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import pytest

from metrics import calculate_metrics

# Reconstructed from docs/img/result_image_sample.png / finautomation_results.pdf:
# Total Revenue and Net Profit are given directly; Expenses = Revenue - Net Profit.
SAMPLE_REVENUES = [100000, 120000, 115000, 130000, 125000, 140000,
                    150000, 145000, 155000, 160000, 170000, 180000]
SAMPLE_EXPENSES = [60000, 70000, 65000, 80000, 75000, 85000,
                    90000, 88000, 95000, 100000, 110000, 115000]

SAMPLE_NET_PROFIT = [40000, 50000, 50000, 50000, 50000, 55000,
                      60000, 57000, 60000, 60000, 60000, 65000]
SAMPLE_GROWTH = [None, 20, -4.166666667, 13.04347826, -3.846153846, 12,
                  7.142857143, -3.333333333, 6.896551724, 3.225806452,
                  6.25, 5.882352941]
SAMPLE_MARGIN = [40, 41.66666667, 43.47826087, 38.46153846, 40, 39.28571429,
                  40, 39.31034483, 38.70967742, 37.5, 35.29411765, 36.11111111]


def test_screenshot_row_2_has_20_percent_growth_and_41_67_percent_margin():
    rows = calculate_metrics(SAMPLE_REVENUES, SAMPLE_EXPENSES)
    assert rows[1]["revenue_growth_pct"] == pytest.approx(20, abs=1e-6)
    assert rows[1]["profit_margin_pct"] == pytest.approx(41.66666667, abs=1e-6)


def test_all_twelve_screenshot_rows_match():
    rows = calculate_metrics(SAMPLE_REVENUES, SAMPLE_EXPENSES)
    assert len(rows) == 12
    for i, row in enumerate(rows):
        assert row["net_profit"] == SAMPLE_NET_PROFIT[i]
        assert row["total_revenue"] == SAMPLE_REVENUES[i]
        if SAMPLE_GROWTH[i] is None:
            assert row["revenue_growth_pct"] is None
        else:
            assert row["revenue_growth_pct"] == pytest.approx(SAMPLE_GROWTH[i], abs=1e-6)
        assert row["profit_margin_pct"] == pytest.approx(SAMPLE_MARGIN[i], abs=1e-6)


def test_first_row_has_no_growth_figure():
    rows = calculate_metrics(SAMPLE_REVENUES, SAMPLE_EXPENSES)
    assert rows[0]["revenue_growth_pct"] is None


def test_zero_previous_revenue_gives_none_instead_of_dividing_by_zero():
    rows = calculate_metrics([0, 50000], [0, 20000])
    assert rows[1]["revenue_growth_pct"] is None
    # Current-row revenue is also zero in row 0, so margin is None there too.
    assert rows[0]["profit_margin_pct"] is None


def test_zero_current_revenue_gives_none_margin_instead_of_dividing_by_zero():
    rows = calculate_metrics([50000, 0], [20000, 5000])
    assert rows[1]["profit_margin_pct"] is None
    assert rows[1]["net_profit"] == -5000


def test_mismatched_length_inputs_raise_value_error():
    with pytest.raises(ValueError):
        calculate_metrics([100000, 120000], [60000])
