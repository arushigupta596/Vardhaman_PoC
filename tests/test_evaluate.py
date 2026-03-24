"""Tests for the evaluate module."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest
import numpy as np
import pandas as pd
from src.evaluate import crps_quantile, evaluate_forecast, diebold_mariano_test


class TestCRPS:
    def test_perfect_forecast(self):
        actual = 80.0
        crps = crps_quantile(79.0, 80.0, 81.0, actual)
        assert crps >= 0

    def test_crps_positive(self):
        crps = crps_quantile(70.0, 75.0, 80.0, 85.0)
        assert crps > 0

    def test_crps_increases_with_error(self):
        crps_good = crps_quantile(79.0, 80.0, 81.0, 80.0)
        crps_bad = crps_quantile(70.0, 75.0, 80.0, 90.0)
        assert crps_bad > crps_good


class TestEvaluateForecast:
    def test_basic_evaluation(self):
        forecast_dict = {
            "median": np.array([80.0, 81.0, 82.0]),
            "q10": np.array([78.0, 79.0, 80.0]),
            "q90": np.array([82.0, 83.0, 84.0]),
        }
        actuals = pd.Series([79.5, 80.5, 81.5, 82.5])
        metrics = evaluate_forecast(forecast_dict, actuals, horizon=3)

        assert metrics is not None
        assert "mae" in metrics
        assert "rmse" in metrics
        assert "crps" in metrics
        assert "dir_acc" in metrics
        assert "coverage" in metrics
        assert "mase" in metrics
        assert metrics["mae"] >= 0
        assert metrics["rmse"] >= 0

    def test_returns_none_for_insufficient_data(self):
        forecast_dict = {
            "median": np.array([80.0, 81.0, 82.0]),
            "q10": np.array([78.0, 79.0, 80.0]),
            "q90": np.array([82.0, 83.0, 84.0]),
        }
        actuals = pd.Series([79.5])  # Only 1 point, need 3
        metrics = evaluate_forecast(forecast_dict, actuals, horizon=3)
        assert metrics is None

    def test_coverage_within_interval(self):
        forecast_dict = {
            "median": np.array([80.0]),
            "q10": np.array([78.0]),
            "q90": np.array([82.0]),
        }
        actuals = pd.Series([79.0, 80.0])  # actual=80.0 is within [78, 82]
        metrics = evaluate_forecast(forecast_dict, actuals, horizon=1)
        assert metrics["coverage"] == 1.0

    def test_coverage_outside_interval(self):
        forecast_dict = {
            "median": np.array([80.0]),
            "q10": np.array([78.0]),
            "q90": np.array([82.0]),
        }
        actuals = pd.Series([85.0, 85.0])  # actual at horizon=1 is 85.0, outside [78, 82]
        metrics = evaluate_forecast(forecast_dict, actuals, horizon=1)
        assert metrics["coverage"] == 0.0


class TestDieboldMariano:
    def test_equal_errors(self):
        errors = np.random.normal(0, 1, 100)
        t_stat, p_val = diebold_mariano_test(errors, errors)
        assert abs(t_stat) < 1e-5
        assert p_val > 0.9

    def test_significantly_different(self):
        np.random.seed(42)
        errors_a = np.random.normal(0, 1, 200)
        errors_b = np.random.normal(0, 3, 200)  # Much larger errors
        t_stat, p_val = diebold_mariano_test(errors_a, errors_b)
        # Model A should be significantly better
        assert t_stat < 0
        assert p_val < 0.05
