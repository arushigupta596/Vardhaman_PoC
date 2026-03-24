"""Tests for the covariates module."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest
import numpy as np
import pandas as pd
from src.covariates import (
    load_config,
    validate_covariates,
    build_context_df,
    build_context_df_cross_learning,
    build_future_df,
    get_covariate_summary,
)


@pytest.fixture
def sample_features():
    """Create a sample features DataFrame for testing."""
    dates = pd.bdate_range("2023-01-02", periods=600, freq="B")
    np.random.seed(42)
    n = len(dates)
    df = pd.DataFrame({
        "ct1_close": 80 + np.cumsum(np.random.normal(0, 0.5, n)),
        "ct2_close": 80.2 + np.cumsum(np.random.normal(0, 0.5, n)),
        "ct3_close": 80.4 + np.cumsum(np.random.normal(0, 0.5, n)),
        "dxy": 103 + np.cumsum(np.random.normal(0, 0.2, n)),
        "wti_crude": 75 + np.cumsum(np.random.normal(0, 0.3, n)),
        "comm_net_z": np.random.normal(0, 1, n),
        "traders_noncomm_long": np.abs(np.random.normal(50000, 10000, n)),
        "spec_net_pct": np.random.normal(0.2, 0.1, n),
        "conc_4_short": np.random.normal(30, 5, n),
        "pdsi_severe_drought": (np.random.random(n) > 0.8).astype(float),
        "realised_vol_21d": np.abs(np.random.normal(0.2, 0.05, n)),
        "noaa_pdsi": np.random.normal(-1, 2, n),
        "seas_sin_annual": np.sin(2 * np.pi * dates.dayofyear / 365.25),
        "seas_cos_annual": np.cos(2 * np.pi * dates.dayofyear / 365.25),
        "flag_planting": ((dates.month >= 4) & (dates.month <= 6)).astype(float),
        "flag_boll_dev": ((dates.month >= 7) & (dates.month <= 8)).astype(float),
        "flag_harvest": ((dates.month >= 9) & (dates.month <= 11)).astype(float),
        "flag_wasde": ((dates.day >= 9) & (dates.day <= 13)).astype(float),
    }, index=dates)
    df.index.name = "date"
    return df


@pytest.fixture
def config():
    return load_config()


class TestValidateCovariates:
    def test_valid_features(self, sample_features, config):
        result = validate_covariates(sample_features, config)
        assert result["valid"] is True
        assert len(result["missing"]) == 0

    def test_missing_covariate(self, sample_features, config):
        df = sample_features.drop(columns=["dxy"])
        result = validate_covariates(df, config)
        assert result["valid"] is False
        assert "dxy" in result["missing"]

    def test_high_nan_detection(self, sample_features, config):
        df = sample_features.copy()
        df.loc[df.index[:400], "traders_noncomm_long"] = np.nan  # >10% NaN
        result = validate_covariates(df, config)
        high_nan_cols = [col for col, _ in result["high_nan"]]
        assert "traders_noncomm_long" in high_nan_cols


class TestBuildContextDf:
    def test_output_shape(self, sample_features, config):
        origin_idx = 500
        context_df = build_context_df(sample_features, origin_idx, config)
        assert len(context_df) == origin_idx
        assert "id" in context_df.columns
        assert "timestamp" in context_df.columns
        assert "target" in context_df.columns

    def test_no_leakage(self, sample_features, config):
        origin_idx = 400
        context_df = build_context_df(sample_features, origin_idx, config)
        origin_date = sample_features.index[origin_idx - 1]
        assert context_df["timestamp"].max() <= origin_date

    def test_covariates_present(self, sample_features, config):
        context_df = build_context_df(sample_features, 500, config)
        for cov in config["data"]["past_covariates"]:
            assert cov in context_df.columns
        for cov in config["data"]["known_future_covariates"]:
            assert cov in context_df.columns

    def test_target_values_match(self, sample_features, config):
        origin_idx = 300
        # Use normalize=False to verify raw target values are passed through
        context_df = build_context_df(sample_features, origin_idx, config, normalize=False)
        np.testing.assert_array_almost_equal(
            context_df["target"].values,
            sample_features["ct1_close"].iloc[:origin_idx].values,
        )


class TestBuildContextDfCrossLearning:
    def test_multiple_series(self, sample_features, config):
        context_df = build_context_df_cross_learning(sample_features, 500, config)
        unique_ids = context_df["id"].unique()
        # Should have target + cross-learning series
        assert "ct1_close" in unique_ids
        assert len(unique_ids) > 1

    def test_cross_learning_series_present(self, sample_features, config):
        context_df = build_context_df_cross_learning(sample_features, 500, config)
        unique_ids = set(context_df["id"].unique())
        for series in config["data"]["cross_learning_series"]:
            assert series in unique_ids


class TestBuildFutureDf:
    def test_output_length(self, sample_features, config):
        last_date = sample_features.index[-1]
        future_df = build_future_df(last_date, 90, config)
        # Should have rows for target + cross-learning series
        target_rows = future_df[future_df["id"] == "ct1_close"]
        assert len(target_rows) == 90

    def test_no_target_column(self, sample_features, config):
        last_date = sample_features.index[-1]
        future_df = build_future_df(last_date, 30, config)
        assert "target" not in future_df.columns

    def test_seasonal_values_bounded(self, sample_features, config):
        last_date = sample_features.index[-1]
        future_df = build_future_df(last_date, 90, config)
        target_future = future_df[future_df["id"] == "ct1_close"]
        if "seas_sin_annual" in target_future.columns:
            assert target_future["seas_sin_annual"].between(-1, 1).all()
            assert target_future["seas_cos_annual"].between(-1, 1).all()

    def test_flags_binary(self, sample_features, config):
        last_date = sample_features.index[-1]
        future_df = build_future_df(last_date, 90, config)
        target_future = future_df[future_df["id"] == "ct1_close"]
        for flag in ["flag_planting", "flag_boll_dev", "flag_harvest", "flag_wasde"]:
            if flag in target_future.columns:
                assert target_future[flag].isin([0.0, 1.0]).all()

    def test_future_dates_after_last(self, sample_features, config):
        last_date = sample_features.index[-1]
        future_df = build_future_df(last_date, 30, config)
        target_future = future_df[future_df["id"] == "ct1_close"]
        assert target_future["timestamp"].min() > last_date


class TestGetCovariateSummary:
    def test_output_columns(self, sample_features, config):
        summary = get_covariate_summary(sample_features, config)
        assert "Covariate" in summary.columns
        assert "Current" in summary.columns
        assert "Z-Score" in summary.columns
        assert "Status" in summary.columns

    def test_all_covariates_present(self, sample_features, config):
        summary = get_covariate_summary(sample_features, config)
        for cov in config["data"]["past_covariates"]:
            assert cov in summary["Covariate"].values
