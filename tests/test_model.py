"""Tests for the model module (with mocked Chronos-2 pipeline)."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pytest
import numpy as np
import pandas as pd
from unittest.mock import patch, MagicMock


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


def _make_mock_pipeline(horizon):
    """Create a mock Chronos2Pipeline that returns realistic predictions."""
    mock_pipeline = MagicMock()

    def mock_predict_df(context_df, future_df=None, prediction_length=90,
                        quantile_levels=None, id_column="id",
                        timestamp_column="timestamp", target="target"):
        if quantile_levels is None:
            quantile_levels = [0.1, 0.25, 0.5, 0.75, 0.9]

        unique_ids = context_df[id_column].unique()
        frames = []
        for item_id in unique_ids:
            item_ctx = context_df[context_df[id_column] == item_id]
            last_val = item_ctx[target].iloc[-1]
            last_ts = item_ctx[timestamp_column].max()
            future_dates = pd.bdate_range(last_ts + pd.Timedelta(days=1), periods=prediction_length)

            pred_data = {id_column: item_id, timestamp_column: future_dates}
            for q in quantile_levels:
                noise = np.random.normal(0, 2, prediction_length).cumsum()
                offset = (q - 0.5) * 5
                pred_data[str(q)] = last_val + noise + offset
            frames.append(pd.DataFrame(pred_data))

        return pd.concat(frames, ignore_index=True)

    mock_pipeline.predict_df = mock_predict_df
    return mock_pipeline


class TestForecast:
    @patch("src.model.get_pipeline")
    def test_forecast_returns_expected_keys(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(90)
        from src.model import forecast
        result = forecast(sample_features, horizon=90)

        assert "q10" in result
        assert "q25" in result
        assert "median" in result
        assert "q75" in result
        assert "q90" in result
        assert "dates" in result

    @patch("src.model.get_pipeline")
    def test_forecast_output_length(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(30)
        from src.model import forecast
        result = forecast(sample_features, horizon=30)

        assert len(result["median"]) == 30
        assert len(result["q10"]) == 30

    @patch("src.model.get_pipeline")
    def test_quantile_ordering(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(30)
        from src.model import forecast
        result = forecast(sample_features, horizon=30)

        # q10 <= q25 <= median <= q75 <= q90 (on average)
        assert np.mean(result["q10"]) <= np.mean(result["median"])
        assert np.mean(result["median"]) <= np.mean(result["q90"])


class TestForecastAtOrigin:
    @patch("src.model.get_pipeline")
    def test_no_leakage(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(30)
        from src.model import forecast_at_origin

        origin_idx = 400
        origin_date = sample_features.index[origin_idx - 1]

        # Capture the context_df passed to predict_df
        calls = []
        original_predict = mock_get_pipeline.return_value.predict_df

        def capture_predict(context_df, **kwargs):
            calls.append(context_df)
            return original_predict(context_df, **kwargs)

        mock_get_pipeline.return_value.predict_df = capture_predict

        forecast_at_origin(sample_features, origin_idx, 30)

        assert len(calls) == 1
        ctx = calls[0]
        target_ctx = ctx[ctx["id"] == "ct1_close"]
        assert target_ctx["timestamp"].max() <= origin_date


class TestForecastMultiHorizon:
    @patch("src.model.get_pipeline")
    def test_all_horizons_present(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(90)
        from src.model import forecast_multi_horizon

        results = forecast_multi_horizon(sample_features, horizons=[30, 60, 90])

        assert 30 in results
        assert 60 in results
        assert 90 in results

    @patch("src.model.get_pipeline")
    def test_horizon_slice_lengths(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(90)
        from src.model import forecast_multi_horizon

        results = forecast_multi_horizon(sample_features, horizons=[30, 60, 90])

        assert len(results[30]["median"]) == 30
        assert len(results[60]["median"]) == 60
        assert len(results[90]["median"]) == 90

    @patch("src.model.get_pipeline")
    def test_point_estimates_present(self, mock_get_pipeline, sample_features):
        mock_get_pipeline.return_value = _make_mock_pipeline(90)
        from src.model import forecast_multi_horizon

        results = forecast_multi_horizon(sample_features, horizons=[30, 60, 90])

        for h in [30, 60, 90]:
            assert "point" in results[h]
            assert "low" in results[h]
            assert "high" in results[h]
