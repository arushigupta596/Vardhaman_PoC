"""
Chronos-2 multivariate forecasting model with covariates and cross-learning.

Uses Chronos-2 (encoder-only) via predict_df API:
  - Past covariates: DXY, WTI, CFTC positioning (noncomm long, spec net, conc short),
    drought, volatility, PDSI
  - Known future covariates: seasonality, crop calendar, WASDE flags
  - Cross-learning: group attention across DXY + WTI series
"""
import torch
import numpy as np
import pandas as pd
from pathlib import Path

from src import covariates

RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"
QUANTILES = [0.1, 0.25, 0.5, 0.75, 0.9]

_pipeline = None
_bolt_pipeline = None
_config = None


def get_config():
    global _config
    if _config is None:
        _config = covariates.load_config()
    return _config


def get_pipeline():
    """Lazy-load the Chronos-2 pipeline."""
    global _pipeline
    if _pipeline is None:
        from chronos import Chronos2Pipeline

        config = get_config()
        model_id = config["model"]["model_id"]
        device_map = config["model"]["device_map"]

        print(f"[model] Loading {model_id} …")
        _pipeline = Chronos2Pipeline.from_pretrained(
            model_id,
            device_map=device_map,
        )
        print(f"  → Loaded Chronos-2 pipeline on {device_map}")
    return _pipeline


def get_bolt_pipeline():
    """Lazy-load the Chronos-Bolt pipeline for univariate ensemble."""
    global _bolt_pipeline
    if _bolt_pipeline is None:
        from chronos import BaseChronosPipeline

        bolt_id = "amazon/chronos-bolt-base"
        config = get_config()
        device_map = config["model"]["device_map"]

        print(f"[model] Loading {bolt_id} …")
        _bolt_pipeline = BaseChronosPipeline.from_pretrained(
            bolt_id, device_map=device_map, dtype=torch.float32
        )
        print(f"  → Loaded Chronos-Bolt pipeline on {device_map}")
    return _bolt_pipeline


def forecast_bolt_univariate(
    series: pd.Series,
    horizon: int,
    quantile_levels: list = None,
) -> dict:
    """
    Univariate forecast using Chronos-Bolt-Base.
    Returns same dict format as forecast() for easy ensemble combination.
    """
    if quantile_levels is None:
        quantile_levels = QUANTILES

    pipeline = get_bolt_pipeline()
    context = series.dropna().values[-512:]
    context_tensor = torch.tensor(context, dtype=torch.float32).unsqueeze(0)

    quantile_preds, _ = pipeline.predict_quantiles(
        inputs=context_tensor,
        prediction_length=horizon,
        quantile_levels=quantile_levels,
    )

    # quantile_preds shape: (1, horizon, n_quantiles)
    preds = quantile_preds[0].numpy()  # (horizon, n_quantiles)

    return {
        "q10": preds[:, 0],
        "q25": preds[:, 1],
        "median": preds[:, 2],
        "q75": preds[:, 3],
        "q90": preds[:, 4],
        "mean": preds[:, 2],
    }


def forecast_ensemble(
    features_df: pd.DataFrame,
    origin_idx: int,
    horizon: int,
    config: dict = None,
    weight_chronos2: float = 0.6,
    weight_bolt: float = 0.4,
) -> dict:
    """
    Ensemble forecast combining Chronos-2 multivariate + Chronos-Bolt univariate.
    Weighted average of quantile predictions.
    """
    if config is None:
        config = get_config()

    # Chronos-2 multivariate forecast
    c2_result = forecast_at_origin(features_df, origin_idx, horizon, config=config, cross_learning=True)

    # Chronos-Bolt univariate forecast
    target_col = config["data"]["target"]
    series = features_df[target_col].iloc[:origin_idx]
    bolt_result = forecast_bolt_univariate(series, horizon, config["forecast"]["quantile_levels"])

    # Weighted combination
    w1, w2 = weight_chronos2, weight_bolt
    combined = {}
    for key in ["q10", "q25", "median", "q75", "q90", "mean"]:
        if key in c2_result and key in bolt_result:
            combined[key] = w1 * c2_result[key] + w2 * bolt_result[key]
        elif key in c2_result:
            combined[key] = c2_result[key]

    return combined


def forecast(
    features_df: pd.DataFrame,
    horizon: int,
    config: dict = None,
    cross_learning: bool = True,
) -> dict:
    """
    Generate quantile forecasts using Chronos-2 with covariates.

    Args:
        features_df: Full features DataFrame with DatetimeIndex
        horizon: Number of trading days to forecast
        config: Config dict (loaded from yaml if None)
        cross_learning: If True, include cross-learning series

    Returns:
        dict with keys: q10, q25, median, q75, q90, dates
        Each value is a numpy array of length `horizon`.
    """
    if config is None:
        config = get_config()

    pipeline = get_pipeline()
    origin_idx = len(features_df)

    # Build context DataFrame (long-format with covariates)
    if cross_learning:
        context_df = covariates.build_context_df_cross_learning(
            features_df, origin_idx, config
        )
    else:
        context_df = covariates.build_context_df(
            features_df, origin_idx, config
        )

    # Build future covariate DataFrame (with normalization stats from context)
    last_date = features_df.index[-1]
    norm_stats = getattr(context_df, "_norm_stats", None)
    future_df = covariates.build_future_df(last_date, horizon, config, covariate_stats=norm_stats)

    id_col = config["data"]["id_column"]
    ts_col = config["data"]["timestamp_column"]
    quantile_levels = config["forecast"]["quantile_levels"]

    # Chronos-2 predict_df — multivariate with covariates + cross-learning
    pred_df = pipeline.predict_df(
        context_df,
        future_df=future_df,
        prediction_length=horizon,
        quantile_levels=quantile_levels,
        id_column=id_col,
        timestamp_column=ts_col,
        target="target",
    )

    # Extract predictions for the target series (ct1_close)
    target_col = config["data"]["target"]
    ct1_preds = pred_df[pred_df[id_col] == target_col].sort_values(ts_col)

    result = {
        "q10": ct1_preds[str(quantile_levels[0])].values,
        "q25": ct1_preds[str(quantile_levels[1])].values,
        "median": ct1_preds[str(quantile_levels[2])].values,
        "q75": ct1_preds[str(quantile_levels[3])].values,
        "q90": ct1_preds[str(quantile_levels[4])].values,
        "mean": ct1_preds[str(quantile_levels[2])].values,  # median as point estimate
        "dates": ct1_preds[ts_col].values,
    }
    return result


def forecast_at_origin(
    features_df: pd.DataFrame,
    origin_idx: int,
    horizon: int,
    config: dict = None,
    cross_learning: bool = True,
) -> dict:
    """
    Forecast from a specific origin index (for backtesting).

    Strict no-leakage: only uses features_df[:origin_idx].
    """
    if config is None:
        config = get_config()

    pipeline = get_pipeline()

    if cross_learning:
        context_df = covariates.build_context_df_cross_learning(
            features_df, origin_idx, config
        )
    else:
        context_df = covariates.build_context_df(
            features_df, origin_idx, config
        )

    last_date = features_df.index[origin_idx - 1]
    norm_stats = getattr(context_df, "_norm_stats", None)
    future_df = covariates.build_future_df(last_date, horizon, config, covariate_stats=norm_stats)

    id_col = config["data"]["id_column"]
    ts_col = config["data"]["timestamp_column"]
    quantile_levels = config["forecast"]["quantile_levels"]

    pred_df = pipeline.predict_df(
        context_df,
        future_df=future_df,
        prediction_length=horizon,
        quantile_levels=quantile_levels,
        id_column=id_col,
        timestamp_column=ts_col,
        target="target",
    )

    target_col = config["data"]["target"]
    ct1_preds = pred_df[pred_df[id_col] == target_col].sort_values(ts_col)

    return {
        "q10": ct1_preds[str(quantile_levels[0])].values,
        "q25": ct1_preds[str(quantile_levels[1])].values,
        "median": ct1_preds[str(quantile_levels[2])].values,
        "q75": ct1_preds[str(quantile_levels[3])].values,
        "q90": ct1_preds[str(quantile_levels[4])].values,
        "mean": ct1_preds[str(quantile_levels[2])].values,
    }


def forecast_multi_horizon(
    features_df: pd.DataFrame,
    horizons: list = None,
    config: dict = None,
    cross_learning: bool = True,
) -> dict:
    """
    Forecast at multiple horizons (30, 60, 90 days).

    Single call for max horizon, then slices for shorter horizons.
    Returns {horizon: forecast_dict}.
    """
    if config is None:
        config = get_config()
    if horizons is None:
        horizons = config["forecast"]["horizons"]

    max_h = max(horizons)

    # Single call for the longest horizon
    full = forecast(features_df, horizon=max_h, config=config, cross_learning=cross_learning)

    results = {}
    for h in horizons:
        results[h] = {k: v[:h] for k, v in full.items() if isinstance(v, np.ndarray)}
        results[h]["point"] = full["median"][h - 1]
        results[h]["low"] = full["q10"][h - 1]
        results[h]["high"] = full["q90"][h - 1]

    return results


def load_bias_estimates() -> dict:
    """
    Load per-horizon bias estimates from backtest diagnostics.
    Returns {horizon: mean_signed_error} or empty dict if unavailable.
    """
    diag_path = RESULTS_DIR / "backtest_metrics.csv"
    if not diag_path.exists():
        return {}
    try:
        df = pd.read_csv(diag_path)
        if "signed_error" not in df.columns:
            # Compute from predicted - actual
            if "predicted" in df.columns and "actual" in df.columns:
                df["signed_error"] = df["predicted"] - df["actual"]
            else:
                return {}
        bias = df.groupby("horizon")["signed_error"].mean().to_dict()
        return {int(k): v for k, v in bias.items()}
    except Exception:
        return {}


def apply_bias_correction(forecast_dict: dict, horizon: int, bias_estimates: dict = None) -> dict:
    """
    Apply post-hoc bias correction to a forecast.
    Subtracts the estimated bias from all quantile predictions.
    """
    if bias_estimates is None:
        bias_estimates = load_bias_estimates()
    if not bias_estimates or horizon not in bias_estimates:
        return forecast_dict

    bias = bias_estimates[horizon]
    corrected = {}
    for key, val in forecast_dict.items():
        if isinstance(val, np.ndarray) and val.dtype.kind == 'f':
            corrected[key] = val - bias
        else:
            corrected[key] = val
    return corrected


if __name__ == "__main__":
    features = pd.read_parquet("data/features/features.parquet")
    print(f"Features: {len(features)} rows × {len(features.columns)} columns")
    print(f"Last date: {features.index[-1].date()}")

    results = forecast_multi_horizon(features)
    for h, r in results.items():
        print(f"\n{h}d forecast:")
        print(f"  Point: {r['point']:.2f}")
        print(f"  Range: [{r['low']:.2f}, {r['high']:.2f}]")
        print(f"  Median path (first 5): {r['median'][:5].round(2)}")
