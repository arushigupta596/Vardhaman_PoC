"""
Chronos-2 multivariate forecasting model with covariates and cross-learning.

Uses Chronos-2 (encoder-only) via predict_df API:
  - Past covariates: DXY, WTI, CFTC positioning (noncomm long, spec net, conc short),
    drought, volatility, PDSI
  - Known future covariates: seasonality, crop calendar, WASDE flags
  - Cross-learning: group attention across DXY + WTI series
"""
import numpy as np
import pandas as pd
from pathlib import Path

# Lazy imports for heavy dependencies (torch, chronos) — only loaded when inference is needed
torch = None
covariates = None

def _ensure_torch():
    global torch
    if torch is None:
        import torch as _torch
        torch = _torch
    return torch

def _ensure_covariates():
    global covariates
    if covariates is None:
        from src import covariates as _cov
        covariates = _cov
    return covariates

RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"
QUANTILES = [0.1, 0.25, 0.5, 0.75, 0.9]

# Human-readable labels for secondary targets
SECONDARY_TARGET_LABELS = {
    "realised_vol_21d": "Realized Volatility (21d annualized)",
    "open_interest": "Open Interest (5d avg proxy)",
    "roll_yield": "Term Spread CT2-CT1 (% of CT1)",
}


def _extract_secondary_predictions(pred_df, secondary_targets, id_col, ts_col, quantile_levels):
    """
    Extract predictions for secondary targets from the Chronos-2 predict_df output.

    Returns:
        {target_name: {"median": array, "q10": array, "q90": array}}
    """
    results = {}
    for tgt in secondary_targets:
        tgt_preds = pred_df[pred_df[id_col] == tgt].sort_values(ts_col)
        if len(tgt_preds) == 0:
            continue
        results[tgt] = {
            "q10": tgt_preds[str(quantile_levels[0])].values,
            "q25": tgt_preds[str(quantile_levels[1])].values,
            "median": tgt_preds[str(quantile_levels[2])].values,
            "q75": tgt_preds[str(quantile_levels[3])].values,
            "q90": tgt_preds[str(quantile_levels[4])].values,
        }
    return results

_pipeline = None
_bolt_pipeline = None
_config = None


def get_config():
    global _config
    if _config is None:
        _config = _ensure_covariates().load_config()
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
            bolt_id, device_map=device_map, dtype=_ensure_torch().float32
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
    _torch = _ensure_torch()
    context_tensor = _torch.tensor(context, dtype=_torch.float32).unsqueeze(0)

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
    return_components: bool = False,
) -> dict:
    """
    Ensemble forecast combining Chronos-2 multivariate + Chronos-Bolt univariate.
    Weighted average of quantile predictions.

    If return_components=True, also includes 'c2_median' and 'bolt_median' keys
    for ensemble weight optimization.
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

    if return_components:
        combined["c2_median"] = c2_result["median"]
        combined["bolt_median"] = bolt_result["median"]

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
    _cov = _ensure_covariates()
    if cross_learning:
        context_df = _cov.build_context_df_cross_learning(
            features_df, origin_idx, config
        )
    else:
        context_df = _cov.build_context_df(
            features_df, origin_idx, config
        )

    # Build future covariate DataFrame (with normalization stats from context)
    last_date = features_df.index[-1]
    norm_stats = getattr(context_df, "_norm_stats", None)
    future_df = _cov.build_future_df(last_date, horizon, config, covariate_stats=norm_stats)

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

    # Extract secondary target predictions
    secondary_targets = config["data"].get("secondary_targets", [])
    if secondary_targets:
        result["secondary"] = _extract_secondary_predictions(
            pred_df, secondary_targets, id_col, ts_col, quantile_levels
        )

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

    _cov = _ensure_covariates()
    if cross_learning:
        context_df = _cov.build_context_df_cross_learning(
            features_df, origin_idx, config
        )
    else:
        context_df = _cov.build_context_df(
            features_df, origin_idx, config
        )

    last_date = features_df.index[origin_idx - 1]
    norm_stats = getattr(context_df, "_norm_stats", None)
    future_df = _cov.build_future_df(last_date, horizon, config, covariate_stats=norm_stats)

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

    result = {
        "q10": ct1_preds[str(quantile_levels[0])].values,
        "q25": ct1_preds[str(quantile_levels[1])].values,
        "median": ct1_preds[str(quantile_levels[2])].values,
        "q75": ct1_preds[str(quantile_levels[3])].values,
        "q90": ct1_preds[str(quantile_levels[4])].values,
        "mean": ct1_preds[str(quantile_levels[2])].values,
    }

    # Extract secondary target predictions
    secondary_targets = config["data"].get("secondary_targets", [])
    if secondary_targets:
        result["secondary"] = _extract_secondary_predictions(
            pred_df, secondary_targets, id_col, ts_col, quantile_levels
        )

    return result


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


def detect_regime(price_series: pd.Series, lookback: int = 20) -> str:
    """
    Detect market regime from recent price action.

    Returns "up", "down", or "sideways" based on the trend over the
    last `lookback` trading days.

    Thresholds:
      - Up: price rose >2% over lookback
      - Down: price fell >2% over lookback
      - Sideways: within ±2%
    """
    if len(price_series) < lookback:
        return "sideways"
    recent = price_series.iloc[-lookback:]
    pct_change = (recent.iloc[-1] - recent.iloc[0]) / recent.iloc[0]
    if pct_change > 0.02:
        return "up"
    elif pct_change < -0.02:
        return "down"
    else:
        return "sideways"


def _compute_static_bias() -> dict:
    """Compute per-horizon bias as simple mean of signed errors (original method)."""
    diag_path = RESULTS_DIR / "backtest_metrics.csv"
    if not diag_path.exists():
        return {}
    try:
        df = pd.read_csv(diag_path)
        if "signed_error" not in df.columns:
            if "predicted" in df.columns and "actual" in df.columns:
                df["signed_error"] = df["predicted"] - df["actual"]
            else:
                return {}
        bias = df.groupby("horizon")["signed_error"].mean().to_dict()
        return {int(k): v for k, v in bias.items()}
    except Exception:
        return {}


def compute_ewma_bias(alpha: float = 0.3) -> dict:
    """
    Compute per-horizon EWMA bias from backtest signed errors.

    Orders errors chronologically by forecast origin date and applies
    exponential weighting so recent errors have more influence.

    Args:
        alpha: Smoothing factor (0-1). Higher = more reactive to recent errors.
               Default 0.3 gives ~95% weight to last ~10 observations.

    Returns:
        {horizon: ewma_bias_value} or empty dict if unavailable.
    """
    diag_path = RESULTS_DIR / "backtest_metrics.csv"
    if not diag_path.exists():
        return {}
    try:
        df = pd.read_csv(diag_path, parse_dates=["as_of"])
        if "signed_error" not in df.columns:
            if "predicted" in df.columns and "actual" in df.columns:
                df["signed_error"] = df["predicted"] - df["actual"]
            else:
                return {}

        bias = {}
        for h in df["horizon"].unique():
            sub = df[df["horizon"] == h].sort_values("as_of")
            errors = sub["signed_error"].values
            if len(errors) == 0:
                continue
            ewma = errors[0]
            for e in errors[1:]:
                ewma = alpha * e + (1.0 - alpha) * ewma
            bias[int(h)] = float(ewma)
        return bias
    except Exception:
        return {}


def compute_regime_ewma_bias(alpha: float = 0.3) -> dict:
    """
    Compute per-horizon, per-regime EWMA bias from backtest signed errors.

    Returns nested dict: {horizon: {"up": bias, "down": bias, "sideways": bias, "global": bias}}
    The "global" key is the standard (non-regime) EWMA fallback.
    """
    diag_path = RESULTS_DIR / "backtest_metrics.csv"
    if not diag_path.exists():
        return {}
    try:
        df = pd.read_csv(diag_path, parse_dates=["as_of"])
        if "signed_error" not in df.columns:
            if "predicted" in df.columns and "actual" in df.columns:
                df["signed_error"] = df["predicted"] - df["actual"]
            else:
                return {}
        if "regime" not in df.columns:
            # No regime data — fall back to global EWMA
            global_bias = compute_ewma_bias(alpha)
            return {h: {"global": v, "up": v, "down": v, "sideways": v}
                    for h, v in global_bias.items()}

        result = {}
        for h in df["horizon"].unique():
            h_int = int(h)
            sub = df[df["horizon"] == h].sort_values("as_of")

            # Global EWMA (all regimes)
            errors = sub["signed_error"].values
            ewma = errors[0]
            for e in errors[1:]:
                ewma = alpha * e + (1.0 - alpha) * ewma
            global_ewma = float(ewma)

            # Per-regime EWMA
            regime_bias = {"global": global_ewma}
            for regime in ["up", "down", "sideways"]:
                rsub = sub[sub["regime"] == regime].sort_values("as_of")
                if len(rsub) >= 2:
                    errors_r = rsub["signed_error"].values
                    ewma_r = errors_r[0]
                    for e in errors_r[1:]:
                        ewma_r = alpha * e + (1.0 - alpha) * ewma_r
                    regime_bias[regime] = float(ewma_r)
                else:
                    regime_bias[regime] = global_ewma  # fallback

            result[h_int] = regime_bias
        return result
    except Exception:
        return {}


def load_bias_estimates(method: str = "regime_ewma", alpha: float = 0.3,
                        regime: str = None) -> dict:
    """
    Load per-horizon bias estimates from backtest diagnostics.

    Args:
        method: "regime_ewma", "ewma", or "static".
        alpha: EWMA smoothing factor.
        regime: Current market regime ("up"/"down"/"sideways"). Used with regime_ewma.

    Returns:
        {horizon: bias_value} or empty dict if unavailable.
    """
    if method == "regime_ewma":
        regime_bias = compute_regime_ewma_bias(alpha)
        if not regime_bias:
            return {}
        r = regime or "global"
        return {h: biases.get(r, biases["global"]) for h, biases in regime_bias.items()}
    elif method == "ewma":
        return compute_ewma_bias(alpha)
    else:
        return _compute_static_bias()


def optimize_ensemble_weights(horizons: list = None) -> dict:
    """
    Learn optimal per-horizon ensemble weights from backtest data.

    Requires backtest_metrics.csv to have 'c2_median' and 'bolt_median' columns
    (produced when backtest runs with return_components=True).

    Uses grid search to minimize MAE for each horizon.

    Returns:
        {horizon: {"weight_chronos2": w, "weight_bolt": 1-w, "mae_improvement": pct}}
        or empty dict if component data unavailable.
    """
    diag_path = RESULTS_DIR / "backtest_metrics.csv"
    if not diag_path.exists():
        return {}
    try:
        df = pd.read_csv(diag_path)
        if "c2_median" not in df.columns or "bolt_median" not in df.columns:
            return {}

        if horizons is None:
            horizons = sorted(df["horizon"].unique())

        result = {}
        for h in horizons:
            sub = df[df["horizon"] == h]
            if len(sub) < 3:
                continue

            actual = sub["actual"].values
            c2 = sub["c2_median"].values
            bolt = sub["bolt_median"].values

            # Grid search over weights 0.0 to 1.0
            best_w, best_mae = 0.6, float("inf")
            for w_int in range(0, 101, 5):  # 0%, 5%, 10%, ..., 100%
                w = w_int / 100.0
                ensemble = w * c2 + (1 - w) * bolt
                mae = np.mean(np.abs(ensemble - actual))
                if mae < best_mae:
                    best_mae = mae
                    best_w = w

            # Compare to default 60/40
            default_mae = np.mean(np.abs(0.6 * c2 + 0.4 * bolt - actual))
            improvement = (default_mae - best_mae) / default_mae * 100 if default_mae > 0 else 0

            result[int(h)] = {
                "weight_chronos2": round(best_w, 2),
                "weight_bolt": round(1 - best_w, 2),
                "optimized_mae": round(float(best_mae), 3),
                "default_mae": round(float(default_mae), 3),
                "improvement_pct": round(float(improvement), 1),
            }
        return result
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
