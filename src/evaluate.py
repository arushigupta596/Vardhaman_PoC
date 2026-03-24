"""
Walk-forward backtesting engine and evaluation metrics.

Supports multivariate Chronos-2 forecasting with covariate leakage prevention.
"""
import numpy as np
import pandas as pd
from pathlib import Path
from scipy import stats

RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"
RESULTS_DIR.mkdir(parents=True, exist_ok=True)


def crps_quantile(q10, q50, q90, actual):
    """CRPS approximation from 3 quantiles using pinball loss."""
    def pinball(q, y, tau):
        e = y - q
        return np.where(e >= 0, tau * e, (tau - 1) * e)
    return (pinball(q10, actual, 0.1) + pinball(q50, actual, 0.5) + pinball(q90, actual, 0.9)) / 3


def evaluate_forecast(forecast_dict, actual_series, horizon):
    """
    Evaluate a single forecast against actuals.

    forecast_dict: output from model.forecast() with keys q10, median, q90
    actual_series: pd.Series of actual prices following the forecast origin
    horizon: int, the target horizon
    """
    actual_vals = actual_series.values[:horizon]
    if len(actual_vals) < horizon:
        return None

    actual_h = actual_vals[horizon - 1]
    pred_h = forecast_dict["median"][horizon - 1]
    q10_h = forecast_dict["q10"][horizon - 1]
    q90_h = forecast_dict["q90"][horizon - 1]

    # Point metrics
    mae = abs(pred_h - actual_h)
    rmse = np.sqrt((pred_h - actual_h) ** 2)

    # CRPS
    crps = crps_quantile(q10_h, pred_h, q90_h, actual_h)

    # Directional accuracy (vs origin price)
    origin_price = actual_series.iloc[0] if len(actual_series) > 0 else pred_h
    dir_actual = 1 if actual_h > origin_price else 0
    dir_pred = 1 if pred_h > origin_price else 0
    dir_acc = float(dir_actual == dir_pred)

    # Coverage (80% interval)
    coverage = float(q10_h <= actual_h <= q90_h)

    # Interval width
    interval_width = q90_h - q10_h

    # MASE (using naive seasonal baseline: persistence forecast)
    naive_error = abs(actual_h - origin_price)
    mase = mae / naive_error if naive_error > 1e-10 else np.nan

    return {
        "mae": mae,
        "rmse": rmse,
        "crps": crps,
        "dir_acc": dir_acc,
        "coverage": coverage,
        "interval_width": interval_width,
        "mase": mase,
        "actual": actual_h,
        "predicted": pred_h,
        "q10": q10_h,
        "q90": q90_h,
        "origin_price": origin_price,
    }


def walk_forward_backtest(
    features_df: pd.DataFrame,
    horizons: list = None,
    test_days: int = 500,
    step: int = 20,
    model_fn=None,
    config: dict = None,
    output_file: str = "backtest_metrics.csv",
):
    """
    Walk-forward backtest with multivariate Chronos-2.

    Args:
        features_df: Full features DataFrame with all covariates
        horizons: List of forecast horizons (default from config)
        test_days: Number of days in test period
        step: Step size between forecast origins
        model_fn: callable(features_df, origin_idx, horizon) -> forecast_dict
                  If None, uses src.model.forecast_at_origin
        config: Configuration dict
        output_file: Filename for results CSV
    """
    from src.covariates import load_config

    if config is None:
        config = load_config()
    if horizons is None:
        horizons = config["forecast"]["horizons"]

    if model_fn is None:
        from src.model import forecast_at_origin
        model_fn = lambda df, idx, h: forecast_at_origin(df, idx, h, config=config)

    series = features_df[config["data"]["target"]]
    n = len(features_df)
    test_start = n - test_days
    min_context = config.get("backtest", {}).get("min_context_length", 0)
    warmup_origins = config.get("backtest", {}).get("warmup_origins", 0)
    all_origins = [idx for idx in range(test_start, n - max(horizons), step) if idx >= min_context]
    origins = all_origins[warmup_origins:]  # skip first N origins (regime-shift warm-up)

    print(f"[backtest] {len(origins)} forecast origins, horizons={horizons}")
    if min_context > 0:
        print(f"  → Min context guard: {min_context} days (skipped {len(list(range(test_start, n - max(horizons), step))) - len(all_origins)} origins)")
    if warmup_origins > 0:
        print(f"  → Warm-up skip: {warmup_origins} origins skipped")
    print(f"  → Test period: {features_df.index[origins[0]].date() if origins else 'N/A'} – {features_df.index[-1].date()}")

    records = []
    for i, origin_idx in enumerate(origins):
        origin_date = features_df.index[origin_idx]

        for h in horizons:
            if origin_idx + h >= n:
                continue

            try:
                # Call model with full features + origin index (strict no-leakage)
                forecast_dict = model_fn(features_df, origin_idx, h)
                actual_series = series.iloc[origin_idx:origin_idx + h + 1]
                metrics = evaluate_forecast(forecast_dict, actual_series, h)
                if metrics:
                    metrics["as_of"] = origin_date
                    metrics["horizon"] = h
                    records.append(metrics)
            except Exception as e:
                print(f"  → Error at {origin_date.date()} h={h}: {e}")

        if (i + 1) % 5 == 0:
            print(f"  → {i+1}/{len(origins)} complete")

    results_df = pd.DataFrame(records)

    if len(results_df) == 0:
        print(f"\n[backtest] Done. 0 forecasts evaluated — check model errors above.")
        results_df.to_csv(RESULTS_DIR / output_file, index=False)
        return results_df

    # Add signed error for bias analysis
    results_df["signed_error"] = results_df["predicted"] - results_df["actual"]

    results_df.to_csv(RESULTS_DIR / output_file, index=False)

    # Save per-origin diagnostics
    diag_file = output_file.replace(".csv", "_diagnostics.csv")
    diag = results_df.groupby(["as_of", "horizon"]).agg(
        mae=("mae", "mean"),
        signed_error=("signed_error", "mean"),
        coverage=("coverage", "mean"),
    ).reset_index()
    diag.to_csv(RESULTS_DIR / diag_file, index=False)
    print(f"  → Per-origin diagnostics saved to {diag_file}")

    print(f"\n[backtest] Done. {len(results_df)} forecasts evaluated.")

    # Flag outlier origins (MAE > 3× median for any horizon)
    outlier_origins = set()
    for h in horizons:
        sub = results_df[results_df["horizon"] == h]
        if len(sub) > 2:
            median_mae = sub["mae"].median()
            threshold = 3 * median_mae
            outliers = sub[sub["mae"] > threshold]
            for _, row in outliers.iterrows():
                outlier_origins.add(str(row["as_of"].date()) if hasattr(row["as_of"], "date") else str(row["as_of"]))

    if outlier_origins:
        print(f"\n  ⚠ Outlier origins (MAE > 3× median): {', '.join(sorted(outlier_origins))}")

    # Summary with bias reporting
    for h in horizons:
        sub = results_df[results_df["horizon"] == h]
        if len(sub) > 0:
            bias = sub["signed_error"].mean()
            print(f"\n  {h}d: MAE={sub['mae'].mean():.2f}  CRPS={sub['crps'].mean():.3f}  "
                  f"Dir={sub['dir_acc'].mean():.1%}  Cov={sub['coverage'].mean():.1%}  "
                  f"RMSE={sub['rmse'].mean():.2f}  MASE={sub['mase'].mean():.2f}  "
                  f"Bias={bias:+.2f}")

    return results_df


def diebold_mariano_test(errors_a, errors_b, h=1):
    """
    Diebold-Mariano test for equal predictive accuracy.
    Negative t-stat means model A is better.
    """
    d = errors_a ** 2 - errors_b ** 2
    n = len(d)
    mean_d = d.mean()

    # HAC variance (Newey-West with bandwidth h-1)
    gamma_0 = np.var(d, ddof=1)
    hac_var = gamma_0
    for k in range(1, min(h, n - 1)):
        gamma_k = np.cov(d[k:], d[:-k])[0, 1]
        hac_var += 2 * (1 - k / h) * gamma_k

    se = np.sqrt(hac_var / n)
    if se < 1e-10:
        return 0.0, 1.0

    t_stat = mean_d / se
    p_val = 2 * stats.t.sf(abs(t_stat), df=n - 1)
    return t_stat, p_val


if __name__ == "__main__":
    features = pd.read_parquet("data/features/features.parquet")
    walk_forward_backtest(features)
