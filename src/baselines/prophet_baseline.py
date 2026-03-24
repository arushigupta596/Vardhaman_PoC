"""
Prophet baseline with enriched regressors including CFTC COT data.
"""
import warnings
import numpy as np
import pandas as pd
from pathlib import Path
from prophet import Prophet

warnings.filterwarnings("ignore")

RESULTS_DIR = Path(__file__).resolve().parent.parent.parent / "results"


def prophet_forecast(
    features_df: pd.DataFrame,
    origin_idx: int,
    horizon: int,
) -> dict:
    """
    Fit Prophet on features up to origin_idx and forecast horizon steps.
    Returns dict with keys: median, q10, q90.
    """
    train = features_df.iloc[max(0, origin_idx - 512):origin_idx].copy()

    # Build Prophet dataframe
    pdf = pd.DataFrame({
        "ds": train.index,
        "y": train["ct1_close"].values,
    })

    # Regressors available in features
    regressor_cols = []
    candidate_regressors = [
        "dxy_dev", "roll_yield", "crude_5d_ret", "noaa_pdsi",
        "water_stress_30d", "term_spread",
        # NEW: COT positioning signals
        "comm_net_pct", "spec_net_pct", "mm_net_pct", "cot_oi_z",
        "comm_net_z", "spec_net_z",
    ]
    for col in candidate_regressors:
        if col in train.columns and train[col].notna().sum() > 100:
            pdf[col] = train[col].values
            regressor_cols.append(col)

    m = Prophet(
        growth="flat",
        yearly_seasonality=False,
        weekly_seasonality=True,
        daily_seasonality=False,
        changepoint_prior_scale=0.05,
        interval_width=0.80,
    )

    # Annual + semi-annual Fourier
    m.add_seasonality(name="annual", period=365.25, fourier_order=8)
    m.add_seasonality(name="semi_annual", period=365.25 / 2, fourier_order=4)

    # WASDE holidays
    wasde_dates = pd.to_datetime([
        f"{y}-{m:02d}-{d}" for y in range(2018, 2027)
        for m, d in [(1, 12), (2, 8), (3, 8), (4, 11), (5, 10), (6, 12),
                     (7, 12), (8, 12), (9, 12), (10, 11), (11, 9), (12, 10)]
    ])
    m.add_country_holidays(country_name="US")

    for col in regressor_cols:
        m.add_regressor(col, prior_scale=10.0)

    m.fit(pdf)

    # Future dataframe with flat-forward regressors
    future = m.make_future_dataframe(periods=horizon, freq="B")
    for col in regressor_cols:
        last_val = pdf[col].iloc[-1] if col in pdf.columns else 0
        future[col] = last_val
        future.loc[future["ds"].isin(pdf["ds"]), col] = pdf[col].values if len(pdf) == len(future.loc[future["ds"].isin(pdf["ds"])]) else last_val

    # Simple flat-forward fill
    for col in regressor_cols:
        if col in train.columns:
            vals = train[col].reindex(pd.to_datetime(future["ds"]), method="ffill").values
            future[col] = vals if len(vals) == len(future) else future[col].fillna(train[col].iloc[-1])

    forecast = m.predict(future)
    fcast = forecast.iloc[-horizon:]

    return {
        "median": fcast["yhat"].values,
        "q10": fcast["yhat_lower"].values,
        "q90": fcast["yhat_upper"].values,
    }


def prophet_backtest(
    features_df: pd.DataFrame,
    horizons: list = [30, 60, 90],
    test_days: int = 500,
    step: int = 20,
) -> pd.DataFrame:
    """Walk-forward backtest for Prophet."""
    from src.evaluate import evaluate_forecast

    n = len(features_df)
    test_start = n - test_days
    origins = list(range(test_start, n - max(horizons), step))
    series = features_df["ct1_close"]

    print(f"[prophet] {len(origins)} forecast origins, horizons={horizons}")

    records = []
    for i, origin_idx in enumerate(origins):
        origin_date = features_df.index[origin_idx]

        for h in horizons:
            if origin_idx + h >= n:
                continue
            try:
                fcast = prophet_forecast(features_df, origin_idx, h)
                actual_series = series.iloc[origin_idx:origin_idx + h + 1]
                metrics = evaluate_forecast(fcast, actual_series, h)
                if metrics:
                    metrics["as_of"] = origin_date
                    metrics["horizon"] = h
                    records.append(metrics)
            except Exception as e:
                print(f"  → Prophet error at {origin_date.date()} h={h}: {e}")

        if (i + 1) % 5 == 0:
            print(f"  → {i+1}/{len(origins)} complete")

    results_df = pd.DataFrame(records)
    results_df.to_csv(RESULTS_DIR / "prophet_backtest.csv", index=False)
    print(f"\n[prophet] Done. {len(results_df)} forecasts evaluated.")

    for h in horizons:
        sub = results_df[results_df.horizon == h]
        if len(sub) > 0:
            print(f"  {h}d: MAE={sub.mae.mean():.2f}  Dir={sub.dir_acc.mean():.1%}  Cov={sub.coverage.mean():.1%}")

    return results_df


if __name__ == "__main__":
    features = pd.read_parquet("data/features/features.parquet")
    prophet_backtest(features)
