"""
Feature engineering pipeline — builds enriched feature matrix from raw data.

Feature groups:
  1. Price & Term Structure (CT1 + synthetic deferred contracts)
  2. Technical Indicators (RSI, Bollinger, MACD, realized vol)
  3. Macro (DXY, WTI crude)
  4. CFTC COT Positioning (commercial/managed money net, ratios, momentum)
  5. Weather & Drought (temp, precip, ET0, GDD, PDSI, stress indices)
  6. Seasonality (Fourier, crop calendar, WASDE flags)
  7. ICE Certified Stocks (warehouse inventory, supply pressure)
"""
import pandas as pd
import numpy as np
from pathlib import Path

RAW_DIR = Path(__file__).resolve().parent.parent / "data" / "raw"
FEAT_DIR = Path(__file__).resolve().parent.parent / "data" / "features"
FEAT_DIR.mkdir(parents=True, exist_ok=True)


def _smooth_seasonal_flag(doy_series, start_doy, end_doy, ramp_days=15):
    """Smooth seasonal flag using sigmoid transitions at boundaries.

    Instead of hard 0→1 switches, ramps up/down over ~ramp_days using a sigmoid.
    This avoids z-score spikes when the flag is normalized.
    """
    k = 4.0 / ramp_days  # steepness: 4/ramp gives ~95% transition within ramp_days
    rise = 1.0 / (1.0 + np.exp(-k * (doy_series - start_doy)))
    fall = 1.0 / (1.0 + np.exp(-k * (end_doy - doy_series)))
    return rise * fall

# Covariate metadata: classifies each feature for Chronos-2 multivariate input
COVARIATE_META = {
    # Past-only covariates (historical values only, no future projection)
    "dxy": "past",
    "wti_crude": "past",
    "traders_noncomm_long": "past",      # CFTC non-commercial long (r=0.71)
    "spec_net_pct": "past",              # Speculator net % of OI (r=0.67)
    "conc_4_short": "past",              # Top-4 short concentration (r=0.63)
    "pdsi_severe_drought": "past",       # Binary severe drought flag (r=0.42)
    "realised_vol_21d": "past",
    "noaa_pdsi": "past",
    "ct1_ret_5d": "past",               # 5-day price momentum
    "ct1_ret_21d": "past",              # 21-day price momentum
    "dxy_5d_ret": "past",               # DXY 5-day return
    "wti_5d_ret": "past",               # WTI 5-day return
    "noncomm_long_chg_5d": "past",      # CFTC positioning change
    "spec_net_pct_chg_5d": "past",      # Spec positioning change
    "cert_stocks_z": "past",            # ICE certified stocks z-score (supply pressure)
    "cert_stocks_chg_5d": "past",       # 5-day change in certified stocks
    "cert_stocks_chg_21d": "past",      # 21-day change in certified stocks (trend)
    # Known future covariates (deterministic, can be computed for future dates)
    "seas_sin_annual": "future",
    "seas_cos_annual": "future",
    "flag_planting": "future",
    "flag_boll_dev": "future",
    "flag_harvest": "future",
    "flag_wasde": "future",
}


def get_covariate_columns():
    """Return (past_covariates, known_future_covariates) column lists."""
    past = [k for k, v in COVARIATE_META.items() if v == "past"]
    future = [k for k, v in COVARIATE_META.items() if v == "future"]
    return past, future


def build_features() -> pd.DataFrame:
    print("[features] Building enriched feature matrix …")

    # ── Load raw data ──
    ct1 = pd.read_parquet(RAW_DIR / "ct1.parquet")
    macro = pd.read_parquet(RAW_DIR / "macro.parquet")
    cot = pd.read_parquet(RAW_DIR / "cftc_cot.parquet")
    drought = pd.read_parquet(RAW_DIR / "drought.parquet")
    weather = pd.read_parquet(RAW_DIR / "weather.parquet")

    # ── 1. Price & Term Structure ──
    df = ct1[["ct1_close"]].copy()
    df["ct1_volume"] = ct1["ct1_volume"] if "ct1_volume" in ct1.columns else np.nan

    # Synthetic deferred contracts (contango proxies)
    df["ct2_close"] = df["ct1_close"] * (1 + 0.002 + 0.001 * np.sin(np.arange(len(df)) * 2 * np.pi / 252))
    df["ct3_close"] = df["ct1_close"] * (1 + 0.004 + 0.002 * np.sin(np.arange(len(df)) * 2 * np.pi / 252))
    df["ct5_close"] = df["ct1_close"] * (1 + 0.008 + 0.003 * np.sin(np.arange(len(df)) * 2 * np.pi / 252))

    # Term structure signals
    df["roll_yield"] = (df["ct2_close"] - df["ct1_close"]) / df["ct1_close"]
    df["mid_spread"] = (df["ct3_close"] - df["ct1_close"]) / df["ct1_close"]
    df["term_spread"] = (df["ct5_close"] - df["ct1_close"]) / df["ct1_close"]
    df["curve_curvature"] = df["ct3_close"] - 0.5 * (df["ct1_close"] + df["ct5_close"])
    df["roll_yield_5d_chg"] = df["roll_yield"].diff(5)

    # ── 2. Technical Indicators ──
    close = df["ct1_close"]
    df["log_ret_1d"] = np.log(close / close.shift(1))
    df["log_ret_5d"] = np.log(close / close.shift(5))
    df["log_ret_20d"] = np.log(close / close.shift(20))
    df["realised_vol_21d"] = df["log_ret_1d"].rolling(21).std() * np.sqrt(252)

    # RSI-14
    delta = close.diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = (-delta.clip(upper=0)).rolling(14).mean()
    rs = gain / loss.replace(0, np.nan)
    df["rsi_14"] = 100 - (100 / (1 + rs))

    # Bollinger %B
    sma20 = close.rolling(20).mean()
    std20 = close.rolling(20).std()
    df["bollinger_pctb"] = (close - (sma20 - 2 * std20)) / (4 * std20)

    # MACD signal
    ema12 = close.ewm(span=12).mean()
    ema26 = close.ewm(span=26).mean()
    macd_line = ema12 - ema26
    df["macd_signal"] = macd_line - macd_line.ewm(span=9).mean()

    # Open interest
    df["open_interest"] = df["ct1_volume"].rolling(5).mean()  # proxy
    df["oi_pct_change"] = df["open_interest"].pct_change(5)

    # Price momentum features (for use as covariates)
    df["ct1_ret_5d"] = close.pct_change(5)
    df["ct1_ret_10d"] = close.pct_change(10)
    df["ct1_ret_21d"] = close.pct_change(21)

    # ── 3. Macro ──
    df = df.join(macro[["dxy", "wti_crude"]], how="left")
    df["dxy"] = df["dxy"].ffill()
    df["wti_crude"] = df["wti_crude"].ffill()
    df["dxy_dev"] = df["dxy"] / df["dxy"].rolling(63).mean() - 1
    df["crude_5d_ret"] = np.log(df["wti_crude"] / df["wti_crude"].shift(5))
    df["dxy_5d_ret"] = df["dxy"].pct_change(5)
    df["wti_5d_ret"] = df["wti_crude"].pct_change(5)

    # ── 4. CFTC COT Positioning ──
    cot_daily = _resample_cot_to_daily(cot, df.index)
    df = df.join(cot_daily, how="left")

    # Week-over-week changes in key positioning metrics
    if "traders_noncomm_long" in df.columns:
        df["noncomm_long_chg_5d"] = df["traders_noncomm_long"].diff(5)
    if "spec_net_pct" in df.columns:
        df["spec_net_pct_chg_5d"] = df["spec_net_pct"].diff(5)

    # ── 5. Weather & Drought ──
    weather_cols = ["temp_max", "temp_min", "precip_sum", "et0"]
    for col in weather_cols:
        if col in weather.columns:
            df[col] = weather[col].reindex(df.index).ffill()

    df["temp_range"] = df["temp_max"] - df["temp_min"]
    df["heat_stress"] = (df["temp_max"] > 95).astype(float)
    df["heat_stress_5d"] = df["heat_stress"].rolling(5).sum()
    df["precip_7d"] = df["precip_sum"].rolling(7).sum()
    df["precip_30d"] = df["precip_sum"].rolling(30).sum()
    df["precip_deficit_30d"] = df["precip_30d"] - df["et0"].rolling(30).sum()

    # GDD (Growing Degree Days, base 60°F for cotton)
    df["gdd"] = ((df["temp_max"] + df["temp_min"]) / 2 - 60).clip(lower=0)
    df["gdd_cumulative_30d"] = df["gdd"].rolling(30).sum()

    # Drought
    drought_daily = drought.reindex(df.index, method="ffill")
    df["noaa_pdsi"] = drought_daily["noaa_pdsi"]
    df["pdsi_severe_drought"] = (df["noaa_pdsi"] < -3).astype(float)
    df["pdsi_30d_chg"] = df["noaa_pdsi"].diff(20)
    df["drought_proxy"] = (df["et0"].rolling(30).mean() / df["precip_sum"].rolling(30).mean().replace(0, np.nan)).clip(upper=20)

    # Water stress
    df["water_stress_7d"] = (df["et0"].rolling(7).sum() - df["precip_7d"]).clip(lower=0)
    df["water_stress_30d"] = (df["et0"].rolling(30).sum() - df["precip_30d"]).clip(lower=0)

    # ── 6. ICE Certified Stocks ──
    # Try CSV first (real data), then parquet (cached)
    cert_csv_path = RAW_DIR / "ice_certified_stocks.csv"
    cert_pq_path = RAW_DIR / "certified_stocks.parquet"
    cert = None
    if cert_csv_path.exists():
        cert = pd.read_csv(cert_csv_path, parse_dates=["date"]).set_index("date")
    elif cert_pq_path.exists():
        cert = pd.read_parquet(cert_pq_path)

    if cert is not None and "certified_stocks" in cert.columns:
        df["certified_stocks"] = cert["certified_stocks"].reindex(df.index, method="ffill")

        # For dates before the first real data point, fill with 0 (neutral)
        # so dropna() doesn't remove the entire 2018-2023 history
        cert_cols_to_fill = []

        # Z-score: how far current stocks are from 6-month average
        cs = df["certified_stocks"]
        cs_mean = cs.rolling(126, min_periods=21).mean()
        cs_std = cs.rolling(126, min_periods=21).std().replace(0, np.nan)
        df["cert_stocks_z"] = (cs - cs_mean) / cs_std

        # Absolute change over 5 and 21 trading days
        df["cert_stocks_chg_5d"] = cs.diff(5)
        df["cert_stocks_chg_21d"] = cs.diff(21)

        # Percent change for relative moves
        df["cert_stocks_pct_5d"] = cs.pct_change(5)
        df["cert_stocks_pct_21d"] = cs.pct_change(21)

        # Supply pressure signal: falling stocks = bullish (positive signal)
        df["cert_stocks_supply_pressure"] = -df["cert_stocks_z"]  # Inverted: low stocks = high pressure

        # Seasonal deviation: stocks vs seasonal norm
        month_mean = df.groupby(df.index.month)["certified_stocks"].transform("mean")
        month_std = df.groupby(df.index.month)["certified_stocks"].transform("std").replace(0, np.nan)
        df["cert_stocks_seasonal_dev"] = (cs - month_mean) / month_std

        # Fill NaN in cert stocks features with 0 (neutral) for pre-data period
        # This prevents dropna() from removing 2018-2023 rows
        cert_feature_cols = [c for c in df.columns if c.startswith("cert_stocks")]
        cert_feature_cols.append("certified_stocks")
        for col in cert_feature_cols:
            df[col] = df[col].fillna(0)

        non_zero = (df["certified_stocks"] > 0).sum()
        print(f"  → ICE certified stocks: {cs.dropna().min():.0f} – {cs.dropna().max():.0f} bales")
        print(f"  → Real data coverage: {non_zero} of {len(df)} trading days ({non_zero/len(df)*100:.0f}%)")
    else:
        print("  → WARN: No certified stocks data found — run ingest first")

    # ── 7. Seasonality ──
    doy = df.index.dayofyear
    df["seas_sin_annual"] = np.sin(2 * np.pi * doy / 365.25)
    df["seas_cos_annual"] = np.cos(2 * np.pi * doy / 365.25)
    df["seas_sin_semi"] = np.sin(4 * np.pi * doy / 365.25)
    df["seas_cos_semi"] = np.cos(4 * np.pi * doy / 365.25)

    # Crop calendar flags — smooth sigmoid transitions to avoid z-score spikes
    doy = df.index.dayofyear
    df["flag_planting"] = _smooth_seasonal_flag(doy, 91, 181)   # Apr 1 – Jun 30
    df["flag_boll_dev"] = _smooth_seasonal_flag(doy, 182, 243)  # Jul 1 – Aug 31
    df["flag_harvest"] = _smooth_seasonal_flag(doy, 244, 334)   # Sep 1 – Nov 30

    # WASDE release approximation (around 10th–12th of each month)
    dom = df.index.day
    df["flag_wasde"] = ((dom >= 9) & (dom <= 13)).astype(float)

    # ── Finalize ──
    df = df.dropna(subset=["ct1_close"])
    df = df.ffill().dropna()

    df.to_parquet(FEAT_DIR / "features.parquet")
    print(f"  → {len(df)} rows × {len(df.columns)} features")
    print(f"  → Date range: {df.index.min().date()} – {df.index.max().date()}")
    print(f"  → Features: {list(df.columns)}")
    return df


def _resample_cot_to_daily(cot: pd.DataFrame, daily_index: pd.DatetimeIndex) -> pd.DataFrame:
    """Resample weekly COT data to daily and derive positioning features."""
    result = pd.DataFrame(index=daily_index)

    for col in cot.columns:
        result[col] = cot[col].reindex(daily_index, method="ffill")

    if "comm_long" in result.columns and "comm_short" in result.columns:
        result["comm_net"] = result["comm_long"] - result["comm_short"]
        result["comm_net_pct"] = result["comm_net"] / (result["comm_long"] + result["comm_short"]).replace(0, np.nan)
        result["comm_net_5w_chg"] = result["comm_net"].diff(25)
        result["comm_net_z"] = (result["comm_net"] - result["comm_net"].rolling(126).mean()) / result["comm_net"].rolling(126).std().replace(0, np.nan)

    if "noncomm_long" in result.columns and "noncomm_short" in result.columns:
        result["spec_net"] = result["noncomm_long"] - result["noncomm_short"]
        result["spec_net_pct"] = result["spec_net"] / (result["noncomm_long"] + result["noncomm_short"]).replace(0, np.nan)
        result["spec_net_5w_chg"] = result["spec_net"].diff(25)
        result["spec_net_z"] = (result["spec_net"] - result["spec_net"].rolling(126).mean()) / result["spec_net"].rolling(126).std().replace(0, np.nan)

    if "mm_long" in result.columns and "mm_short" in result.columns:
        result["mm_net"] = result["mm_long"] - result["mm_short"]
        result["mm_net_pct"] = result["mm_net"] / (result["mm_long"] + result["mm_short"]).replace(0, np.nan)
        result["mm_net_5w_chg"] = result["mm_net"].diff(25)
        result["mm_net_z"] = (result["mm_net"] - result["mm_net"].rolling(126).mean()) / result["mm_net"].rolling(126).std().replace(0, np.nan)

    if "cot_oi" in result.columns:
        result["cot_oi_z"] = (result["cot_oi"] - result["cot_oi"].rolling(126).mean()) / result["cot_oi"].rolling(126).std().replace(0, np.nan)

    return result


if __name__ == "__main__":
    build_features()
