"""
Covariate preparation module for Chronos-2 multivariate forecasting.

Converts the wide-format features DataFrame into the long-format DataFrames
that Chronos-2's predict_df API expects:

  - context_df: [id, timestamp, target, <past_covariates>, <known_future_covariates>]
  - future_df:  [id, timestamp, <known_future_covariates>]  (no target column)
"""
import yaml
import numpy as np
import pandas as pd
from pathlib import Path
from src.features import _smooth_seasonal_flag

CONFIG_PATH = Path(__file__).resolve().parent.parent / "configs" / "config.yaml"


def load_config() -> dict:
    """Load configuration from config.yaml."""
    with open(CONFIG_PATH) as f:
        return yaml.safe_load(f)


def validate_covariates(features_df: pd.DataFrame, config: dict = None) -> dict:
    """
    Validate that all configured covariates exist in features_df.
    Returns dict with validation results.
    """
    if config is None:
        config = load_config()

    past_covs = config["data"]["past_covariates"]
    future_covs = config["data"]["known_future_covariates"]
    target = config["data"]["target"]

    results = {"valid": True, "missing": [], "high_nan": [], "available_past": [], "available_future": []}

    # Check target
    if target not in features_df.columns:
        results["valid"] = False
        results["missing"].append(target)
        return results

    # Check past covariates
    for col in past_covs:
        if col not in features_df.columns:
            results["missing"].append(col)
            results["valid"] = False
        else:
            nan_pct = features_df[col].isna().mean()
            if nan_pct > 0.10:
                results["high_nan"].append((col, f"{nan_pct:.1%}"))
            else:
                results["available_past"].append(col)

    # Check known future covariates
    for col in future_covs:
        if col not in features_df.columns:
            results["missing"].append(col)
            results["valid"] = False
        else:
            results["available_future"].append(col)

    return results


def _reindex_to_business_days(df: pd.DataFrame) -> pd.DataFrame:
    """
    Reindex DataFrame to a regular business-day frequency, forward-filling
    any market holiday gaps. This allows Chronos-2 to infer 'B' frequency.
    """
    if len(df) == 0:
        return df
    bdays = pd.bdate_range(df.index[0], df.index[-1], freq="B")
    reindexed = df.reindex(bdays, method="ffill")
    reindexed.index.name = df.index.name
    return reindexed


def _normalize_covariates(context: pd.DataFrame, cols: list) -> tuple:
    """
    Z-score normalize covariate columns using context window statistics.
    Returns (normalized context, stats dict for applying to future_df).
    """
    stats = {}
    context = context.copy()
    for col in cols:
        if col in context.columns:
            mean = context[col].mean()
            std = context[col].std()
            if std > 1e-10:
                context[col] = (context[col] - mean) / std
                stats[col] = {"mean": mean, "std": std}
            else:
                stats[col] = {"mean": mean, "std": 1.0}
    return context, stats


def build_context_df(
    features_df: pd.DataFrame,
    origin_idx: int,
    config: dict = None,
    normalize: bool = True,
) -> pd.DataFrame:
    """
    Build the context DataFrame for Chronos-2 predict_df.

    Converts wide-format features (up to origin_idx) into long-format:
      [id, timestamp, target, <past_covariates>, <known_future_covariates>]

    Args:
        features_df: Full features DataFrame with DatetimeIndex
        origin_idx: Forecast origin index (exclusive upper bound — no data at or after this index)
        config: Configuration dict (loaded from config.yaml if None)
        normalize: If True, z-score normalize covariates

    Returns:
        context_df suitable for Chronos2Pipeline.predict_df()
    """
    if config is None:
        config = load_config()

    target_col = config["data"]["target"]
    past_covs = config["data"]["past_covariates"]
    future_covs = config["data"]["known_future_covariates"]
    id_col = config["data"]["id_column"]
    ts_col = config["data"]["timestamp_column"]

    # Slice to before origin (strict no-leakage)
    context = features_df.iloc[:origin_idx]

    # Reindex to regular business-day frequency (fills holiday gaps)
    context = _reindex_to_business_days(context)

    # Normalize covariates if requested
    norm_stats = {}
    if normalize:
        context, norm_stats = _normalize_covariates(context, past_covs + future_covs)

    # Build long-format DataFrame
    data = {
        id_col: target_col,
        ts_col: context.index,
        "target": context[target_col].values,
    }

    # Add past covariates
    for col in past_covs:
        if col in context.columns:
            data[col] = context[col].values

    # Add known future covariates (historical values in context)
    for col in future_covs:
        if col in context.columns:
            data[col] = context[col].values

    context_df = pd.DataFrame(data)
    context_df._norm_stats = norm_stats  # attach for future_df normalization
    return context_df


def build_context_df_cross_learning(
    features_df: pd.DataFrame,
    origin_idx: int,
    config: dict = None,
    normalize: bool = True,
) -> pd.DataFrame:
    """
    Build context DataFrame with cross-learning: includes target series + additional
    related series (dxy, wti) as separate items.

    Chronos-2's group attention mechanism learns dependencies across items
    that share the same context window.
    """
    if config is None:
        config = load_config()

    target_col = config["data"]["target"]
    cross_series = config["data"].get("cross_learning_series", [])
    past_covs = config["data"]["past_covariates"]
    future_covs = config["data"]["known_future_covariates"]
    id_col = config["data"]["id_column"]
    ts_col = config["data"]["timestamp_column"]

    context = features_df.iloc[:origin_idx]

    # Reindex to regular business-day frequency (fills holiday gaps)
    context = _reindex_to_business_days(context)

    # Normalize covariates if requested (using training window stats)
    norm_stats = {}
    if normalize:
        context, norm_stats = _normalize_covariates(context, past_covs + future_covs)

    frames = []

    # Target series with all covariates
    target_data = {
        id_col: target_col,
        ts_col: context.index,
        "target": context[target_col].values,
    }
    for col in past_covs:
        if col in context.columns:
            target_data[col] = context[col].values
    for col in future_covs:
        if col in context.columns:
            target_data[col] = context[col].values
    frames.append(pd.DataFrame(target_data))

    # Secondary targets — each as a separate forecasting target with shared covariates
    secondary_targets = config["data"].get("secondary_targets", [])
    for sec_col in secondary_targets:
        if sec_col in context.columns and sec_col != target_col:
            sec_data = {
                id_col: sec_col,
                ts_col: context.index,
                "target": context[sec_col].values,
            }
            # Secondary targets get shared future covariates (like cross-learning)
            for col in future_covs:
                if col in context.columns:
                    sec_data[col] = context[col].values
            frames.append(pd.DataFrame(sec_data))

    # Cross-learning series — each as a separate item
    for series_col in cross_series:
        if series_col in context.columns and series_col != target_col:
            cl_data = {
                id_col: series_col,
                ts_col: context.index,
                "target": context[series_col].values,
            }
            # Add shared known future covariates
            for col in future_covs:
                if col in context.columns:
                    cl_data[col] = context[col].values
            frames.append(pd.DataFrame(cl_data))

    result = pd.concat(frames, ignore_index=True)
    result._norm_stats = norm_stats  # attach for future_df normalization
    return result


def build_future_df(
    last_date: pd.Timestamp,
    horizon: int,
    config: dict = None,
    covariate_stats: dict = None,
) -> pd.DataFrame:
    """
    Generate future covariate DataFrame with deterministic values.

    Only known_future_covariates are included (seasonality + calendar flags).
    Past-only covariates are NOT projected into the future.
    """
    if config is None:
        config = load_config()

    target_col = config["data"]["target"]
    future_covs = config["data"]["known_future_covariates"]
    cross_series = config["data"].get("cross_learning_series", [])
    id_col = config["data"]["id_column"]
    ts_col = config["data"]["timestamp_column"]

    future_dates = pd.bdate_range(last_date + pd.Timedelta(days=1), periods=horizon)
    doy = future_dates.dayofyear
    month = future_dates.month
    dom = future_dates.day

    # Compute deterministic future values
    future_values = {}
    for col in future_covs:
        if col == "seas_sin_annual":
            future_values[col] = np.sin(2 * np.pi * doy / 365.25)
        elif col == "seas_cos_annual":
            future_values[col] = np.cos(2 * np.pi * doy / 365.25)
        elif col == "flag_planting":
            future_values[col] = _smooth_seasonal_flag(doy, 91, 181)
        elif col == "flag_boll_dev":
            future_values[col] = _smooth_seasonal_flag(doy, 182, 243)
        elif col == "flag_harvest":
            future_values[col] = _smooth_seasonal_flag(doy, 244, 334)
        elif col == "flag_wasde":
            future_values[col] = ((dom >= 9) & (dom <= 13)).astype(float)

    # Normalize future covariate values if stats provided
    if covariate_stats:
        for col in list(future_values.keys()):
            if col in covariate_stats:
                s = covariate_stats[col]
                future_values[col] = (future_values[col] - s["mean"]) / s["std"]

    # Build future_df for target series
    frames = []
    target_future = {id_col: target_col, ts_col: future_dates, **future_values}
    frames.append(pd.DataFrame(target_future))

    # Build future_df for secondary targets
    secondary_targets = config["data"].get("secondary_targets", [])
    for sec_col in secondary_targets:
        if sec_col != target_col:
            sec_future = {id_col: sec_col, ts_col: future_dates, **future_values}
            frames.append(pd.DataFrame(sec_future))

    # Build future_df for cross-learning series
    for series_col in cross_series:
        if series_col != target_col:
            cl_future = {id_col: series_col, ts_col: future_dates, **future_values}
            frames.append(pd.DataFrame(cl_future))

    return pd.concat(frames, ignore_index=True)


def get_covariate_summary(features_df: pd.DataFrame, config: dict = None) -> pd.DataFrame:
    """
    Return current covariate values and z-scores for dashboard display.
    """
    if config is None:
        config = load_config()

    past_covs = config["data"]["past_covariates"]
    records = []
    for col in past_covs:
        if col in features_df.columns:
            current = features_df[col].iloc[-1]
            mean = features_df[col].mean()
            std = features_df[col].std()
            z_score = (current - mean) / std if std > 0 else 0.0

            if abs(z_score) > 2:
                status = "Extreme"
            elif abs(z_score) > 1:
                status = "Elevated"
            else:
                status = "Normal"

            records.append({
                "Covariate": col,
                "Current": f"{current:.4f}",
                "Z-Score": f"{z_score:.2f}",
                "Status": status,
            })

    return pd.DataFrame(records)
