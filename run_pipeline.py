#!/usr/bin/env python3
"""
Master pipeline — Chronos-2 Multivariate Cotton Futures Forecasting.

Usage:
    python run_pipeline.py             # Full pipeline (ingest + features + backtest + prophet + live)
    python run_pipeline.py --live      # Live forecast only (skip backtest)
    python run_pipeline.py --backtest  # Backtest only (skip ingest)
    python run_pipeline.py --univariate  # Run univariate (no covariates) for comparison
"""
import sys
import argparse
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))


def main():
    parser = argparse.ArgumentParser(description="Chronos-2 Multivariate Cotton Forecasting Pipeline")
    parser.add_argument("--live", action="store_true", help="Live forecast only")
    parser.add_argument("--backtest", action="store_true", help="Backtest only (skip ingest)")
    parser.add_argument("--univariate", action="store_true", help="Run without covariates for comparison")
    args = parser.parse_args()

    t0 = time.time()

    if not args.backtest and not args.live:
        # ── Stage 1: Ingest ──
        print("=" * 60)
        print("STAGE 1: DATA INGESTION")
        print("=" * 60)
        from src.ingest import ingest_all
        ingest_all()

        # ── Stage 2: Features ──
        print("\n" + "=" * 60)
        print("STAGE 2: FEATURE ENGINEERING")
        print("=" * 60)
        from src.features import build_features
        build_features()

    # ── Stage 3: Covariate Validation ──
    print("\n" + "=" * 60)
    print("STAGE 3: COVARIATE VALIDATION")
    print("=" * 60)
    import pandas as pd
    from src.covariates import validate_covariates, load_config

    config = load_config()
    features = pd.read_parquet("data/features/features.parquet")

    validation = validate_covariates(features, config)
    if validation["valid"]:
        print(f"  → All covariates validated")
        print(f"  → Past covariates: {validation['available_past']}")
        print(f"  → Future covariates: {validation['available_future']}")
    else:
        print(f"  → WARNING: Missing covariates: {validation['missing']}")
    if validation["high_nan"]:
        print(f"  → High NaN covariates (>10%): {validation['high_nan']}")

    if not args.live:
        # ── Stage 4: Chronos-2 Multivariate Walk-Forward Backtest ──
        print("\n" + "=" * 60)
        print("STAGE 4: CHRONOS-2 MULTIVARIATE WALK-FORWARD BACKTEST")
        print("=" * 60)
        from src.evaluate import walk_forward_backtest

        use_cross_learning = not args.univariate
        use_ensemble = config.get("ensemble", {}).get("enabled", False) and not args.univariate

        if use_ensemble:
            w_c2 = config["ensemble"].get("weight_chronos2", 0.6)
            w_bolt = config["ensemble"].get("weight_bolt", 0.4)
            print(f"  → Ensemble mode: Chronos-2 ({w_c2}) + Chronos-Bolt ({w_bolt})")

            def multivariate_model_fn(df, origin_idx, h):
                from src.model import forecast_ensemble
                return forecast_ensemble(
                    df, origin_idx, h,
                    config=config,
                    weight_chronos2=w_c2,
                    weight_bolt=w_bolt,
                    return_components=True,
                )
        else:
            def multivariate_model_fn(df, origin_idx, h):
                from src.model import forecast_at_origin
                return forecast_at_origin(
                    df, origin_idx, h,
                    config=config,
                    cross_learning=use_cross_learning,
                )

        output_file = "backtest_metrics.csv" if not args.univariate else "backtest_metrics_univariate.csv"
        chronos_results = walk_forward_backtest(
            features,
            config=config,
            model_fn=multivariate_model_fn,
            output_file=output_file,
        )

        # ── Stage 5: Prophet Baseline Backtest ──
        print("\n" + "=" * 60)
        print("STAGE 5: PROPHET BASELINE BACKTEST")
        print("=" * 60)
        from src.baselines.prophet_baseline import prophet_backtest
        prophet_results = prophet_backtest(features)

        # ── Stage 6: Model Comparison & DM Test ──
        print("\n" + "=" * 60)
        print("STAGE 6: MODEL COMPARISON & DM TEST")
        print("=" * 60)
        from src.evaluate import diebold_mariano_test

        merged = chronos_results.merge(
            prophet_results, on=["as_of", "horizon"], suffixes=("_c", "_p")
        )

        mode_label = "Chronos-2 Multivariate" if not args.univariate else "Chronos-2 Univariate"
        for h in config["forecast"]["horizons"]:
            sub = merged[merged.horizon == h]
            if len(sub) > 0:
                t_stat, p_val = diebold_mariano_test(sub["mae_c"].values, sub["mae_p"].values)
                c_mae = sub["mae_c"].mean()
                p_mae = sub["mae_p"].mean()
                imp = (p_mae - c_mae) / p_mae * 100
                winner = mode_label if t_stat < 0 else "Prophet"
                sig = "***" if p_val < 0.001 else ("**" if p_val < 0.01 else ("*" if p_val < 0.05 else "n.s."))
                print(f"\n  {h}d: {mode_label} MAE={c_mae:.2f}  Prophet MAE={p_mae:.2f}")
                print(f"       DM t={t_stat:.3f}  p={p_val:.6f} {sig}")
                print(f"       Winner: {winner}  Improvement: {imp:.1f}%")

    # ── Stage 7: Live Forecast ──
    print("\n" + "=" * 60)
    print("STAGE 7: LIVE FORECAST (MULTIVARIATE)")
    print("=" * 60)
    import pandas as pd

    features = pd.read_parquet("data/features/features.parquet")
    from src.model import (forecast, load_bias_estimates, apply_bias_correction,
                            forecast_bolt_univariate, detect_regime,
                            optimize_ensemble_weights)

    max_horizon = max(config["forecast"]["horizons"])
    use_ensemble = config.get("ensemble", {}).get("enabled", False) and not args.univariate

    result = forecast(
        features,
        horizon=max_horizon,
        config=config,
        cross_learning=not args.univariate,
    )

    # Ensemble with Chronos-Bolt — use optimized per-horizon weights if available
    if use_ensemble:
        opt_weights = optimize_ensemble_weights(config["forecast"]["horizons"])
        bolt_result = forecast_bolt_univariate(
            features[config["data"]["target"]], max_horizon, config["forecast"]["quantile_levels"]
        )

        if opt_weights:
            # Apply optimized per-horizon weights to the full forecast arrays
            # Use the max-horizon optimized weight as default for the full array
            max_h = max(config["forecast"]["horizons"])
            w_c2 = opt_weights.get(max_h, {}).get("weight_chronos2", config["ensemble"].get("weight_chronos2", 0.6))
            w_bolt = 1 - w_c2
            for key in ["q10", "q25", "median", "q75", "q90", "mean"]:
                if key in result and key in bolt_result:
                    result[key] = w_c2 * result[key] + w_bolt * bolt_result[key]
            weight_info = ", ".join(f"{h}d: C2={info['weight_chronos2']:.0%}/Bolt={info['weight_bolt']:.0%}"
                                    for h, info in sorted(opt_weights.items()))
            print(f"  → Ensemble (optimized weights): {weight_info}")
        else:
            w_c2 = config["ensemble"].get("weight_chronos2", 0.6)
            w_bolt = config["ensemble"].get("weight_bolt", 0.4)
            for key in ["q10", "q25", "median", "q75", "q90", "mean"]:
                if key in result and key in bolt_result:
                    result[key] = w_c2 * result[key] + w_bolt * bolt_result[key]
            print(f"  → Ensemble applied: Chronos-2 ({w_c2}) + Chronos-Bolt ({w_bolt})")

    # Apply regime-dependent bias correction if enabled
    bias_cfg = config.get("bias_correction", {})
    if bias_cfg.get("enabled", False):
        method = bias_cfg.get("method", "regime_ewma")
        alpha = bias_cfg.get("alpha", 0.3)

        # Detect current market regime
        current_regime = detect_regime(features[config["data"]["target"]])
        print(f"  → Current market regime: {current_regime}")

        bias_estimates = load_bias_estimates(method=method, alpha=alpha, regime=current_regime)
        if bias_estimates:
            for h in config["forecast"]["horizons"]:
                result = apply_bias_correction(result, h, bias_estimates)
            print(f"  → Bias correction ({method}, α={alpha}, regime={current_regime}): "
                  f"{{{', '.join(f'{k}d: {v:+.2f}' for k, v in sorted(bias_estimates.items()))}}}")
        else:
            print(f"  → Bias correction enabled but no estimates available (run backtest first)")

    last_date = features.index[-1]
    dates = pd.bdate_range(last_date + pd.Timedelta(days=1), periods=max_horizon)

    live_df = pd.DataFrame({
        "date": dates,
        "median": result["median"],
        "q10": result["q10"],
        "q25": result["q25"],
        "q75": result["q75"],
        "q90": result["q90"],
    })
    live_df.to_csv("results/live_forecast.csv", index=False)
    print(f"  → Live forecast saved: {len(live_df)} days")

    for h in config["forecast"]["horizons"]:
        row = live_df.iloc[h - 1]
        direction = "↑" if row["median"] > features["ct1_close"].iloc[-1] else "↓"
        print(f"  → {h}d: {row['median']:.2f} ¢/lb {direction}  [{row['q10']:.2f}, {row['q90']:.2f}]")

    # Save and display secondary target forecasts
    secondary_preds = result.get("secondary", {})
    if secondary_preds:
        from src.model import SECONDARY_TARGET_LABELS
        print(f"\n  Secondary Futures Feature Forecasts:")
        for tgt, preds in secondary_preds.items():
            label = SECONDARY_TARGET_LABELS.get(tgt, tgt)
            sec_df = pd.DataFrame({
                "date": dates,
                "median": preds["median"],
                "q10": preds["q10"],
                "q90": preds["q90"],
            })
            sec_df.to_csv(f"results/live_forecast_{tgt}.csv", index=False)

            for h in config["forecast"]["horizons"]:
                val = preds["median"][h - 1]
                low = preds["q10"][h - 1]
                high = preds["q90"][h - 1]
                # Get current value for comparison
                current = features[tgt].iloc[-1] if tgt in features.columns else 0
                direction = "↑" if val > current else "↓"
                print(f"    {tgt} {h}d: {val:.4f} {direction}  [{low:.4f}, {high:.4f}]")

    # Print covariate summary
    print("\n  Active Covariates:")
    past_covs = config["data"]["past_covariates"]
    future_covs = config["data"]["known_future_covariates"]
    print(f"    Past: {past_covs}")
    print(f"    Future: {future_covs}")
    print(f"    Cross-learning: {'enabled' if not args.univariate else 'disabled'}")

    elapsed = time.time() - t0
    print(f"\n{'=' * 60}")
    print(f"PIPELINE COMPLETE — {elapsed:.0f}s total")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
