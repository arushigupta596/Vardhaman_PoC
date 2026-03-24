"""
Streamlit dashboard — Cotton Futures Forecasting with Chronos-2 Multivariate.
Full analysis with covariates, CFTC COT, Weather, Macro, Technical signals,
and Bias & Regime Analysis.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from datetime import datetime

st.set_page_config(
    page_title="Cotton Futures — Chronos-2 Multivariate",
    page_icon="🏭",
    layout="wide",
)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
RESULTS_DIR = Path(__file__).resolve().parent.parent / "results"

# ── Color palette ──
NAVY = "#1B3A5C"
DARK_NAVY = "#0F2440"
BLUE = "#2E86C1"
RED = "#E74C3C"
GREEN = "#27AE60"
ORANGE = "#F39C12"
PURPLE = "#8E44AD"

REGIME_COLORS = {"up": GREEN, "down": RED, "sideways": ORANGE}
REGIME_ICONS = {"up": "🟢", "down": "🔴", "sideways": "🟡"}


@st.cache_data
def load_data():
    features = pd.read_parquet(DATA_DIR / "features" / "features.parquet")
    chronos = None
    prophet = None
    try:
        chronos = pd.read_csv(RESULTS_DIR / "backtest_metrics.csv", parse_dates=["as_of"])
    except FileNotFoundError:
        pass
    try:
        prophet = pd.read_csv(RESULTS_DIR / "prophet_backtest.csv", parse_dates=["as_of"])
    except FileNotFoundError:
        pass
    return features, chronos, prophet


@st.cache_data
def load_live_forecast():
    try:
        return pd.read_csv(RESULTS_DIR / "live_forecast.csv", parse_dates=["date"])
    except FileNotFoundError:
        return None


def main():
    features, chronos_bt, prophet_bt = load_data()
    live = load_live_forecast()

    # ── Compute regime & bias info ──
    from src.model import (detect_regime, compute_ewma_bias,
                            compute_regime_ewma_bias, optimize_ensemble_weights,
                            _compute_static_bias)
    current_regime = detect_regime(features["ct1_close"])
    ewma_bias = compute_ewma_bias(alpha=0.3)
    regime_bias = compute_regime_ewma_bias(alpha=0.3)
    static_bias = _compute_static_bias()
    opt_weights = optimize_ensemble_weights([30, 60, 90])

    # ── Header ──
    st.title("Cotton Futures Forecasting — Chronos-2 Multivariate")
    st.markdown("**ICE CT1** | Chronos-2 (Multivariate + Covariates + Cross-Learning) | 30/60/90-day horizons")

    # ── Current Price Card ──
    last_price = features["ct1_close"].iloc[-1]
    last_date = features.index[-1]
    price_chg = features["ct1_close"].iloc[-1] - features["ct1_close"].iloc[-2]
    pct_chg = price_chg / features["ct1_close"].iloc[-2] * 100

    col1, col2, col3, col4, col5 = st.columns(5)
    col1.metric("CT1 Close", f"{last_price:.2f} ¢/lb", f"{price_chg:+.2f} ({pct_chg:+.1f}%)")
    col2.metric("21d Realised Vol", f"{features['realised_vol_21d'].iloc[-1]:.1%}")
    col3.metric("RSI-14", f"{features['rsi_14'].iloc[-1]:.1f}")
    if "comm_net_z" in features.columns:
        col4.metric("Commercial Net Z", f"{features['comm_net_z'].iloc[-1]:.2f}")
    col5.metric("Market Regime", f"{REGIME_ICONS.get(current_regime, '')} {current_regime.upper()}")

    st.divider()

    # ── TAB LAYOUT ──
    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "Live Forecast", "Backtest Results", "Bias & Regime Analysis",
        "CFTC COT Analysis", "Weather & Drought", "Signal Dashboard",
        "Covariate Impact"
    ])

    # ═══════════════════════════════════════════════════════════
    # TAB 1: LIVE FORECAST
    # ═══════════════════════════════════════════════════════════
    with tab1:
        st.header("Live Chronos-2 Multivariate Forecast")

        # Regime & bias correction info
        r_color = REGIME_COLORS.get(current_regime, ORANGE)
        st.markdown(
            f"**Current Regime:** <span style='color:{r_color}; font-weight:bold;'>"
            f"{REGIME_ICONS.get(current_regime, '')} {current_regime.upper()}</span> "
            f"&nbsp;|&nbsp; **Bias Correction:** Regime-EWMA (α=0.3)",
            unsafe_allow_html=True
        )
        if regime_bias:
            bias_parts = []
            for h in [30, 60, 90]:
                if h in regime_bias:
                    b = regime_bias[h].get(current_regime, regime_bias[h].get("global", 0))
                    bias_parts.append(f"{h}d: {b:+.2f}")
            if bias_parts:
                st.caption(f"Applied bias correction ({current_regime} regime): {', '.join(bias_parts)} ¢/lb")

        with st.expander("Active Covariates & Model Info", expanded=False):
            st.markdown("**Past Covariates** (17 total):")
            st.markdown("- `dxy`, `wti_crude` — Macro factors")
            st.markdown("- `spec_net_pct`, `cert_stocks_z/chg_5d/chg_21d` — Positioning & supply")
            st.markdown("- `realised_vol_21d`, `ct1_ret_5d/21d` — Momentum & volatility")
            st.markdown("- `noaa_pdsi`, `pdsi_severe_drought` — Weather/drought")
            st.markdown("**Known Future Covariates** (6): Seasonal Fourier terms, smooth crop calendar flags, WASDE windows")
            st.markdown("**Cross-Learning**: Group attention across CT1, DXY, WTI")
            st.markdown("**Ensemble**: Chronos-2 + Chronos-Bolt with optimized per-horizon weights")
            if opt_weights:
                for h, info in sorted(opt_weights.items()):
                    st.markdown(f"- {h}d: C2={info['weight_chronos2']:.0%} / Bolt={info['weight_bolt']:.0%}")

        if live is not None:
            fig = go.Figure()
            hist = features["ct1_close"].tail(120)
            fig.add_trace(go.Scatter(
                x=hist.index, y=hist.values,
                name="Historical", line=dict(color=NAVY, width=2)
            ))
            fig.add_trace(go.Scatter(
                x=live["date"], y=live["median"],
                name="Median Forecast", line=dict(color=RED, width=2, dash="dash")
            ))
            fig.add_trace(go.Scatter(
                x=pd.concat([live["date"], live["date"][::-1]]),
                y=pd.concat([live["q90"], live["q10"][::-1]]),
                fill="toself", fillcolor="rgba(231,76,60,0.15)",
                line=dict(width=0), name="80% Interval (q10-q90)"
            ))
            if "q25" in live.columns and "q75" in live.columns:
                fig.add_trace(go.Scatter(
                    x=pd.concat([live["date"], live["date"][::-1]]),
                    y=pd.concat([live["q75"], live["q25"][::-1]]),
                    fill="toself", fillcolor="rgba(231,76,60,0.25)",
                    line=dict(width=0), name="50% Interval (q25-q75)"
                ))
            fig.update_layout(
                title="CT1 Cotton Futures — 90-Day Multivariate Forecast (Regime-Corrected)",
                yaxis_title="¢/lb", height=500, template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

            fcols = st.columns(3)
            for i, h in enumerate([30, 60, 90]):
                row = live.iloc[h - 1] if h <= len(live) else None
                if row is not None:
                    direction = "↑" if row["median"] > last_price else "↓"
                    fcols[i].metric(
                        f"{h}-Day Forecast",
                        f"{row['median']:.2f} ¢/lb {direction}",
                        f"Range: [{row['q10']:.2f}, {row['q90']:.2f}]"
                    )
        else:
            st.info("Run live forecast: `python run_pipeline.py --live`")
            if st.button("Generate Live Forecast"):
                with st.spinner("Running Chronos-2 multivariate inference ..."):
                    from src.model import forecast
                    result = forecast(features, horizon=90)
                    dates = pd.bdate_range(last_date + pd.Timedelta(days=1), periods=90)
                    live_df = pd.DataFrame({
                        "date": dates, "median": result["median"],
                        "q10": result["q10"], "q25": result["q25"],
                        "q75": result["q75"], "q90": result["q90"],
                    })
                    live_df.to_csv(RESULTS_DIR / "live_forecast.csv", index=False)
                    st.success("Forecast generated!")
                    st.rerun()

    # ═══════════════════════════════════════════════════════════
    # TAB 2: BACKTEST RESULTS
    # ═══════════════════════════════════════════════════════════
    with tab2:
        st.header("Walk-Forward Backtest Comparison")

        if chronos_bt is None or prophet_bt is None:
            st.warning("Run full pipeline first: `python run_pipeline.py`")
        else:
            horizon = st.selectbox("Horizon", [30, 60, 90], key="bt_horizon")

            c_sub = chronos_bt[chronos_bt.horizon == horizon]
            p_sub = prophet_bt[prophet_bt.horizon == horizon]

            col1, col2, col3 = st.columns(3)
            with col1:
                st.subheader("Chronos-2 Ensemble")
                st.metric("MAE", f"{c_sub.mae.mean():.2f} ¢/lb")
                st.metric("CRPS", f"{c_sub.crps.mean():.3f}")
                st.metric("Dir. Accuracy", f"{c_sub.dir_acc.mean():.1%}")
                st.metric("80% Coverage", f"{c_sub.coverage.mean():.1%}")
                if "signed_error" in c_sub.columns:
                    st.metric("Bias", f"{c_sub.signed_error.mean():+.2f} ¢/lb")
            with col2:
                st.subheader("Prophet")
                st.metric("MAE", f"{p_sub.mae.mean():.2f} ¢/lb")
                st.metric("Dir. Accuracy", f"{p_sub.dir_acc.mean():.1%}")
                st.metric("80% Coverage", f"{p_sub.coverage.mean():.1%}")
            with col3:
                st.subheader("DM Test")
                merged = chronos_bt.merge(prophet_bt, on=["as_of", "horizon"], suffixes=("_c", "_p"))
                sub = merged[merged.horizon == horizon]
                if len(sub) > 0:
                    from src.evaluate import diebold_mariano_test
                    t_stat, p_val = diebold_mariano_test(sub["mae_c"].values, sub["mae_p"].values)
                    st.metric("t-statistic", f"{t_stat:.3f}")
                    st.metric("p-value", f"{p_val:.6f}")
                    improvement = (p_sub.mae.mean() - c_sub.mae.mean()) / p_sub.mae.mean() * 100
                    st.metric("MAE Improvement", f"{improvement:.1f}%")

            # Error over time chart
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=c_sub["as_of"], y=c_sub["mae"],
                name="Chronos-2 Ensemble MAE", line=dict(color=BLUE)
            ))
            fig.add_trace(go.Scatter(
                x=p_sub["as_of"], y=p_sub["mae"],
                name="Prophet MAE", line=dict(color=RED)
            ))
            fig.update_layout(
                title=f"{horizon}-Day MAE Over Time",
                yaxis_title="MAE (¢/lb)", height=400, template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Actual vs Predicted scatter — colored by regime
            if "regime" in c_sub.columns:
                fig2 = go.Figure()
                for regime in ["up", "down", "sideways"]:
                    r_data = c_sub[c_sub["regime"] == regime]
                    if len(r_data) > 0:
                        fig2.add_trace(go.Scatter(
                            x=r_data["actual"], y=r_data["predicted"],
                            mode="markers", name=f"{regime.capitalize()}",
                            marker=dict(color=REGIME_COLORS.get(regime, BLUE), size=8)
                        ))
                min_v = min(c_sub["actual"].min(), c_sub["predicted"].min())
                max_v = max(c_sub["actual"].max(), c_sub["predicted"].max())
                fig2.add_trace(go.Scatter(
                    x=[min_v, max_v], y=[min_v, max_v],
                    mode="lines", line=dict(dash="dash", color="gray"),
                    showlegend=False
                ))
                fig2.update_layout(
                    height=400, template="plotly_white",
                    title=f"{horizon}d: Actual vs Predicted (by Regime)",
                    xaxis_title="Actual (¢/lb)", yaxis_title="Predicted (¢/lb)"
                )
                st.plotly_chart(fig2, use_container_width=True)
            else:
                fig2 = make_subplots(rows=1, cols=2, subplot_titles=["Chronos-2", "Prophet"])
                fig2.add_trace(go.Scatter(
                    x=c_sub["actual"], y=c_sub["predicted"],
                    mode="markers", name="Chronos-2", marker=dict(color=BLUE, size=6)
                ), row=1, col=1)
                fig2.add_trace(go.Scatter(
                    x=p_sub["actual"], y=p_sub["predicted"],
                    mode="markers", name="Prophet", marker=dict(color=RED, size=6)
                ), row=1, col=2)
                min_v = min(c_sub["actual"].min(), c_sub["predicted"].min())
                max_v = max(c_sub["actual"].max(), c_sub["predicted"].max())
                for col in [1, 2]:
                    fig2.add_trace(go.Scatter(
                        x=[min_v, max_v], y=[min_v, max_v],
                        mode="lines", line=dict(dash="dash", color="gray"),
                        showlegend=False
                    ), row=1, col=col)
                fig2.update_layout(height=400, template="plotly_white",
                                   title=f"{horizon}d: Actual vs Predicted")
                st.plotly_chart(fig2, use_container_width=True)

    # ═══════════════════════════════════════════════════════════
    # TAB 3: BIAS & REGIME ANALYSIS (NEW)
    # ═══════════════════════════════════════════════════════════
    with tab3:
        st.header("Bias & Regime Analysis")
        st.markdown("Analysis of systematic forecast bias and market regime-dependent corrections.")

        if chronos_bt is None:
            st.warning("Run backtest first: `python run_pipeline.py`")
        else:
            # ── Current regime indicator ──
            st.subheader("Current Market Regime")
            r_color = REGIME_COLORS.get(current_regime, ORANGE)
            st.markdown(
                f"<div style='background-color:{r_color}20; border-left:4px solid {r_color}; "
                f"padding:15px; border-radius:5px; margin-bottom:20px;'>"
                f"<span style='font-size:24px;'>{REGIME_ICONS.get(current_regime, '')}</span> "
                f"<span style='font-size:20px; font-weight:bold; color:{r_color};'>"
                f"{current_regime.upper()} MARKET</span><br>"
                f"<span style='color:#666;'>Based on 20-day price trend ({'> +2%' if current_regime == 'up' else '< -2%' if current_regime == 'down' else 'within +/-2%'})</span>"
                f"</div>",
                unsafe_allow_html=True
            )

            # ── Before/After improvement cards ──
            st.subheader("Accuracy Improvement (v3.5 → v4.0)")
            before_mae = {30: 2.70, 60: 3.38, 90: 3.11}
            before_bias = {30: 1.27, 60: 2.13, 90: 2.60}
            icols = st.columns(3)
            for i, h in enumerate([30, 60, 90]):
                after_mae = chronos_bt[chronos_bt.horizon == h]["mae"].mean()
                after_bias = chronos_bt[chronos_bt.horizon == h]["signed_error"].mean()
                mae_pct = (before_mae[h] - after_mae) / before_mae[h] * 100
                bias_pct = (before_bias[h] - after_bias) / before_bias[h] * 100
                with icols[i]:
                    st.metric(f"{h}d MAE", f"{after_mae:.2f} ¢/lb",
                              f"{mae_pct:+.0f}% (was {before_mae[h]:.2f})")
                    st.metric(f"{h}d Bias", f"{after_bias:+.2f} ¢/lb",
                              f"{bias_pct:+.0f}% (was +{before_bias[h]:.2f})")

            st.divider()

            # ── Signed error over time by horizon ──
            st.subheader("Forecast Bias Over Time")
            fig_bias = go.Figure()
            horizon_colors = {30: BLUE, 60: PURPLE, 90: RED}
            for h in [30, 60, 90]:
                h_data = chronos_bt[chronos_bt.horizon == h].sort_values("as_of")
                fig_bias.add_trace(go.Scatter(
                    x=h_data["as_of"], y=h_data["signed_error"],
                    name=f"{h}d", mode="lines+markers",
                    line=dict(color=horizon_colors[h]),
                    marker=dict(size=6)
                ))
            fig_bias.add_hline(y=0, line_dash="dash", line_color="gray")
            fig_bias.update_layout(
                title="Signed Error (Predicted - Actual) Over Time",
                yaxis_title="Signed Error (¢/lb)", height=400,
                template="plotly_white",
                annotations=[dict(
                    x=0.01, y=0.98, xref="paper", yref="paper",
                    text="Above 0 = Over-prediction<br>Below 0 = Under-prediction",
                    showarrow=False, bgcolor="white", bordercolor="gray",
                    borderwidth=1, font=dict(size=11)
                )]
            )
            st.plotly_chart(fig_bias, use_container_width=True)

            # ── Signed error by regime (box plot) ──
            if "regime" in chronos_bt.columns:
                st.subheader("Bias by Market Regime")
                fig_box = go.Figure()
                for regime in ["up", "down", "sideways"]:
                    r_data = chronos_bt[chronos_bt["regime"] == regime]
                    if len(r_data) > 0:
                        fig_box.add_trace(go.Box(
                            y=r_data["signed_error"],
                            name=regime.capitalize(),
                            marker_color=REGIME_COLORS[regime],
                            boxmean=True
                        ))
                fig_box.add_hline(y=0, line_dash="dash", line_color="gray")
                fig_box.update_layout(
                    title="Signed Error Distribution by Regime",
                    yaxis_title="Signed Error (¢/lb)", height=400,
                    template="plotly_white"
                )
                st.plotly_chart(fig_box, use_container_width=True)

            # ── Regime-dependent EWMA bias table ──
            st.subheader("Regime-Dependent EWMA Bias (¢/lb)")
            if regime_bias:
                regime_table = []
                for h in [30, 60, 90]:
                    if h in regime_bias:
                        row = {"Horizon": f"{h}-day"}
                        for r in ["up", "down", "sideways", "global"]:
                            row[r.capitalize()] = f"{regime_bias[h].get(r, 0):+.2f}"
                        regime_table.append(row)
                if regime_table:
                    st.dataframe(pd.DataFrame(regime_table).set_index("Horizon"),
                                 use_container_width=True)

                # Regime distribution
                if "regime" in chronos_bt.columns:
                    regime_counts = chronos_bt.groupby("regime").size()
                    st.caption(f"Regime distribution in backtest: {dict(regime_counts)}")

            st.divider()

            # ── EWMA vs Static bias comparison ──
            st.subheader("Bias Correction Methods Comparison")
            if ewma_bias and static_bias:
                bias_comp = []
                for h in [30, 60, 90]:
                    s = static_bias.get(h, 0)
                    e = ewma_bias.get(h, 0)
                    r_val = regime_bias.get(h, {}).get(current_regime, e) if regime_bias else e
                    bias_comp.append({
                        "Horizon": f"{h}-day",
                        "Static (mean)": f"{s:+.2f}",
                        "EWMA (α=0.3)": f"{e:+.2f}",
                        f"Regime-EWMA ({current_regime})": f"{r_val:+.2f}",
                    })
                st.dataframe(pd.DataFrame(bias_comp).set_index("Horizon"),
                             use_container_width=True)
                st.caption("Regime-EWMA adapts to current market conditions; lower absolute values = less bias to correct.")

            # ── Optimized ensemble weights ──
            st.subheader("Optimized Ensemble Weights")
            if opt_weights:
                weight_data = []
                for h in [30, 60, 90]:
                    if h in opt_weights:
                        info = opt_weights[h]
                        weight_data.append({
                            "Horizon": f"{h}-day",
                            "Default": "C2: 60% / Bolt: 40%",
                            "Optimized": f"C2: {info['weight_chronos2']:.0%} / Bolt: {info['weight_bolt']:.0%}",
                            "MAE Improvement": f"{info['improvement_pct']:+.1f}%",
                        })
                st.dataframe(pd.DataFrame(weight_data).set_index("Horizon"),
                             use_container_width=True)

                # Component model comparison chart
                if "c2_median" in chronos_bt.columns and "bolt_median" in chronos_bt.columns:
                    st.subheader("Chronos-2 vs Chronos-Bolt by Horizon")
                    bt_horizon = st.selectbox("Select Horizon", [30, 60, 90], key="comp_horizon")
                    h_data = chronos_bt[chronos_bt.horizon == bt_horizon].sort_values("as_of")
                    if len(h_data) > 0:
                        fig_comp = go.Figure()
                        fig_comp.add_trace(go.Scatter(
                            x=h_data["as_of"], y=h_data["actual"],
                            name="Actual", line=dict(color="black", width=2)
                        ))
                        fig_comp.add_trace(go.Scatter(
                            x=h_data["as_of"], y=h_data["c2_median"],
                            name="Chronos-2", line=dict(color=BLUE, dash="dot")
                        ))
                        fig_comp.add_trace(go.Scatter(
                            x=h_data["as_of"], y=h_data["bolt_median"],
                            name="Chronos-Bolt", line=dict(color=ORANGE, dash="dot")
                        ))
                        fig_comp.add_trace(go.Scatter(
                            x=h_data["as_of"], y=h_data["predicted"],
                            name="Ensemble", line=dict(color=GREEN, width=2)
                        ))
                        fig_comp.update_layout(
                            title=f"{bt_horizon}d: Individual Models vs Ensemble vs Actual",
                            yaxis_title="¢/lb", height=400, template="plotly_white"
                        )
                        st.plotly_chart(fig_comp, use_container_width=True)

            # ── Signed error histogram ──
            st.subheader("Signed Error Distribution")
            fig_hist = go.Figure()
            for h in [30, 60, 90]:
                h_data = chronos_bt[chronos_bt.horizon == h]
                fig_hist.add_trace(go.Histogram(
                    x=h_data["signed_error"], name=f"{h}d",
                    marker_color=horizon_colors[h], opacity=0.6,
                    nbinsx=15
                ))
            fig_hist.add_vline(x=0, line_dash="dash", line_color="gray")
            fig_hist.update_layout(
                title="Distribution of Signed Errors (All Horizons)",
                xaxis_title="Signed Error (¢/lb)", yaxis_title="Count",
                barmode="overlay", height=350, template="plotly_white"
            )
            st.plotly_chart(fig_hist, use_container_width=True)

    # ═══════════════════════════════════════════════════════════
    # TAB 4: CFTC COT ANALYSIS
    # ═══════════════════════════════════════════════════════════
    with tab4:
        st.header("CFTC Commitments of Traders — Cotton")

        cot_cols = [c for c in features.columns if any(x in c for x in ["comm_", "spec_", "mm_", "cot_"])]
        if cot_cols:
            fig = make_subplots(rows=3, cols=1, shared_xaxes=True,
                                subplot_titles=["CT1 Price", "Net Positioning (Z-score)", "COT Open Interest Z"],
                                vertical_spacing=0.08)

            fig.add_trace(go.Scatter(
                x=features.index, y=features["ct1_close"],
                name="CT1", line=dict(color=NAVY)
            ), row=1, col=1)

            z_cols = {"comm_net_z": ("Commercial", GREEN),
                      "spec_net_z": ("Speculator", RED),
                      "mm_net_z": ("Managed Money", PURPLE)}
            for col, (name, color) in z_cols.items():
                if col in features.columns:
                    fig.add_trace(go.Scatter(
                        x=features.index, y=features[col],
                        name=name, line=dict(color=color)
                    ), row=2, col=1)

            if "cot_oi_z" in features.columns:
                fig.add_trace(go.Scatter(
                    x=features.index, y=features["cot_oi_z"],
                    name="OI Z-score", line=dict(color=ORANGE)
                ), row=3, col=1)

            fig.update_layout(height=700, template="plotly_white",
                              title="CFTC COT Positioning vs Cotton Price")
            st.plotly_chart(fig, use_container_width=True)

            st.subheader("Current Positioning")
            pos_cols = {
                "comm_net_z": "Commercial Net Z",
                "spec_net_z": "Speculator Net Z",
                "mm_net_z": "Managed Money Net Z",
                "comm_net_pct": "Commercial Net %",
                "spec_net_pct": "Speculator Net %",
                "cot_oi_z": "OI Z-score",
            }
            pos_data = []
            for col, label in pos_cols.items():
                if col in features.columns:
                    val = features[col].iloc[-1]
                    pos_data.append({"Signal": label, "Value": f"{val:.3f}",
                                    "Interpretation": "Bullish" if val > 0.5 else ("Bearish" if val < -0.5 else "Neutral")})
            st.dataframe(pd.DataFrame(pos_data), use_container_width=True)
        else:
            st.warning("No COT data available. Run ingest to fetch CFTC data.")

    # ═══════════════════════════════════════════════════════════
    # TAB 5: WEATHER & DROUGHT
    # ═══════════════════════════════════════════════════════════
    with tab5:
        st.header("Weather & Drought — Lubbock TX (Cotton Belt)")

        fig = make_subplots(rows=4, cols=1, shared_xaxes=True,
                            subplot_titles=["CT1 Price", "Temperature & Heat Stress",
                                          "Precipitation & Water Stress", "NOAA PDSI Drought Index"],
                            vertical_spacing=0.06)

        fig.add_trace(go.Scatter(
            x=features.index, y=features["ct1_close"],
            name="CT1", line=dict(color=NAVY)
        ), row=1, col=1)

        if "temp_max" in features.columns:
            fig.add_trace(go.Scatter(
                x=features.index, y=features["temp_max"],
                name="Temp Max", line=dict(color=RED, width=1)
            ), row=2, col=1)
            fig.add_trace(go.Scatter(
                x=features.index, y=features["temp_min"],
                name="Temp Min", line=dict(color="#3498DB", width=1)
            ), row=2, col=1)
            fig.add_hline(y=95, line_dash="dash", line_color="red",
                         annotation_text="Heat Stress (95F)", row=2, col=1)

        if "precip_30d" in features.columns:
            fig.add_trace(go.Bar(
                x=features.index, y=features["precip_30d"],
                name="30d Precip", marker_color="#3498DB", opacity=0.5
            ), row=3, col=1)
        if "water_stress_30d" in features.columns:
            fig.add_trace(go.Scatter(
                x=features.index, y=features["water_stress_30d"],
                name="Water Stress 30d", line=dict(color=RED)
            ), row=3, col=1)

        if "noaa_pdsi" in features.columns:
            fig.add_trace(go.Scatter(
                x=features.index, y=features["noaa_pdsi"],
                name="PDSI", fill="tozeroy",
                fillcolor="rgba(231,76,60,0.2)", line=dict(color="#C0392B")
            ), row=4, col=1)
            fig.add_hline(y=-3, line_dash="dash", line_color="red",
                         annotation_text="Severe Drought", row=4, col=1)

        fig.update_layout(height=900, template="plotly_white",
                          title="Weather Impact on Cotton")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Current Conditions")
        cond = {
            "PDSI": f"{features['noaa_pdsi'].iloc[-1]:.2f}" if "noaa_pdsi" in features.columns else "N/A",
            "Heat Stress 5d": f"{features['heat_stress_5d'].iloc[-1]:.0f}" if "heat_stress_5d" in features.columns else "N/A",
            "Water Stress 30d": f"{features['water_stress_30d'].iloc[-1]:.2f}" if "water_stress_30d" in features.columns else "N/A",
            "GDD 30d": f"{features['gdd_cumulative_30d'].iloc[-1]:.0f}" if "gdd_cumulative_30d" in features.columns else "N/A",
        }
        cols = st.columns(len(cond))
        for i, (k, v) in enumerate(cond.items()):
            cols[i].metric(k, v)

    # ═══════════════════════════════════════════════════════════
    # TAB 6: SIGNAL DASHBOARD
    # ═══════════════════════════════════════════════════════════
    with tab6:
        st.header("Multi-Signal Dashboard")

        fig = make_subplots(rows=5, cols=1, shared_xaxes=True,
                            subplot_titles=["CT1 Price + Bollinger", "RSI-14", "DXY Deviation",
                                          "Realized Volatility", "MACD Signal"],
                            vertical_spacing=0.05)

        sma20 = features["ct1_close"].rolling(20).mean()
        std20 = features["ct1_close"].rolling(20).std()
        fig.add_trace(go.Scatter(x=features.index, y=features["ct1_close"],
                                  name="CT1", line=dict(color=NAVY)), row=1, col=1)
        fig.add_trace(go.Scatter(x=features.index, y=sma20 + 2*std20,
                                  name="Upper BB", line=dict(color="gray", dash="dash", width=1)), row=1, col=1)
        fig.add_trace(go.Scatter(x=features.index, y=sma20 - 2*std20,
                                  name="Lower BB", line=dict(color="gray", dash="dash", width=1),
                                  fill="tonexty", fillcolor="rgba(128,128,128,0.1)"), row=1, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["rsi_14"],
                                  name="RSI-14", line=dict(color=PURPLE)), row=2, col=1)
        fig.add_hline(y=70, line_dash="dash", line_color="red", row=2, col=1)
        fig.add_hline(y=30, line_dash="dash", line_color="green", row=2, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["dxy_dev"],
                                  name="DXY Dev", line=dict(color=ORANGE)), row=3, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["realised_vol_21d"],
                                  name="21d Vol", line=dict(color=RED)), row=4, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["macd_signal"],
                                  name="MACD Signal", line=dict(color="#2ECC71")), row=5, col=1)

        fig.update_layout(height=1000, template="plotly_white", showlegend=False,
                          title="Technical & Macro Signals")
        st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════════════════
    # TAB 7: COVARIATE IMPACT
    # ═══════════════════════════════════════════════════════════
    with tab7:
        st.header("Covariate Impact Analysis")
        st.markdown("Covariates used by Chronos-2 for multivariate forecasting.")

        # ICE Certified Stocks
        if "certified_stocks" in features.columns:
            st.subheader("ICE Certified Stocks")
            cert = features[features["certified_stocks"] > 0]
            if len(cert) > 0:
                fig_cert = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                         subplot_titles=["CT1 Price", "ICE Certified Stocks (bales)"])
                fig_cert.add_trace(go.Scatter(
                    x=cert.index, y=cert["ct1_close"],
                    name="CT1", line=dict(color=NAVY)
                ), row=1, col=1)
                fig_cert.add_trace(go.Scatter(
                    x=cert.index, y=cert["certified_stocks"],
                    name="Certified Stocks", fill="tozeroy",
                    fillcolor="rgba(46,134,193,0.2)", line=dict(color=BLUE)
                ), row=2, col=1)
                fig_cert.update_layout(height=400, template="plotly_white",
                                       title="ICE Certified Stocks vs Cotton Price")
                st.plotly_chart(fig_cert, use_container_width=True)

        # Current covariate status
        st.subheader("Current Covariate Values & Z-Scores")
        try:
            from src.covariates import get_covariate_summary
            cov_summary = get_covariate_summary(features)
            if len(cov_summary) > 0:
                st.dataframe(cov_summary, use_container_width=True)
        except Exception:
            st.info("Covariate summary not available.")

        # Covariate correlation with target
        st.subheader("Covariate Correlation with CT1")
        past_covs = ["dxy", "wti_crude", "comm_net_z", "spec_net_z",
                      "realised_vol_21d", "noaa_pdsi", "cert_stocks_z"]
        available_covs = [c for c in past_covs if c in features.columns]

        if available_covs:
            corr_data = []
            for col in available_covs:
                for lag in [0, 5, 10, 20]:
                    if lag == 0:
                        corr = features["ct1_close"].corr(features[col])
                    else:
                        corr = features["ct1_close"].corr(features[col].shift(lag))
                    corr_data.append({"Covariate": col, "Lag (days)": lag, "Correlation": corr})

            corr_df = pd.DataFrame(corr_data)
            pivot = corr_df.pivot(index="Covariate", columns="Lag (days)", values="Correlation")
            fig = go.Figure(data=go.Heatmap(
                z=pivot.values,
                x=[f"Lag {c}d" for c in pivot.columns],
                y=pivot.index,
                colorscale="RdBu_r", zmid=0,
                text=pivot.values.round(3), texttemplate="%{text}",
            ))
            fig.update_layout(
                title="Covariate-Target Correlation at Different Lags",
                height=400, template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

        # Covariate time series
        st.subheader("Covariate History")
        selected_cov = st.selectbox("Select Covariate", available_covs if available_covs else ["N/A"])
        if selected_cov != "N/A" and selected_cov in features.columns:
            fig = make_subplots(rows=2, cols=1, shared_xaxes=True,
                                subplot_titles=["CT1 Price", selected_cov])
            fig.add_trace(go.Scatter(x=features.index, y=features["ct1_close"],
                                      name="CT1", line=dict(color=NAVY)), row=1, col=1)
            fig.add_trace(go.Scatter(x=features.index, y=features[selected_cov],
                                      name=selected_cov, line=dict(color=RED)), row=2, col=1)
            fig.update_layout(height=500, template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

    # ── Footer ──
    st.divider()
    st.caption(f"Data: {features.index.min().date()} - {features.index.max().date()} | "
               f"{len(features)} trading days | {len(features.columns)} features | "
               f"Model: Chronos-2 Multivariate + Chronos-Bolt Ensemble | "
               f"Bias: Regime-EWMA | Regime: {current_regime.upper()}")


if __name__ == "__main__":
    main()
