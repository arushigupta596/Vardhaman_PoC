"""
Streamlit dashboard — Cotton Futures Forecasting with Chronos-2 Multivariate.
Full analysis with covariates, CFTC COT, Weather, Macro, and Technical signals.
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

    # ── Header ──
    st.title("Cotton Futures Forecasting — Chronos-2 Multivariate")
    st.markdown("**ICE CT1** | Chronos-2 (Multivariate + Covariates + Cross-Learning) | 30/60/90-day horizons")

    # ── Current Price Card ──
    last_price = features["ct1_close"].iloc[-1]
    last_date = features.index[-1]
    price_chg = features["ct1_close"].iloc[-1] - features["ct1_close"].iloc[-2]
    pct_chg = price_chg / features["ct1_close"].iloc[-2] * 100

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("CT1 Close", f"{last_price:.2f} ¢/lb", f"{price_chg:+.2f} ({pct_chg:+.1f}%)")
    col2.metric("21d Realised Vol", f"{features['realised_vol_21d'].iloc[-1]:.1%}")
    col3.metric("RSI-14", f"{features['rsi_14'].iloc[-1]:.1f}")

    if "comm_net_z" in features.columns:
        col4.metric("Commercial Net Z-score", f"{features['comm_net_z'].iloc[-1]:.2f}")

    st.divider()

    # ── TAB LAYOUT ──
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "Live Forecast", "Backtest Results", "CFTC COT Analysis",
        "Weather & Drought", "Signal Dashboard", "Covariate Impact"
    ])

    # ═══════════════════════════════════════════════════════════
    # TAB 1: LIVE FORECAST
    # ═══════════════════════════════════════════════════════════
    with tab1:
        st.header("Live Chronos-2 Multivariate Forecast")

        # Sidebar: Active covariates
        with st.expander("Active Covariates", expanded=False):
            st.markdown("**Past Covariates** (historical only):")
            st.markdown("- `ct2_close`, `ct3_close` — Term structure")
            st.markdown("- `dxy` — US Dollar Index")
            st.markdown("- `wti_crude` — WTI Crude Oil")
            st.markdown("- `comm_net_z` — CFTC Commercial Net Z-score")
            st.markdown("- `realised_vol_21d` — 21-day Realized Volatility")
            st.markdown("- `noaa_pdsi` — Drought Index")
            st.markdown("**Known Future Covariates** (deterministic):")
            st.markdown("- Seasonal Fourier terms, Crop calendar flags, WASDE windows")
            st.markdown("**Cross-Learning**: Group attention across CT1, CT2, CT3, DXY, WTI")

        if live is not None:
            fig = go.Figure()
            # Historical
            hist = features["ct1_close"].tail(120)
            fig.add_trace(go.Scatter(
                x=hist.index, y=hist.values,
                name="Historical", line=dict(color="#1B4F72", width=2)
            ))
            # Forecast
            fig.add_trace(go.Scatter(
                x=live["date"], y=live["median"],
                name="Median Forecast", line=dict(color="#E74C3C", width=2, dash="dash")
            ))
            # 80% interval (q10-q90)
            fig.add_trace(go.Scatter(
                x=pd.concat([live["date"], live["date"][::-1]]),
                y=pd.concat([live["q90"], live["q10"][::-1]]),
                fill="toself", fillcolor="rgba(231,76,60,0.15)",
                line=dict(width=0), name="80% Interval (q10–q90)"
            ))
            # 50% interval (q25-q75)
            if "q25" in live.columns and "q75" in live.columns:
                fig.add_trace(go.Scatter(
                    x=pd.concat([live["date"], live["date"][::-1]]),
                    y=pd.concat([live["q75"], live["q25"][::-1]]),
                    fill="toself", fillcolor="rgba(231,76,60,0.25)",
                    line=dict(width=0), name="50% Interval (q25–q75)"
                ))
            fig.update_layout(
                title="CT1 Cotton Futures — 90-Day Multivariate Forecast",
                yaxis_title="¢/lb", height=500,
                template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Forecast table
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
            st.info("Run live forecast to see predictions. Use: `python run_pipeline.py --live`")

            if st.button("Generate Live Forecast"):
                with st.spinner("Running Chronos-2 multivariate inference …"):
                    from src.model import forecast
                    result = forecast(features, horizon=90)
                    dates = pd.bdate_range(last_date + pd.Timedelta(days=1), periods=90)
                    live_df = pd.DataFrame({
                        "date": dates,
                        "median": result["median"],
                        "q10": result["q10"],
                        "q25": result["q25"],
                        "q75": result["q75"],
                        "q90": result["q90"],
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

            # Summary metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.subheader("Chronos-2 Multivariate")
                st.metric("MAE", f"{c_sub.mae.mean():.2f} ¢/lb")
                st.metric("CRPS", f"{c_sub.crps.mean():.3f}")
                st.metric("Dir. Accuracy", f"{c_sub.dir_acc.mean():.1%}")
                st.metric("80% Coverage", f"{c_sub.coverage.mean():.1%}")
                if "mase" in c_sub.columns:
                    st.metric("MASE", f"{c_sub.mase.mean():.2f}")
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
                name="Chronos-2 Multivariate MAE", line=dict(color="#2E86C1")
            ))
            fig.add_trace(go.Scatter(
                x=p_sub["as_of"], y=p_sub["mae"],
                name="Prophet MAE", line=dict(color="#E74C3C")
            ))
            fig.update_layout(
                title=f"{horizon}-Day MAE Over Time",
                yaxis_title="MAE (¢/lb)", height=400,
                template="plotly_white"
            )
            st.plotly_chart(fig, use_container_width=True)

            # Actual vs Predicted scatter
            fig2 = make_subplots(rows=1, cols=2, subplot_titles=["Chronos-2 Multivariate", "Prophet"])
            fig2.add_trace(go.Scatter(
                x=c_sub["actual"], y=c_sub["predicted"],
                mode="markers", name="Chronos-2", marker=dict(color="#2E86C1", size=6)
            ), row=1, col=1)
            fig2.add_trace(go.Scatter(
                x=p_sub["actual"], y=p_sub["predicted"],
                mode="markers", name="Prophet", marker=dict(color="#E74C3C", size=6)
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
    # TAB 3: CFTC COT ANALYSIS
    # ═══════════════════════════════════════════════════════════
    with tab3:
        st.header("CFTC Commitments of Traders — Cotton")

        cot_cols = [c for c in features.columns if any(x in c for x in ["comm_", "spec_", "mm_", "cot_"])]
        if cot_cols:
            fig = make_subplots(rows=3, cols=1, shared_xaxes=True,
                                subplot_titles=["CT1 Price", "Net Positioning (Z-score)", "COT Open Interest Z"],
                                vertical_spacing=0.08)

            fig.add_trace(go.Scatter(
                x=features.index, y=features["ct1_close"],
                name="CT1", line=dict(color="#1B4F72")
            ), row=1, col=1)

            z_cols = {"comm_net_z": ("Commercial", "#27AE60"),
                      "spec_net_z": ("Speculator", "#E74C3C"),
                      "mm_net_z": ("Managed Money", "#8E44AD")}
            for col, (name, color) in z_cols.items():
                if col in features.columns:
                    fig.add_trace(go.Scatter(
                        x=features.index, y=features[col],
                        name=name, line=dict(color=color)
                    ), row=2, col=1)

            if "cot_oi_z" in features.columns:
                fig.add_trace(go.Scatter(
                    x=features.index, y=features["cot_oi_z"],
                    name="OI Z-score", line=dict(color="#F39C12")
                ), row=3, col=1)

            fig.update_layout(height=700, template="plotly_white",
                              title="CFTC COT Positioning vs Cotton Price")
            st.plotly_chart(fig, use_container_width=True)

            # Current positioning summary
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
    # TAB 4: WEATHER & DROUGHT
    # ═══════════════════════════════════════════════════════════
    with tab4:
        st.header("Weather & Drought — Lubbock TX (Cotton Belt)")

        fig = make_subplots(rows=4, cols=1, shared_xaxes=True,
                            subplot_titles=["CT1 Price", "Temperature & Heat Stress",
                                          "Precipitation & Water Stress", "NOAA PDSI Drought Index"],
                            vertical_spacing=0.06)

        fig.add_trace(go.Scatter(
            x=features.index, y=features["ct1_close"],
            name="CT1", line=dict(color="#1B4F72")
        ), row=1, col=1)

        if "temp_max" in features.columns:
            fig.add_trace(go.Scatter(
                x=features.index, y=features["temp_max"],
                name="Temp Max (°F)", line=dict(color="#E74C3C", width=1)
            ), row=2, col=1)
            fig.add_trace(go.Scatter(
                x=features.index, y=features["temp_min"],
                name="Temp Min (°F)", line=dict(color="#3498DB", width=1)
            ), row=2, col=1)
            fig.add_hline(y=95, line_dash="dash", line_color="red",
                         annotation_text="Heat Stress (95°F)", row=2, col=1)

        if "precip_30d" in features.columns:
            fig.add_trace(go.Bar(
                x=features.index, y=features["precip_30d"],
                name="30d Precip (in)", marker_color="#3498DB", opacity=0.5
            ), row=3, col=1)
        if "water_stress_30d" in features.columns:
            fig.add_trace(go.Scatter(
                x=features.index, y=features["water_stress_30d"],
                name="Water Stress 30d", line=dict(color="#E74C3C")
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
            "Heat Stress Days (5d)": f"{features['heat_stress_5d'].iloc[-1]:.0f}" if "heat_stress_5d" in features.columns else "N/A",
            "Water Stress 30d": f"{features['water_stress_30d'].iloc[-1]:.2f}" if "water_stress_30d" in features.columns else "N/A",
            "GDD 30d Cumulative": f"{features['gdd_cumulative_30d'].iloc[-1]:.0f}" if "gdd_cumulative_30d" in features.columns else "N/A",
        }
        cols = st.columns(len(cond))
        for i, (k, v) in enumerate(cond.items()):
            cols[i].metric(k, v)

    # ═══════════════════════════════════════════════════════════
    # TAB 5: SIGNAL DASHBOARD
    # ═══════════════════════════════════════════════════════════
    with tab5:
        st.header("Multi-Signal Dashboard")

        fig = make_subplots(rows=5, cols=1, shared_xaxes=True,
                            subplot_titles=["CT1 Price + Bollinger", "RSI-14", "DXY Deviation",
                                          "Realized Volatility", "MACD Signal"],
                            vertical_spacing=0.05)

        sma20 = features["ct1_close"].rolling(20).mean()
        std20 = features["ct1_close"].rolling(20).std()
        fig.add_trace(go.Scatter(x=features.index, y=features["ct1_close"],
                                  name="CT1", line=dict(color="#1B4F72")), row=1, col=1)
        fig.add_trace(go.Scatter(x=features.index, y=sma20 + 2*std20,
                                  name="Upper BB", line=dict(color="gray", dash="dash", width=1)), row=1, col=1)
        fig.add_trace(go.Scatter(x=features.index, y=sma20 - 2*std20,
                                  name="Lower BB", line=dict(color="gray", dash="dash", width=1),
                                  fill="tonexty", fillcolor="rgba(128,128,128,0.1)"), row=1, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["rsi_14"],
                                  name="RSI-14", line=dict(color="#8E44AD")), row=2, col=1)
        fig.add_hline(y=70, line_dash="dash", line_color="red", row=2, col=1)
        fig.add_hline(y=30, line_dash="dash", line_color="green", row=2, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["dxy_dev"],
                                  name="DXY Dev", line=dict(color="#F39C12")), row=3, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["realised_vol_21d"],
                                  name="21d Vol", line=dict(color="#E74C3C")), row=4, col=1)

        fig.add_trace(go.Scatter(x=features.index, y=features["macd_signal"],
                                  name="MACD Signal", line=dict(color="#2ECC71")), row=5, col=1)

        fig.update_layout(height=1000, template="plotly_white", showlegend=False,
                          title="Technical & Macro Signals")
        st.plotly_chart(fig, use_container_width=True)

    # ═══════════════════════════════════════════════════════════
    # TAB 6: COVARIATE IMPACT (NEW)
    # ═══════════════════════════════════════════════════════════
    with tab6:
        st.header("Covariate Impact Analysis")
        st.markdown("Covariates used by Chronos-2 for multivariate forecasting.")

        # Current covariate status
        st.subheader("Current Covariate Values & Z-Scores")
        from src.covariates import get_covariate_summary
        cov_summary = get_covariate_summary(features)
        if len(cov_summary) > 0:
            st.dataframe(cov_summary, use_container_width=True)

        # Covariate correlation with target
        st.subheader("Covariate Correlation with CT1")
        past_covs = ["ct2_close", "ct3_close", "dxy", "wti_crude", "comm_net_z", "realised_vol_21d", "noaa_pdsi"]
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

            # Heatmap
            pivot = corr_df.pivot(index="Covariate", columns="Lag (days)", values="Correlation")
            fig = go.Figure(data=go.Heatmap(
                z=pivot.values,
                x=[f"Lag {c}d" for c in pivot.columns],
                y=pivot.index,
                colorscale="RdBu_r",
                zmid=0,
                text=pivot.values.round(3),
                texttemplate="%{text}",
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
                                      name="CT1", line=dict(color="#1B4F72")), row=1, col=1)
            fig.add_trace(go.Scatter(x=features.index, y=features[selected_cov],
                                      name=selected_cov, line=dict(color="#E74C3C")), row=2, col=1)
            fig.update_layout(height=500, template="plotly_white")
            st.plotly_chart(fig, use_container_width=True)

    # ── Footer ──
    st.divider()
    st.caption(f"Data: {features.index.min().date()} – {features.index.max().date()} | "
               f"{len(features)} trading days | {len(features.columns)} features | "
               f"Model: Chronos-2 Multivariate + Covariates + Cross-Learning (amazon/chronos-2)")


if __name__ == "__main__":
    main()
