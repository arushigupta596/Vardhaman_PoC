"""
FastAPI serving endpoint for Chronos-2 multivariate cotton futures forecasting.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from datetime import date
from typing import Optional
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
import pandas as pd
import numpy as np

app = FastAPI(
    title="Cotton Futures Chronos-2 Multivariate API",
    description="30/60/90-day cotton price forecasting with covariates and cross-learning",
    version="3.0.0",
)

DATA_DIR = Path(__file__).resolve().parent.parent / "data"


class ForecastRequest(BaseModel):
    as_of_date: Optional[str] = None
    horizon: int = 90


class ForecastResponse(BaseModel):
    as_of: str
    horizon: int
    p10: list[float]
    p25: list[float]
    p50: list[float]
    p75: list[float]
    p90: list[float]
    dates: list[str]
    covariates_used: dict


class CovariateStatus(BaseModel):
    covariate: str
    current_value: float
    z_score: float
    status: str


@app.get("/health")
async def health():
    """Health check — verify model and data availability."""
    features_exist = (DATA_DIR / "features" / "features.parquet").exists()
    return {
        "status": "healthy" if features_exist else "degraded",
        "model": "amazon/chronos-2",
        "mode": "multivariate + covariates + cross-learning",
        "features_available": features_exist,
    }


@app.post("/forecast", response_model=ForecastResponse)
async def forecast_endpoint(req: ForecastRequest):
    """Generate multivariate forecast with covariates."""
    try:
        features = pd.read_parquet(DATA_DIR / "features" / "features.parquet")
    except FileNotFoundError:
        raise HTTPException(status_code=503, detail="Features not built. Run pipeline first.")

    if req.as_of_date:
        as_of = pd.Timestamp(req.as_of_date)
        if as_of not in features.index:
            # Find nearest date
            idx = features.index.searchsorted(as_of)
            idx = min(idx, len(features) - 1)
            features = features.iloc[:idx + 1]
        else:
            features = features.loc[:as_of]
    as_of_str = features.index[-1].strftime("%Y-%m-%d")

    from src.model import forecast
    from src.covariates import load_config

    config = load_config()
    result = forecast(features, horizon=req.horizon, config=config)

    dates = pd.bdate_range(features.index[-1] + pd.Timedelta(days=1), periods=req.horizon)

    return ForecastResponse(
        as_of=as_of_str,
        horizon=req.horizon,
        p10=result["q10"].tolist(),
        p25=result["q25"].tolist(),
        p50=result["median"].tolist(),
        p75=result["q75"].tolist(),
        p90=result["q90"].tolist(),
        dates=[d.strftime("%Y-%m-%d") for d in dates],
        covariates_used={
            "past": config["data"]["past_covariates"],
            "future": config["data"]["known_future_covariates"],
            "cross_learning": config["data"].get("cross_learning_series", []),
        },
    )


@app.get("/covariates", response_model=list[CovariateStatus])
async def covariates_endpoint():
    """Return current covariate values and z-scores."""
    try:
        features = pd.read_parquet(DATA_DIR / "features" / "features.parquet")
    except FileNotFoundError:
        raise HTTPException(status_code=503, detail="Features not built.")

    from src.covariates import get_covariate_summary
    summary = get_covariate_summary(features)

    return [
        CovariateStatus(
            covariate=row["Covariate"],
            current_value=float(row["Current"]),
            z_score=float(row["Z-Score"]),
            status=row["Status"],
        )
        for _, row in summary.iterrows()
    ]


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
