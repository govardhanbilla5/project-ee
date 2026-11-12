"""
Application configuration and physical constants.
 
All tunable constants live here so they can be changed in one place
without touching business logic. Every value can be overridden with an
environment variable or a `.env` file (pydantic-settings reads both),
e.g.  DEADBAND_PCT=0.02  or  PRICE_LAG_MONTHS=2.
 
The forecast engine reads these values in forecast_engine.py:
    price_next = P_current + [(AMS_Q - AMS_Q-1) x PWt] + (PPI_Factor x P_current)
    AMS        = (MC x DF_C) + CNG
    MC         = (LME + Midwest) + MC_OFFSET
"""

from pydantic_settings import BaseSettings
from datetime import datetime, timezone

# NEW Modification -11/12/2026 - Requested by thapse (19-31)
# Fixed forecast starting month
START_YEAR = 2026
START_MONTH = 10
 
# Get current month
now_utc = datetime.now(timezone.utc)
 
# Calculate inclusive month gap
month_gap = (
    (now_utc.year - START_YEAR) * 12
    + now_utc.month - START_MONTH
    + 1
)
class Settings(BaseSettings):
    """App-level settings. Override via environment variables."""

    APP_NAME: str = "Aluminium Price Forecast API"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = False

    # Density / conversion drauss factor used in AMS = (MC x DF_C) + CNG
    DF_C: float = 1.44

    #── MODIFICATION ──────────────────────────────────
    # Here MC_OFFSET is added newly from thapase/srikanth requirements on 09/30/2026
    # Historical MC formula was (LME + Midwest) x MC_MULTIPLIER + MC_OFFSET.
    # The CURRENT engine no longer multiplies: MC = (LME + Midwest) + MC_OFFSET.
    # MC_MULTIPLIER is therefore NOT used by forecast_engine.py. It is kept only
    MC_MULTIPLIER: float = 0.91
    # CHANGED from 1.2 in earlier versions and this is the new value as per thapase/srikanth requirements on 09/30/2026
    MC_OFFSET: float = 0.558883

    # Actual vs projected data 
    # Published (actual) data ends this many months before the current month;
    # anything later in the sheet is a projection.
    # Example: sitting in October with lag 1 -> actuals run through September.
    # Used for the previous-quarter values (MC_Q-1 and PPI_Q-1):
    #   - last month of previous quarter is actual   -> use that month's value
    #   - otherwise (partly projected)               -> use the 3-month average of the prevoius quarter
    PPI_ACTUAL_LAG_MONTHS: int = 1

    # Number of future months to forecast.
    # (13 months are produced when include_current_month = "YES".)
    #FORECAST_HORIZON_MONTHS: int = 12
    FORECAST_HORIZON_MONTHS: int = 12 + month_gap -1

    # ── MODIFICATION ────────────────────────────────────────
    # This is the new feature as per thapase/srikanth requirements on 09/30/2026
    # Deadband: if the calculated price is within +/- this fraction of the base
    # price (0.02 = 2%), the base price is retained (no change).
    # Formula: |raw_price - P_current| / P_current <= DEADBAND_PCT
    DEADBAND_PCT: float = 0.02
    # Implementation lag in months (~50 days).
    PRICE_LAG_MONTHS: int = 2

    class Config:
        env_file = ".env"
        case_sensitive = True


settings = Settings()
