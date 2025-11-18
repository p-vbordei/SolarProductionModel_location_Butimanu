# Solar Forecast Email System - CEF Tomnatic

Solar production forecasting system for CEF Tomnatic, Timiș County, Romania.

## Core Scripts

1. **run_intraday_cm.py** - Generates solar production forecast
2. **export_forecast_to_excel.py** - Creates Excel report from forecast data
3. **send_forecast_zoho.py** - Sends email with Excel attachment

## Wrapper Script

**run_forecast_and_email.py** - Runs all 3 scripts in sequence

## Quick Start

```bash
# Using uv (recommended)
uv run python scripts/run_forecast_and_email.py

# Using traditional Python
python scripts/run_forecast_and_email.py
```

## Docker

```bash
# Run scheduled (6 AM and 2 PM)
docker-compose up -d

# Run once
docker-compose --profile manual up solar-forecast-once
```

## CEF Tomnatic Configuration

- **Location**: Tomnatic, Timiș County, Romania
- **Coordinates**: 45°59'33.4"N, 20°40'52.7"E (45.9926°N, 20.6813°E)
- **DC Capacity**: 2.919 MW (2919 kW)
- **AC Capacity**: 2.5 MW (limited to 2.425 MW)
- **Panels**: 4,200 × Risen Energy RSM132-8-695BHDG (695W each)
  - Total DC Power: 2.919 MW
  - Panel Tilt: 25°
  - Orientation: South (180° azimuth)
- **Inverters**: 25 × Huawei SUN2000-100KTL-M2 (100 kW each)
  - Total AC Power: 2.5 MW (nameplate)
  - Limited to: 2.425 MW (operational limit)
- **Email Config**: Must have `scripts/email_config_zoho_working.json`
- **Weather Data**: Uses Open-Meteo API (real data only)

## System Features

- ✅ 7-day rolling forecast with 15-minute resolution
- ✅ Probabilistic forecasts (P10, P25, P50, P75, P90)
- ✅ Excel reports with multiple aggregation levels
- ✅ Automated email delivery via Zoho Mail
- ✅ Timezone-aware outputs (CET/CEST)