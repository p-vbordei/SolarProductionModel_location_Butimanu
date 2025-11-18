# Solar Forecast Email System - CEF Butimanu

Solar production forecasting system for CEF Butimanu, Dâmbovița County, Romania.

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

## CEF Butimanu Configuration

- **Location**: Butimanu, Dâmbovița County, Romania
- **Coordinates**: 44°40'59.4"N, 25°54'25.4"E (44.6832°N, 25.9071°E)
- **DC Capacity**: 12.806 MW (12,806 kW)
- **AC Capacity**: 10.8 MW (limited to 10.8 MW)
- **Panels**: 22,080 × LONGi LR5-72HIBD580M (580W each)
  - Total DC Power: 12.806 MW
  - Panel Tilt: 20°
  - Orientation: South (180° azimuth)
- **Inverters**: 54 × Huawei SUN2000-215KTL-H1 (200 kW each)
  - Total AC Power: 10.8 MW (nameplate)
  - Limited to: 10.8 MW (operational limit)
- **Email Config**: Must have `scripts/email_config_zoho_working.json`
- **Weather Data**: Uses Open-Meteo API (real data only)

## System Features

- ✅ 7-day rolling forecast with 15-minute resolution
- ✅ Probabilistic forecasts (P10, P25, P50, P75, P90)
- ✅ Excel reports with multiple aggregation levels
- ✅ Automated email delivery via Zoho Mail
- ✅ Timezone-aware outputs (CET/CEST)