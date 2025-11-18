Weather Parameters Export Summary
================================
Generated: 2025-10-24 14:20:47 UTC
Location: CEF Tomnatic
GPS Coordinates: 45.9926, 20.6813
Capacity: 2.425 MW

Weather Data Source: Open-Meteo API
Data Period: 2025-10-24 22:00:00+00:00 to 2025-10-30 23:45:00+00:00

Key Statistics:
- GHI Max: 509 W/m² (if 'ghi' in weather_df else 'N/A')
- GHI Mean (daylight): 213 W/m² (if 'ghi' in weather_df else 'N/A')
- Cloud Cover Mean: 51.1% (if 'cloud_cover' in weather_df else 'N/A')
- Temperature Range: 5.3°C to 18.6°C (if 'temperature' in weather_df else 'N/A')

Files Created:
- location_info.json: GPS coordinates and configuration
- raw_weather_data.csv: Complete weather data used
- weather_statistics.json: Statistical summary
- daily_weather_summary.csv: Day-by-day breakdown
- model_parameters.json: PV system parameters
