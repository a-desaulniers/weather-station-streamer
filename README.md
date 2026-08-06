# Weather Station Streamer (Google Apps Script)

Automated weather data collection from **Davis WeatherLink** and **HOBOlink RX3000** into Google Sheets with metric conversion, automatic authentication, and 15‑minute scheduled updates.

## Data Sources

### Davis WeatherLink (n stations)
Temperature, humidity, wind, barometric pressure, solar radiation, UV, rainfall, evapotranspiration, daily highs/lows with timestamps, sunrise/sunset, moon phase.

### HOBOlink RX3000
Temperature, RH, dew point, heat index, wet bulb, wind speed/gust/direction, pressure, rainfall, soil water content. Full 10‑minute time‑series from LICOR API.

## Setup

See [docs/setup.md](docs/setup.md) for step‑by‑step instructions.

## Quick Start

1. Open Google Sheets → Extensions → Apps Script
2. Paste `src/Code.gs` into the editor
3. Run `setCredentials()` once with your Davis login
4. Run `setHoboCredentials()` once with your HOBOlink API token
5. Run `streamWeatherLinkNetwork()` to test
6. Add a 15‑minute time‑driven trigger

## License

MIT
