// Combined Davis & HoboLink Weather Streamer
// Author: Alexandre Desaulniers, Quebec Innovative Materials Corporation
// Credentials are stored securely in Script Properties

const WL_USERNAME = PropertiesService.getScriptProperties().getProperty('WL_USERNAME');
const WL_PASSWORD = PropertiesService.getScriptProperties().getProperty('WL_PASSWORD');
const HOBOLINK_TOKEN = PropertiesService.getScriptProperties().getProperty('HOBOLINK_API_KEY');

const TIMEZONE_OFFSET = '-10800000';
const IANA_TIMEZONE = 'America%2FHalifax';
const USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36';

// Davis WeatherLink stations
const DAVIS_STATIONS = [
  { uuid: '', sheetName: 'example1_NSW080' },
  { uuid: '', sheetName: 'example2_NSW039' },
  { uuid: '', sheetName: 'example3_NSW082' },
  { uuid: '', sheetName: 'example3_NSW081' },
  { uuid: '', sheetName: 'example4_NSW065' },
  { uuid: '', sheetName: 'example5_NSW063' },
  { uuid: '', sheetName: 'example6_NSW066' },
  { uuid: '', sheetName: 'example7_NSW091' }
];

// HOBOlink configuration
const HOBO_DEVICE_SERIAL = '';
const HOBO_SHEET_NAME = 'example8_RX3000';
const HOBO_LOOKBACK_HOURS = 24; // How far back to fetch on first run

// unit conversions
function fToC(f) {
  if (f === '' || f === undefined || f === null) return '';
  return Math.round((parseFloat(f) - 32) * 5 / 9 * 10) / 10;
}

function mphToKmh(mph) {
  if (mph === '' || mph === undefined || mph === null) return '';
  return Math.round(parseFloat(mph) * 1.60934 * 10) / 10;
}

function inHgToHpa(inHg) {
  if (inHg === '' || inHg === undefined || inHg === null) return '';
  return Math.round(parseFloat(inHg) * 33.8639 * 10) / 10;
}

function inToMm(inches) {
  if (inches === '' || inches === undefined || inches === null) return '';
  return Math.round(parseFloat(inches) * 25.4 * 10) / 10;
}

function roundVal(val) {
  if (val === undefined || val === null) return '';
  return Math.round(parseFloat(val) * 100) / 100;
}

// main on 15 minute trigger

function streamWeatherLinkNetwork() {
  Logger.log('=== Starting combined fetch ===');
  const logSheet = getLogSheet();

  // WeatherLink auth
  const authCookie = loginAndGetSession(logSheet);
  if (authCookie) {
    DAVIS_STATIONS.forEach((st, i) => {
      if (!st.uuid) return;
      fetchDavisStation(st.uuid, st.sheetName, authCookie, logSheet);
      if (i < DAVIS_STATIONS.length - 1) Utilities.sleep(3000);
    });
  } else {
    Logger.log('Davis login failed — skipping WeatherLink stations.');
  }

  // Hobolink API Auth
  fetchHobolinkData(logSheet);

  Logger.log('=== Combined fetch complete ===');

    // Force dashboard to recalculate
  const dash = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DASHBOARD');
  if (dash) {
    dash.getRange('A1').setValue(dash.getRange('A1').getValue());
  }

  Logger.log('=== Combined fetch complete ===');
}

// Davis login attempt

function loginAndGetSession(logSheet) {
  const loginUrl = 'https://www.weatherlink.com/login';
  const processUrl = 'https://www.weatherlink.com/processLogin';

  try {
    const getResp = UrlFetchApp.fetch(loginUrl, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': USER_AGENT }
    });
    
    const cookies = extractCookies(getResp.getAllHeaders()['Set-Cookie']);
    const sessionId = cookies.find(c => c.startsWith('JSESSIONID='));
    
    if (!sessionId) {
      Logger.log('No JSESSIONID from login page.');
      return null;
    }

    const postResp = UrlFetchApp.fetch(processUrl, {
      method: 'post',
      payload: {
        username: WL_USERNAME,
        password: WL_PASSWORD,
        rememberMe: 'false',
        localTimezoneOffset: TIMEZONE_OFFSET,
        ianaTimeZone: IANA_TIMEZONE
      },
      contentType: 'application/x-www-form-urlencoded',
      muteHttpExceptions: true,
      followRedirects: false,
      headers: {
        'User-Agent': USER_AGENT,
        'Accept': '*/*',
        'X-Requested-With': 'XMLHttpRequest',
        'Referer': loginUrl,
        'Origin': 'https://www.weatherlink.com',
        'Cookie': sessionId
      }
    });

    if (postResp.getResponseCode() === 200) {
      return sessionId;
    }
    
    Logger.log(`Davis login failed: ${postResp.getResponseCode()}`);
    logRun(logSheet, 'DAVIS_LOGIN', 'FAILED', `HTTP ${postResp.getResponseCode()}`);
    return null;
    
  } catch (err) {
    Logger.log(`Davis login error: ${err}`);
    logRun(logSheet, 'DAVIS_LOGIN', 'FAILED', err.toString());
    return null;
  }
}

// Fetch davis data attempt

function fetchDavisStation(uuid, sheetName, cookie, logSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  const headers = [
    "Timestamp", "Temp Out (°C)", "Temp In (°C)", "Hum Out (%)", "Hum In (%)",
    "Heat Index (°C)", "THSW Index (°C)", "Wind Chill (°C)", "THW Index (°C)",
    "Dew Point (°C)", "Wet Bulb (°C)",
    "Wind Speed (km/h)", "Wind Dir (°)", "Wind Rose",
    "2-Min Wind (km/h)", "10-Min Wind (km/h)", "10-Min Gust (km/h)",
    "Barometer (hPa)", "Bar Trend Text", "Bar Trend Val",
    "Solar Rad (W/m²)", "UV Index",
    "Rain Rate (mm/h)", "Rain Day (mm)", "Rain Month (mm)", "Rain Year (mm)", "Rain Storm (mm)",
    "Rain 1-Hour (mm)",
    "ET Day (mm)", "ET Month (mm)", "ET Year (mm)",
    "High Temp Out (°C)", "High Temp Out Time",
    "Low Temp Out (°C)", "Low Temp Out Time",
    "High Hum Out (%)", "High Hum Out Time",
    "Low Hum Out (%)", "Low Hum Out Time",
    "High Dew Point (°C)", "High Dew Point Time",
    "Low Dew Point (°C)", "Low Dew Point Time",
    "High Heat Index (°C)", "High Heat Index Time",
    "Low Wind Chill (°C)", "Low Wind Chill Time",
    "High Wind Speed (km/h)", "High Wind Speed Time",
    "High Solar Rad (W/m²)", "High Solar Rad Time",
    "High Barometer (hPa)", "High Barometer Time",
    "Low Barometer (hPa)", "Low Barometer Time",
    "Low Wet Bulb (°C)", "Low Wet Bulb Time",
    "Sunrise", "Sunset", "Moon Phase",
    "Extra Temp 1 (°C)", "Extra Hum 1 (%)",
    "Extra Temp 2 (°C)", "Extra Hum 2 (%)",
    "Extra Temp 3 (°C)", "Extra Hum 3 (%)",
    "Extra Temp 4 (°C)", "Extra Hum 4 (%)",
    "Extra Temp 5 (°C)", "Extra Hum 5 (%)",
    "Extra Temp 6 (°C)", "Extra Hum 6 (%)",
    "Extra Temp 7 (°C)", "Extra Hum 7 (%)",
    "Soil Temp 1 (°C)", "Soil Moist 1 (cb)",
    "Soil Temp 2 (°C)", "Soil Moist 2 (cb)",
    "Soil Temp 3 (°C)", "Soil Moist 3 (cb)",
    "Soil Temp 4 (°C)", "Soil Moist 4 (cb)",
    "Leaf Temp 1 (°C)", "Leaf Wet 1",
    "Leaf Temp 2 (°C)", "Leaf Wet 2",
    "Leaf Temp 3 (°C)", "Leaf Wet 3",
    "Leaf Temp 4 (°C)", "Leaf Wet 4"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
  }

  try {
    const resp = UrlFetchApp.fetch(`https://www.weatherlink.com/bulletin/${uuid}`, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': USER_AGENT, 'Cookie': cookie }
    });

    if (resp.getResponseCode() !== 200) {
      logRun(logSheet, sheetName, 'FAILED', `HTTP ${resp.getResponseCode()}`);
      return;
    }

    const match = resp.getContentText().match(/wl\.__bootstrap_station_data\s*=\s*(\{[\s\S]*?\});/);
    if (!match) {
      logRun(logSheet, sheetName, 'NODATA', 'No station data block');
      return;
    }

    const data = JSON.parse(match[1]);
    const iss = data.oIssData || {};
    const bar = data.oBarData || {};
    const hilow = data.oIssHilowData || {};

    // --- helpers ---
    const formatTime = (val) => {
      if (!val || val === 65535) return '';
      const h = Math.floor(val / 100), m = val % 100;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    };

    const extract = (val) => {
      if (val === undefined || val === null) return '';
      if (typeof val === 'number' || typeof val === 'string') return val;
      if (typeof val === 'object') {
        if (val.value !== undefined) return val.value;
        if (val.index !== undefined) return val.index;
        for (const k of Object.keys(val)) {
          if (typeof val[k] === 'number') return val[k];
        }
      }
      return '';
    };

    const sensor = (val) => (val === undefined || val === null || val === 255) ? '' : extract(val);

    const windRose = Object.keys(iss.wind_rose || {}).slice(-1)[0] || '';

    const row = [
      new Date(iss.lastReceived || Date.now()).toISOString(),
      fToC(extract(iss.temp_out)), fToC(extract(iss.temp_in)),
      extract(iss.hum_out), extract(iss.hum_in),
      fToC(extract(iss.temp_heat)), fToC(extract(iss.thswIndex)),
      fToC(extract(iss.temp_chill)), fToC(extract(iss.thwIndex)),
      fToC(extract(iss.temp_dew)), fToC(extract(iss.temp_wet_bulb)),
      mphToKmh(extract(iss.wind_speed)), extract(iss.wind_direction), windRose,
      mphToKmh(extract(iss.iWindSpeed2Min)), mphToKmh(extract(iss.iWindSpeed10Min)),
      mphToKmh(extract(iss.iWindGust10Min)),
      inHgToHpa(extract(bar.barCurrent) || extract(iss.barometer)),
      iss.barTrendText || bar.barTrend || '', extract(iss.barTrendVal),
      extract(iss.solar_rad), extract(iss.uv),
      inToMm(extract(iss.rain_rate)), inToMm(extract(iss.rain_day)),
      inToMm(extract(iss.rain_month)), inToMm(extract(iss.rain_year)),
      inToMm(extract(iss.rain_storm)), inToMm(extract(iss.iRain1Hour)),
      inToMm(extract(iss.et_day)), inToMm(extract(iss.et_month)), inToMm(extract(iss.et_year)),
      fToC(extract(hilow.temp_out_hi)), formatTime(hilow.temp_out_hi_time),
      fToC(extract(hilow.temp_out_low)), formatTime(hilow.temp_out_low_time),
      extract(hilow.hum_out_hi), formatTime(hilow.hum_out_hi_time),
      extract(hilow.hum_out_low), formatTime(hilow.hum_out_low_time),
      fToC(extract(hilow.dew_hi)), formatTime(hilow.dew_hi_time),
      fToC(extract(hilow.dew_low)), formatTime(hilow.dew_low_time),
      fToC(extract(hilow.heat_hi)), formatTime(hilow.heat_hi_time),
      fToC(extract(hilow.chill_low)), formatTime(hilow.chill_low_time),
      mphToKmh(extract(hilow.wind_speed_hi)), formatTime(hilow.wind_speed_hi_time),
      extract(hilow.solar_rad_hi), formatTime(hilow.solar_rad_hi_time),
      inHgToHpa(extract(hilow.bar_day_high)), formatTime(hilow.bar_day_high_time),
      inHgToHpa(extract(hilow.bar_day_low)), formatTime(hilow.bar_day_low_time),
      fToC(extract(hilow.wet_bulb_low)), formatTime(hilow.wet_bulb_low_time),
      data.sunrise || '', data.sunset || '', extract(data.moonPhase),
      fToC(sensor(iss.extraTemp1)), sensor(iss.extraHum1),
      fToC(sensor(iss.extraTemp2)), sensor(iss.extraHum2),
      fToC(sensor(iss.extraTemp3)), sensor(iss.extraHum3),
      fToC(sensor(iss.extraTemp4)), sensor(iss.extraHum4),
      fToC(sensor(iss.extraTemp5)), sensor(iss.extraHum5),
      fToC(sensor(iss.extraTemp6)), sensor(iss.extraHum6),
      fToC(sensor(iss.extraTemp7)), sensor(iss.extraHum7),
      fToC(sensor(iss.soilTemp1)), sensor(iss.soilMoist1),
      fToC(sensor(iss.soilTemp2)), sensor(iss.soilMoist2),
      fToC(sensor(iss.soilTemp3)), sensor(iss.soilMoist3),
      fToC(sensor(iss.soilTemp4)), sensor(iss.soilMoist4),
      fToC(sensor(iss.leafTemp1)), sensor(iss.leafWet1),
      fToC(sensor(iss.leafTemp2)), sensor(iss.leafWet2),
      fToC(sensor(iss.leafTemp3)), sensor(iss.leafWet3),
      fToC(sensor(iss.leafTemp4)), sensor(iss.leafWet4)
    ];

    sheet.appendRow(row);
    logRun(logSheet, sheetName, 'SUCCESS', '');

  } catch (err) {
    logRun(logSheet, sheetName, 'FAILED', err.toString());
  }
}

// Fetch hobolink Data

function fetchHobolinkData(logSheet) {
  if (!HOBOLINK_TOKEN) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(HOBO_SHEET_NAME);

  const headers = [
    "Timestamp (UTC)",
    "Temperature (°C)", "RH (%)", "Dew Point (°C)",
    "Heat Index (°C)", "Wet Bulb (°C)",
    "Wind Speed (m/s)", "Wind Speed (km/h)", "Gust Speed (m/s)",
    "Wind Direction (°)",
    "Pressure (mbar)",
    "Rain (mm)",
    "Water Content (m³/m³)"
  ];

  if (!sheet) {
    sheet = ss.insertSheet(HOBO_SHEET_NAME);
    sheet.appendRow(headers);
  }

  try {
    // Last stored timestamp
    let lastTs = 0;
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const lastTime = sheet.getRange(lastRow, 1).getValue();
      if (lastTime) lastTs = new Date(lastTime).getTime();
    }
    if (lastTs === 0) lastTs = Date.now() - HOBO_LOOKBACK_HOURS * 3600000;

    const url = `https://api.licor.cloud/v2/data?deviceSerialNumber=${HOBO_DEVICE_SERIAL}&startTime=${lastTs}&endTime=${Date.now()}`;

    const resp = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': `Bearer ${HOBOLINK_TOKEN}`, 'Accept': 'application/json' },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      logRun(logSheet, HOBO_SHEET_NAME, 'FAILED', `HTTP ${resp.getResponseCode()}`);
      return;
    }

    const data = JSON.parse(resp.getContentText());
    if (!data.sensors?.length) return;

    // Build readings map: timestamp -> { measurementType: value }
    const readings = {};
    data.sensors.forEach(sensor => {
      const type = sensor.data?.[0]?.measurementType;
      const records = sensor.data?.[0]?.records || [];
      if (!type) return;

      records.forEach(([ts, val]) => {
        if (!readings[ts]) readings[ts] = {};
        readings[ts][type] = val;
      });
    });

    // Append new rows
    const timestamps = Object.keys(readings).map(Number).sort((a, b) => a - b);
    let count = 0;

    timestamps.forEach(ts => {
      if (ts <= lastTs) return;

      const r = readings[ts];
      const temp = r.Temperature ?? null;
      const rh = r.RH ?? null;
      const pressure = r.Pressure ?? null;
      const windSpeed = r['Wind Speed'] ?? null;

      sheet.appendRow([
        new Date(ts + (-1 * parseInt(TIMEZONE_OFFSET))).toISOString(),
        roundVal(temp),
        roundVal(rh),
        roundVal(r['Dew Point']),
        calcHeatIndex(temp, rh),
        calcWetBulb(temp, rh, pressure),
        roundVal(windSpeed),
        windSpeed !== null ? roundVal(windSpeed * 3.6) : '',
        roundVal(r['Gust Speed']),
        roundVal(r['Wind Direction']),
        roundVal(pressure),
        roundVal(r.Rain),
        roundVal(r['Water Content'])
      ]);
      count++;
    });

    if (count > 0) {
      Logger.log(`[HOBO] ${count} new rows`);
      logRun(logSheet, HOBO_SHEET_NAME, 'SUCCESS', `${count} rows`);
    }

  } catch (err) {
    Logger.log(`[HOBO] ${err}`);
    try { logRun(logSheet, HOBO_SHEET_NAME, 'FAILED', err.toString()); } catch(e) {}
  }
}

// field calculations

function calcHeatIndex(tempC, rh) {
  if (tempC === null || rh === null) return '';
  const t = tempC * 9/5 + 32;
  if (t < 80) return roundVal(tempC);

  const hi = -42.379 + 2.04901523*t + 10.14333127*rh
    - 0.22475541*t*rh - 0.00683783*t*t - 0.05481717*rh*rh
    + 0.00122874*t*t*rh + 0.00085282*t*rh*rh - 0.00000199*t*t*rh*rh;

  return roundVal((hi - 32) * 5/9);
}

function calcWetBulb(tempC, rh, _pressure) {
  if (tempC === null || rh === null) return '';
  const tw = tempC * Math.atan(0.151977 * Math.sqrt(rh + 8.313659))
    + Math.atan(tempC + rh) - Math.atan(rh - 1.676331)
    + 0.00391838 * Math.pow(rh, 1.5) * Math.atan(0.023101 * rh)
    - 4.686035;
  return roundVal(tw);
}

function roundVal(val) {
  if (val === undefined || val === null) return '';
  return Math.round(parseFloat(val) * 100) / 100;
}

    // --- helpers ---

function extractCookies(setCookie) {
  if (!setCookie) return [];
  return (Array.isArray(setCookie) ? setCookie : [setCookie])
    .map(c => c.split(';')[0].trim())
    .filter(Boolean);
}

function getLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName('Log');
  if (!log) {
    log = ss.insertSheet('Log');
    log.appendRow(['Timestamp (UTC)', 'Sheet', 'Status', 'Details']);
  }
  return log;
}

function logRun(logSheet, sheet, status, details) {
  try {
    logSheet.appendRow([new Date().toISOString(), sheet, status, details]);
  } catch(e) {
    Logger.log(`Log error: ${e}`);
  }
}







