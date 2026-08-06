# Setup Guide

## Prerequisites

- Google account with Google Sheets access
- Davis WeatherLink free account (with shared stations)
- HOBOlink account with API access (for RX3000 data)

## Installation

### 1. Create Google Sheet

Create a new Google Sheet (or use an existing one).

### 2. Open Apps Script

**Extensions → Apps Script**

### 3. Add the script

Copy the contents of `src/Code.gs` into the Apps Script editor.
Replace the default `myFunction()` with the entire script.

### 4. Store credentials

Run these functions **once each**, then delete them:

#### Davis WeatherLink

function setCredentials() {
  PropertiesService.getScriptProperties()
    .setProperty('WL_USERNAME', 'your_email_or_username')
    .setProperty('WL_PASSWORD', 'your_password');
    
}


#### HOBOlink

function setHoboCredentials() {
  PropertiesService.getScriptProperties()
    .setProperty('HOBOLINK_API_KEY', 'your_api_token');
}

### 5. Configure stations

Edit the DAVIS_STATIONS array with your own station UUIDs:

const DAVIS_STATIONS = [
  { uuid: 'your-uuid-here', sheetName: 'Station_Name' },
  // ...
];

### 6. Test

Run streamWeatherLinkNetwork() from the Apps Script editor.
Check the sheet for data.

### 7. Schedule

    In Apps Script, click the clock icon (Triggers)

    + Add Trigger

    Function: streamWeatherLinkNetwork

    Event: Time‑driven -> Minutes timer-> Every 15 minutes

    Save


---

## Customization

| Task | Where to edit |
|------|---------------|
| Add/remove Davis stations | `DAVIS_STATIONS` array in `Code.gs` |
| Change HOBOlink device | `HOBO_DEVICE_SERIAL` constant |
| Switch to imperial units | Remove conversion functions (`fToC`, etc.) |
| Change update frequency | Edit the time‑driven trigger interval |
| Add new HOBOlink sensors | Update `headers` array and `sheet.appendRow()` |

---

## How It Works

Update HOBO_DEVICE_SERIAL constant.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---------|--------------|-----|
| Davis login fails | Expired password | Re‑run `setCredentials()` |
| HOBOlink no data | Invalid token | Re‑run `setHoboCredentials()` |
| Missing columns | Old sheet format | Delete the sheet, re‑run |
| Duplicate HOBO rows | Timestamp tracking broken | Check `lastTs` logic |
| Trigger not firing | Quota exceeded | Check Apps Script dashboard |

---

