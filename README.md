# HHD Performance Testing Tool

A small Flask web app for capturing performance metrics of the Godam/WMS Android
apps on physical HHD devices (Newland, Urovo, Chainway, etc.) over `adb`, and
recording them to a **Google Sheet** or a **local Excel file**.

Metrics captured per app: CPU usage, memory (PSS), app launch time, FPS,
network latency, crash/ANR counts, a manual rating, and free-text remarks.

---

## Requirements

- Python 3
- `adb` on `PATH` (Android platform-tools)
- A device connected and authorised for USB debugging
- For Google mode: a Google service-account JSON with edit access to the target sheet

## Setup

```bash
python3 -m venv venv
./venv/bin/pip install -r requirements.txt
```

---

## Running

The tool supports two modes, selected with the `PERF_MODE` environment variable.
`google` is the default.

### Google mode (writes to a shared Google Sheet)

```bash
PERF_MODE=google ./venv/bin/python app.py
```

Then open http://localhost:5000

### Local mode (writes to a local .xlsx file)

```bash
PERF_MODE=local ./venv/bin/python app.py
```

Local mode writes to `../Godam_HHD_Performance_Testing.xlsx` (one worksheet per
device, named `"<manufacturer> <model>"`).

> Tip: use **local mode** for scratch/practice runs so you don't write test data
> into the shared Google Sheet. Switch to **google mode** for the real capture.

---

## Configuration (environment variables)

| Variable       | Default              | Purpose |
|----------------|----------------------|---------|
| `PERF_MODE`    | `google`             | `google` writes to Google Sheets; `local` writes to the local `.xlsx`. |
| `GSHEET_ID`    | (see `app.py`)       | Target Google Sheet ID (google mode only). |
| `GSHEET_CREDS` | (see `app.py`)       | Path to the Google service-account credentials JSON (google mode only). |

Examples:

```bash
# Google mode against a different sheet
GSHEET_ID=your_sheet_id PERF_MODE=google ./venv/bin/python app.py

# Google mode with custom credentials path
GSHEET_CREDS=/path/to/creds.json PERF_MODE=google ./venv/bin/python app.py
```

### Google mode setup (one-time)

1. Create a Google Cloud service account and download its JSON key.
2. Enable the Google Sheets API for the project.
3. Share the target spreadsheet with the service account's email
   (`client_email` in the JSON) as an **Editor**.
4. Point `GSHEET_CREDS` at the JSON file.

> The credentials JSON is **never** committed — it is excluded by `.gitignore`.

---

## How to capture metrics

1. Connect the HHD via USB and confirm it shows up: `adb devices`.
2. Start the tool (see Running above) and open http://localhost:5000.
3. Select the **device** from the dropdown (auto-detected via adb).
4. Select the **application** (Godam child app) to test.
5. Click **Run All Metrics** (or run each tile individually: CPU, Memory,
   Launch Time, FPS, Network).
6. Add an optional **rating** and **remarks**.
7. Click **Save** — the row is appended to the device's worksheet
   (Google Sheet or local `.xlsx` depending on mode).
8. **Open Sheet** opens the Google Sheet URL (google mode) or the local file
   (local mode).

### Notes on launch time

Launch time is measured with `am start-activity -W`. The tool uses `TotalTime`
when available, and falls back to `WaitTime` for apps that redirect to another
process on launch (e.g. Picking hands off to the Godam `LoginActivity`, which
makes Android report `LaunchState: UNKNOWN` and omit `TotalTime`). When the
`WaitTime` fallback is used, the reading carries a note to that effect.

---

## Endpoints (for reference)

| Endpoint                | Purpose |
|-------------------------|---------|
| `GET /api/devices`      | List connected adb devices. |
| `GET /api/packages`     | List which Godam apps are installed on the device. |
| `GET /api/metric/cpu`   | CPU usage %. |
| `GET /api/metric/memory`| Memory (TOTAL PSS) in MB. |
| `GET /api/metric/launch`| App launch time in ms (TotalTime, falls back to WaitTime). |
| `GET /api/metric/fps`   | Frames per second. |
| `GET /api/metric/network`| Network latency in ms. |
| `GET /api/launch`       | Launch the app (monkey). |
| `GET /api/uninstall-all`| Uninstall all `com.delhivery.*` apps except the Godam drawer. |
| `POST /api/save`        | Save a metrics row (routes to Google or local per `PERF_MODE`). |
| `GET /api/open-sheet`   | Open the Google Sheet URL or local file. |
| `GET /api/mode`         | Return the current mode. |

---

## Troubleshooting

- **Device not listed**: run `adb devices`; accept the USB debugging prompt on
  the device; re-plug if it shows `unauthorized`/`offline`.
- **Launch time N/A**: the tool falls back to `WaitTime` for apps that redirect
  on launch. If still N/A, the app failed to start; check adb.
- **Google save fails**: confirm the service account has Editor access to the
  sheet and `GSHEET_CREDS`/`GSHEET_ID` are correct.
- **Port 5000 in use**: stop the other process, or change the port in `app.py`.
