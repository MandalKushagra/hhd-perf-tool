# HHD Performance Testing Tool

Web-based tool for measuring Android app performance metrics on warehouse HHD (Handheld Device) hardware. Connects to devices over ADB, runs benchmarks, and saves results to a structured Excel sheet.

![Python](https://img.shields.io/badge/python-3.8+-3776ab?style=flat-square) ![Flask](https://img.shields.io/badge/flask-3.x-000?style=flat-square)

## What It Does

- Auto-detects connected ADB devices (manufacturer, model, Android version)
- Lists installed Godam/WMS apps on each device
- Measures per-app metrics:
  - **CPU Usage** (%) via `top` / `dumpsys cpuinfo`
  - **Memory Usage** (MB) via `dumpsys meminfo` (PSS)
  - **App Launch Time** (ms) cold start via `am start -W`
  - **FPS** via `dumpsys gfxinfo` jank frame analysis
  - **Network Latency** (ms) ping to `api-wms.delhivery.com`
- Saves results directly to a formatted `.xlsx` file (one sheet per device)
- Kill/Launch apps remotely
- Bulk uninstall all Delhivery apps from a device
- Real-time terminal output panel showing ADB commands being run

## Prerequisites

| Dependency | Version | Install |
|---|---|---|
| Python | 3.8+ | `sudo apt install python3 python3-pip python3-venv` |
| ADB | any | `sudo apt install adb` or via Android SDK platform-tools |
| USB debugging | - | Enable on each device: Settings > Developer Options > USB Debugging |

Verify ADB sees your device:
```bash
adb devices
# Should list your device serial as "device"
```

## Quick Start

```bash
# 1. Clone
git clone https://github.com/MandalKushagra/hhd-perf-tool.git
cd hhd-perf-tool

# 2. Set up virtualenv
python3 -m venv venv
source venv/bin/activate

# 3. Install dependencies
pip install -r requirements.txt

# 4. Generate the Excel sheet (first time only)
python3 generate_sheet.py

# 5. Start the tool
python3 app.py
```

Open **http://localhost:5000** in your browser.

## Generating the Excel Sheet

The tool writes metric data to an Excel file. Generate it before first use:

```bash
python3 generate_sheet.py
```

This creates `Godam_HHD_Performance_Testing.xlsx` in the project root with:
- A **Summary** sheet
- One sheet per device (pre-configured for 5 devices)
- Formatted headers: S/N, App Name, Package, CPU, Memory, Launch Time, FPS, Network, Crashes, ANRs, Rating, Remarks

To customize devices, edit the `devices` list in `generate_sheet.py`:

```python
devices = [
    ("Newland NLS-MT93L", "Newland", "NLS-MT93L", "YOUR_SERIAL", "13"),
    ("Chainway C66", "Chainway", "C66", "YOUR_SERIAL", "13"),
    # Add more devices as needed...
]
```

## Usage

1. Connect device(s) via USB with debugging enabled
2. Select a device from the dropdown
3. Select an app to benchmark
4. Click **Run All Metrics** or run individual metrics
5. Add a rating (1-5) and optional remarks
6. Click **Save to Excel** to persist the row

You can also:
- **Kill App** / **Launch App** for quick app control
- **Open Sheet** to view the Excel file
- **Uninstall All Delhivery Apps** for a clean slate

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `PERF_EXCEL_PATH` | `./Godam_HHD_Performance_Testing.xlsx` | Path to the Excel file |

```bash
PERF_EXCEL_PATH=/path/to/my/sheet.xlsx python3 app.py
```

## Project Structure

```
hhd-perf-tool/
├── app.py                 # Flask backend (ADB + Excel I/O)
├── static/
│   └── index.html         # Frontend UI (single-page, no build step)
├── generate_sheet.py      # Excel sheet generator
├── requirements.txt       # Python dependencies
└── .gitignore
```

## Supported Apps

| App | Package |
|---|---|
| Store (Godam Drawer) | `com.delhivery.godam.drawer` |
| WMS | `com.delhivery.mobile.godam.wms` |
| Picking | `com.delhivery.mobile.godam.picking` |
| Put | `com.delhivery.mobile.godam.put` |
| Pack | `com.delhivery.mobile.godam.pack` |
| Transfer | `com.delhivery.mobile.godam.transfers` |
| Receiving | `com.delhivery.mobile.godam.receiving` |
| CycleCount | `com.delhivery.cyclecount` |
| Box | `com.delhivery.mobile.godam.box` |
| Dispatch | `com.delhivery.mobile.godam.dispatch` |
| ContainerInfo | `com.delhivery.mobile.godam.containers` |
| Rapid Picking | `com.delhivery.darkstore.pick` |

To add/remove apps, edit `GODAM_PACKAGES` in `app.py`.

## Notes

- ADB commands timeout after 15 seconds
- Network latency is measured from the device (not your machine)
- FPS measurement takes ~3 seconds (waits for frame data)
- Launch time does a cold start (force-stops app first)
- "Open Sheet" uses `xdg-open` (Linux). On macOS use `open`, on Windows use `start`
- Works with multiple devices connected simultaneously

## Troubleshooting

| Problem | Fix |
|---|---|
| No devices showing up | Run `adb devices`, check USB cable, enable USB debugging |
| "App may not be running" for CPU | Launch the app first, then measure |
| Network latency fails | Device needs internet. Connect to WiFi |
| Excel save fails | Run `generate_sheet.py` first |
| Permission denied on ADB | `adb kill-server && adb start-server`, re-authorize on device |

## License

Internal tool. Use as needed.
