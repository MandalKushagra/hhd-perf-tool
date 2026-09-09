# Running the HHD Performance Testing Tool on Windows

This is a Flask app that talks to Android HHD devices over `adb` and records
metrics to a Google Sheet or a local Excel file. It runs on Windows — this guide
covers the Windows-specific setup steps and the couple of gotchas.

> The `windows-support` branch also includes a small code fix: the **Open Sheet**
> button now works on Windows (it previously called the Linux-only `xdg-open`).

---

## 1. Prerequisites

### Python 3
Install from [python.org](https://www.python.org/downloads/windows/). During
install, **tick "Add python.exe to PATH"**. Verify:

```powershell
python --version
```

> On Windows the command is `python`, not `python3`.

### adb (Android platform-tools)
1. Download **SDK Platform-Tools for Windows** from
   [developer.android.com/tools/releases/platform-tools](https://developer.android.com/tools/releases/platform-tools).
2. Extract it somewhere, e.g. `C:\platform-tools`.
3. Add that folder to your **PATH**:
   - Start menu ▸ "Edit the system environment variables" ▸ **Environment Variables**
   - Under **User variables**, select `Path` ▸ **Edit** ▸ **New** ▸ paste `C:\platform-tools`
   - OK out of all dialogs, then open a **new** terminal.
4. Verify:

```powershell
adb version
```

### HHD device drivers
Some rugged HHDs (Newland, Urovo, Chainway) need the OEM USB driver on Windows.
If `adb devices` shows nothing after connecting, install the vendor's Windows USB
driver, then re-plug. Enable **USB debugging** on the device and accept the
authorization prompt.

Verify the device is seen:

```powershell
adb devices
```

---

## 2. Get the Tool

```powershell
git clone <repo-url>
cd hhd-perf-tool
```

Or download the repo ZIP from GitHub (**Code ▸ Download ZIP**) and extract it.

---

## 3. Set Up the Virtual Environment

```powershell
python -m venv venv
venv\Scripts\pip install -r requirements.txt
```

> Windows uses `venv\Scripts\` (backslashes), whereas Linux/macOS use
> `venv/bin/`. You can also "activate" the venv first:
> `venv\Scripts\Activate.ps1` (PowerShell) or `venv\Scripts\activate.bat` (cmd),
> after which plain `python` and `pip` use the venv.
>
> If PowerShell blocks the activation script with an execution-policy error, run:
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once, then retry. You can
> skip activation entirely by calling `venv\Scripts\python` directly as shown below.

---

## 4. Run

The mode is chosen with the `PERF_MODE` environment variable. **Inline
`VAR=value command` syntax does NOT work on Windows** — set the variable first.

### Local mode (writes to a local .xlsx — best for practice runs)

PowerShell:

```powershell
$env:PERF_MODE = "local"
venv\Scripts\python app.py
```

Command Prompt (cmd.exe):

```cmd
set PERF_MODE=local
venv\Scripts\python app.py
```

### Google mode (writes to the shared Google Sheet — the default)

You need a Google service-account JSON (see the main README for the one-time
Google setup). Point `GSHEET_CREDS` at it using a Windows path.

PowerShell:

```powershell
$env:PERF_MODE = "google"
$env:GSHEET_CREDS = "C:\Users\<you>\Downloads\service-account.json"
# optional: target a different sheet
# $env:GSHEET_ID = "your_sheet_id"
venv\Scripts\python app.py
```

Command Prompt (cmd.exe):

```cmd
set PERF_MODE=google
set GSHEET_CREDS=C:\Users\<you>\Downloads\service-account.json
venv\Scripts\python app.py
```

Then open http://localhost:5000

> The default `GSHEET_CREDS` path in `app.py` is a Linux path from the original
> author's machine. On Windows you **must** override it with `GSHEET_CREDS` (as
> above) or edit the default in `app.py`.

---

## 5. Generate the Local Excel Sheet (local mode only)

Local mode writes to `..\Godam_HHD_Performance_Testing.xlsx` (one directory
above the repo). Create it first with:

```powershell
venv\Scripts\python generate_sheet.py
```

That produces `Godam_HHD_Performance_Testing.xlsx` in the repo folder. Edit the
`devices` list at the top of `generate_sheet.py` to match your test devices. Move
or copy the file to the location the app expects, or adjust `EXCEL_PATH` in
`app.py` if you prefer to keep it alongside the code.

---

## 6. Capture Metrics

Usage is identical to other platforms (see the main README):

1. Connect the HHD and confirm `adb devices` lists it.
2. Open http://localhost:5000.
3. Pick the device and the app, then **Run All Metrics** (or individual tiles).
4. Add a rating and remarks, then **Save**.
5. **Open Sheet** opens the Google Sheet (google mode) or the local `.xlsx`
   (local mode) — this now works on Windows.

> All the `adb shell` commands (`ping`, `top`, `dumpsys`, etc.) run **on the
> Android device**, so they behave the same regardless of your host OS.

---

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `python` not recognized | Reinstall Python with **"Add python.exe to PATH"** ticked, or use the full path. |
| `adb` not recognized | Add the platform-tools folder to PATH and open a new terminal. |
| `adb devices` empty / `unauthorized` | Install the HHD's OEM USB driver, enable USB debugging, accept the on-device prompt, re-plug. |
| PowerShell won't run `Activate.ps1` | `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`, or call `venv\Scripts\python` directly. |
| Google save fails | Ensure the service account has **Editor** access to the sheet and `GSHEET_CREDS`/`GSHEET_ID` are correct (Windows path for creds). |
| `PERF_MODE` seems ignored | You likely used inline `PERF_MODE=... python app.py`. Set it first with `$env:` (PowerShell) or `set` (cmd). |
| Port 5000 in use | Stop the other process or change the port in `app.py`. |
| Open Sheet does nothing | Make sure you're on the `windows-support` branch (or main once merged) which includes the cross-platform open fix. |

---

## Summary

- Install **Python** (on PATH) and **adb platform-tools** (on PATH).
- `python -m venv venv` then `venv\Scripts\pip install -r requirements.txt`.
- Set `PERF_MODE` (and `GSHEET_CREDS` for google mode) as separate commands,
  then `venv\Scripts\python app.py`.
- The one required override on Windows is `GSHEET_CREDS` (the default is a Linux path).
