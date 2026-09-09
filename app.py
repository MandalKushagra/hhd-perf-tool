"""
HHD Performance Testing Tool
Run: python app.py   (Windows)  /  python3 app.py  (Linux/macOS)
Open: http://localhost:5000

Supports two modes (set via PERF_MODE env var):
  - "google" (default): writes directly to Google Sheets
  - "local": writes to a local .xlsx file
"""
import subprocess
import json
import re
import os
import sys
from flask import Flask, jsonify, request, send_from_directory
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__, static_folder='static')

# --- Configuration ---
PERF_MODE = os.environ.get('PERF_MODE', 'google')  # "google" or "local"
EXCEL_PATH = os.path.join(os.path.dirname(__file__), '..', 'Godam_HHD_Performance_Testing.xlsx')
GSHEET_ID = os.environ.get('GSHEET_ID', '1wOkihKil908xK5HfpQL2a1hUvBP5uOJsZY8K4paxkeo')
GSHEET_CREDS = os.environ.get('GSHEET_CREDS', '/home/kushagra/Downloads/model-obelisk-465116-v2-a9472f70816f.json')

# adb executable name. On Windows the binary is adb.exe; letting the OS resolve
# it from PATH works for both, but being explicit avoids edge cases.
ADB = 'adb.exe' if sys.platform == 'win32' else 'adb'


def open_path(target):
    """Open a file path or URL in the OS default handler, cross-platform.

    Windows uses os.startfile, macOS uses `open`, and other POSIX systems use
    `xdg-open`. Returns True on a best-effort launch, False on failure.
    """
    try:
        if sys.platform == 'win32':
            os.startfile(target)  # type: ignore[attr-defined]
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', target])
        else:
            subprocess.Popen(['xdg-open', target])
        return True
    except Exception:
        return False


# --- Google Sheets client (lazy init) ---
_gsheet_client = None

def get_gsheet():
    """Get or create the gspread client and open the spreadsheet."""
    global _gsheet_client
    if _gsheet_client is None:
        import gspread
        _gsheet_client = gspread.service_account(filename=GSHEET_CREDS)
    return _gsheet_client.open_by_key(GSHEET_ID)

GODAM_PACKAGES = [
    ("Store (Godam Drawer)", "com.delhivery.godam.drawer"),
    ("WMS", "com.delhivery.mobile.godam.wms"),
    ("Picking", "com.delhivery.mobile.godam.picking"),
    ("Put", "com.delhivery.mobile.godam.put"),
    ("Pack", "com.delhivery.mobile.godam.pack"),
    ("Transfer", "com.delhivery.mobile.godam.transfers"),
    ("Receiving", "com.delhivery.mobile.godam.receiving"),
    ("CycleCount", "com.delhivery.cyclecount"),
    ("Box", "com.delhivery.mobile.godam.box"),
    ("Dispatch", "com.delhivery.mobile.godam.dispatch"),
    ("ContainerInfo", "com.delhivery.mobile.godam.containers"),
    ("Rapid Picking", "com.delhivery.darkstore.pick"),
]


def run_adb(device_id, *args):
    """Run an adb command targeting a specific device."""
    cmd = [ADB, '-s', device_id] + list(args)
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=15)
        return result.stdout.strip()
    except Exception as e:
        return f"ERROR: {e}"


@app.route('/')
def index():
    return send_from_directory('static', 'index.html')


@app.route('/api/devices')
def get_devices():
    """List connected adb devices with info."""
    result = subprocess.run([ADB, 'devices'], capture_output=True, text=True)
    devices = []
    for line in result.stdout.strip().split('\n')[1:]:
        if 'device' in line and 'List' not in line and 'offline' not in line:
            parts = line.split()
            serial = parts[0]
            # Get device model and android version
            model = run_adb(serial, 'shell', 'getprop', 'ro.product.model')
            manufacturer = run_adb(serial, 'shell', 'getprop', 'ro.product.manufacturer')
            android_ver = run_adb(serial, 'shell', 'getprop', 'ro.build.version.release')
            devices.append({
                'serial': serial,
                'model': model,
                'manufacturer': manufacturer,
                'android': android_ver,
                'label': f"{manufacturer} {model} ({serial})"
            })
    return jsonify(devices)


@app.route('/api/packages')
def get_packages():
    """List Godam packages installed on device."""
    device = request.args.get('device')
    if not device:
        return jsonify([])
    
    output = run_adb(device, 'shell', 'pm', 'list', 'packages')
    installed = set()
    for line in output.split('\n'):
        pkg = line.replace('package:', '').strip()
        if pkg:
            installed.add(pkg)
    
    result = []
    for name, pkg in GODAM_PACKAGES:
        if pkg in installed:
            result.append({'name': name, 'package': pkg, 'installed': True})
        else:
            result.append({'name': name, 'package': pkg, 'installed': False})
    return jsonify(result)


@app.route('/api/metric/cpu')
def get_cpu():
    device = request.args.get('device')
    package = request.args.get('package')
    output = run_adb(device, 'shell', 'top', '-n', '1', '-b')
    for line in output.split('\n'):
        if package in line:
            parts = line.split()
            for i, p in enumerate(parts):
                if '%' in p or (i > 0 and parts[0].isdigit()):
                    try:
                        cpu = float(re.sub(r'[^0-9.]', '', parts[8] if len(parts) > 8 else '0'))
                        return jsonify({'value': cpu, 'unit': '%'})
                    except:
                        pass
    # Fallback: use dumpsys cpuinfo
    output2 = run_adb(device, 'shell', 'dumpsys', 'cpuinfo')
    for line in output2.split('\n'):
        if package in line:
            match = re.search(r'([\d.]+)%', line)
            if match:
                return jsonify({'value': float(match.group(1)), 'unit': '%'})
    return jsonify({'value': 0, 'unit': '%', 'note': 'App may not be running'})


@app.route('/api/metric/memory')
def get_memory():
    device = request.args.get('device')
    package = request.args.get('package')
    output = run_adb(device, 'shell', 'dumpsys', 'meminfo', package)
    for line in output.split('\n'):
        if 'TOTAL PSS' in line or 'TOTAL' in line:
            match = re.search(r'TOTAL\s+PSS:\s*([\d,]+)', line)
            if match:
                kb = int(match.group(1).replace(',', ''))
                return jsonify({'value': round(kb / 1024, 1), 'unit': 'MB'})
            match2 = re.search(r'TOTAL\s+([\d,]+)', line)
            if match2:
                kb = int(match2.group(1).replace(',', ''))
                return jsonify({'value': round(kb / 1024, 1), 'unit': 'MB'})
    return jsonify({'value': 0, 'unit': 'MB', 'note': 'Could not read'})


@app.route('/api/metric/launch')
def get_launch_time():
    device = request.args.get('device')
    package = request.args.get('package')
    # First stop the app
    run_adb(device, 'shell', 'am', 'force-stop', package)
    import time
    time.sleep(1)
    # Get launch activity
    output = run_adb(device, 'shell', 'cmd', 'package', 'resolve-activity', '--brief', package)
    activity = None
    for line in output.split('\n'):
        if '/' in line and package in line:
            activity = line.strip()
            break
    if not activity:
        # Fallback: use monkey
        output = run_adb(device, 'shell', 'am', 'start-activity', '-W', '-n',
                        f'{package}/.ui.activity.SplashActivity')
        if 'TotalTime' not in output:
            output = run_adb(device, 'shell', 'monkey', '-p', package,
                           '-c', 'android.intent.category.LAUNCHER', '1')
            return jsonify({'value': -1, 'unit': 'ms', 'note': 'Used monkey, no timing available'})
    else:
        output = run_adb(device, 'shell', 'am', 'start-activity', '-W', '-n', activity)

    # Prefer TotalTime, but fall back to WaitTime. Apps that redirect to another
    # process on launch (e.g. Picking -> Godam LoginActivity) report
    # LaunchState UNKNOWN and omit TotalTime, emitting only WaitTime.
    match = re.search(r'TotalTime:\s*(\d+)', output)
    if match:
        return jsonify({'value': int(match.group(1)), 'unit': 'ms'})
    wait_match = re.search(r'WaitTime:\s*(\d+)', output)
    if wait_match:
        return jsonify({'value': int(wait_match.group(1)), 'unit': 'ms',
                        'note': 'WaitTime (app redirects on launch; TotalTime unavailable)'})
    return jsonify({'value': -1, 'unit': 'ms', 'note': 'Could not measure'})


@app.route('/api/metric/fps')
def get_fps():
    device = request.args.get('device')
    package = request.args.get('package')
    # Reset gfxinfo
    run_adb(device, 'shell', 'dumpsys', 'gfxinfo', package, 'reset')
    import time
    time.sleep(3)
    output = run_adb(device, 'shell', 'dumpsys', 'gfxinfo', package)
    
    total_frames = 0
    janky = 0
    for line in output.split('\n'):
        if 'Total frames rendered' in line:
            match = re.search(r'(\d+)', line)
            if match:
                total_frames = int(match.group(1))
        if 'Janky frames' in line:
            match = re.search(r'(\d+)', line)
            if match:
                janky = int(match.group(1))
    
    # Estimate FPS from frame data
    fps = 60  # default if idle
    if total_frames > 0 and janky > 0:
        jank_pct = (janky / total_frames) * 100
        fps = max(1, int(60 * (1 - jank_pct / 100)))
    
    return jsonify({'value': fps, 'unit': 'fps', 'totalFrames': total_frames, 'janky': janky})


@app.route('/api/metric/network')
def get_network():
    device = request.args.get('device')
    # Try ping first
    output = run_adb(device, 'shell', 'ping', '-c', '3', '-W', '3', 'api-wms.delhivery.com')
    match = re.search(r'min/avg/max.*?=\s*([\d.]+)/([\d.]+)/([\d.]+)', output)
    if match:
        return jsonify({'value': float(match.group(2)), 'unit': 'ms'})
    # Try alternate format
    match2 = re.search(r'(\d+\.\d+)/(\d+\.\d+)/(\d+\.\d+)', output)
    if match2:
        return jsonify({'value': float(match2.group(2)), 'unit': 'ms'})
    # Fallback: use curl timing
    output2 = run_adb(device, 'shell', 'curl', '-o', '/dev/null', '-s', '-w', '%{time_connect}', 'https://api-wms.delhivery.com')
    try:
        seconds = float(output2.strip())
        return jsonify({'value': round(seconds * 1000, 1), 'unit': 'ms', 'note': 'Measured via curl connect time'})
    except:
        pass
    # Last fallback: wget timing
    output3 = run_adb(device, 'shell', 'ping', '-c', '1', '8.8.8.8')
    match3 = re.search(r'time=(\d+\.?\d*)', output3)
    if match3:
        return jsonify({'value': float(match3.group(1)), 'unit': 'ms', 'note': 'Ping to 8.8.8.8'})
    return jsonify({'value': -1, 'unit': 'ms', 'note': 'All methods failed - check device network'})


@app.route('/api/kill')
def kill_app():
    device = request.args.get('device')
    package = request.args.get('package')
    output = run_adb(device, 'shell', 'am', 'force-stop', package)
    return jsonify({'success': True, 'output': output})


@app.route('/api/launch')
def launch_app():
    device = request.args.get('device')
    package = request.args.get('package')
    output = run_adb(device, 'shell', 'monkey', '-p', package, '-c', 'android.intent.category.LAUNCHER', '1')
    return jsonify({'success': True, 'output': output})


@app.route('/api/uninstall-all')
def uninstall_all():
    """Uninstall all com.delhivery apps except godam drawer."""
    device = request.args.get('device')
    if not device:
        return jsonify({'success': False, 'error': 'No device'})
    output = run_adb(device, 'shell', 'pm', 'list', 'packages')
    removed = []
    for line in output.split('\n'):
        pkg = line.replace('package:', '').strip()
        if 'delhivery' in pkg and pkg != 'com.delhivery.godam.drawer':
            run_adb(device, 'uninstall', pkg)
            removed.append(pkg)
    return jsonify({'success': True, 'removed': removed, 'count': len(removed)})


@app.route('/api/save', methods=['POST'])
def save_to_excel():
    """Save a metric reading — to Google Sheet or local Excel based on PERF_MODE."""
    data = request.json
    model = data.get('model', '')
    manufacturer = data.get('manufacturer', '')
    sheet_name = f"{manufacturer} {model}"

    if PERF_MODE == 'google':
        return _save_to_gsheet(data, model, sheet_name)
    else:
        return _save_to_local(data, model, sheet_name)


def _save_to_gsheet(data, model, sheet_name):
    """Write metrics to Google Sheet."""
    try:
        spreadsheet = get_gsheet()

        # Find or create worksheet
        ws = None
        for worksheet in spreadsheet.worksheets():
            if model in worksheet.title:
                ws = worksheet
                break

        if not ws:
            ws = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=12)
            headers = ["S/N", "App Name", "Package", "CPU Usage (%)", "Memory Usage (MB)",
                       "App Launch Time (ms)", "FPS", "Network Latency (ms)",
                       "Crash Count", "ANR Count", "Rating (1-5)", "Remarks"]
            ws.update('A4:L4', [headers])

        # Find next empty row (data starts at row 5)
        col_b = ws.col_values(2)
        next_row = max(len(col_b) + 1, 5)
        sn = next_row - 4

        row_data = [
            sn,
            data.get('appName', ''),
            data.get('package', ''),
            data.get('cpu'),
            data.get('memory'),
            data.get('launch'),
            data.get('fps'),
            data.get('network'),
            data.get('crashes', 0),
            data.get('anr', 0),
            data.get('rating'),
            data.get('remarks', '')
        ]
        ws.update(f'A{next_row}:L{next_row}', [row_data])
        return jsonify({'success': True, 'row': next_row, 'sheet': ws.title, 'mode': 'google'})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e), 'mode': 'google'}), 500


def _save_to_local(data, model, sheet_name):
    """Write metrics to local .xlsx file."""
    wb = load_workbook(EXCEL_PATH)
    ws = None
    for name in wb.sheetnames:
        if model in name:
            ws = wb[name]
            break
    if not ws:
        ws = wb.create_sheet(title=sheet_name)
        headers = ["S/N", "App Name", "Package", "CPU Usage (%)", "Memory Usage (MB)",
                   "App Launch Time (ms)", "FPS", "Network Latency (ms)",
                   "Crash Count", "ANR Count", "Rating (1-5)", "Remarks"]
        for col, h in enumerate(headers, 1):
            ws.cell(row=4, column=col, value=h)
    next_row = 5
    while ws.cell(row=next_row, column=2).value is not None:
        next_row += 1
    sn = next_row - 4
    ws.cell(row=next_row, column=1, value=sn)
    ws.cell(row=next_row, column=2, value=data.get('appName', ''))
    ws.cell(row=next_row, column=3, value=data.get('package', ''))
    ws.cell(row=next_row, column=4, value=data.get('cpu'))
    ws.cell(row=next_row, column=5, value=data.get('memory'))
    ws.cell(row=next_row, column=6, value=data.get('launch'))
    ws.cell(row=next_row, column=7, value=data.get('fps'))
    ws.cell(row=next_row, column=8, value=data.get('network'))
    ws.cell(row=next_row, column=9, value=data.get('crashes', 0))
    ws.cell(row=next_row, column=10, value=data.get('anr', 0))
    ws.cell(row=next_row, column=11, value=data.get('rating'))
    ws.cell(row=next_row, column=12, value=data.get('remarks', ''))
    wb.save(EXCEL_PATH)
    return jsonify({'success': True, 'row': next_row, 'sheet': ws.title, 'mode': 'local'})


@app.route('/api/open-sheet')
def open_sheet():
    """Open the sheet — Google Sheet URL or local file. Cross-platform."""
    if PERF_MODE == 'google':
        url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}"
        ok = open_path(url)
        return jsonify({'success': ok, 'url': url})
    else:
        ok = open_path(os.path.abspath(EXCEL_PATH))
        return jsonify({'success': ok})


@app.route('/api/mode')
def get_mode():
    """Return current save mode."""
    return jsonify({'mode': PERF_MODE, 'gsheet_id': GSHEET_ID if PERF_MODE == 'google' else None})


if __name__ == '__main__':
    print("\n🔧 HHD Performance Testing Tool")
    print(f"📋 Mode: {PERF_MODE.upper()}")
    if PERF_MODE == 'google':
        print(f"📊 Google Sheet: https://docs.google.com/spreadsheets/d/{GSHEET_ID}")
    else:
        print(f"📊 Excel: {os.path.abspath(EXCEL_PATH)}")
    print(f"🌐 Open: http://localhost:5000\n")
    app.run(debug=True, port=5000)
