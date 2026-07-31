"""
HHD Performance Testing Tool
Run: python3 app.py
Open: http://localhost:5000
"""
import subprocess
import json
import re
import os
from flask import Flask, jsonify, request, send_from_directory
from openpyxl import load_workbook
from datetime import datetime

app = Flask(__name__, static_folder='static')

EXCEL_PATH = os.environ.get('PERF_EXCEL_PATH', os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Godam_HHD_Performance_Testing.xlsx'))

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
    cmd = ['adb', '-s', device_id] + list(args)
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
    result = subprocess.run(['adb', 'devices'], capture_output=True, text=True)
    devices = []
    for line in result.stdout.strip().split('\n')[1:]:
        if 'device' in line and 'List' not in line and 'offline' not in line:
            parts = line.split()
            serial = parts[0]
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
        result.append({'name': name, 'package': pkg, 'installed': pkg in installed})
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
    run_adb(device, 'shell', 'am', 'force-stop', package)
    import time
    time.sleep(1)
    output = run_adb(device, 'shell', 'cmd', 'package', 'resolve-activity', '--brief', package)
    activity = None
    for line in output.split('\n'):
        if '/' in line and package in line:
            activity = line.strip()
            break
    if not activity:
        output = run_adb(device, 'shell', 'am', 'start-activity', '-W', '-n',
                        f'{package}/.ui.activity.SplashActivity')
        if 'TotalTime' not in output:
            output = run_adb(device, 'shell', 'monkey', '-p', package,
                           '-c', 'android.intent.category.LAUNCHER', '1')
            return jsonify({'value': -1, 'unit': 'ms', 'note': 'Used monkey, no timing available'})
    else:
        output = run_adb(device, 'shell', 'am', 'start-activity', '-W', '-n', activity)
    match = re.search(r'TotalTime:\s*(\d+)', output)
    if match:
        return jsonify({'value': int(match.group(1)), 'unit': 'ms'})
    return jsonify({'value': -1, 'unit': 'ms', 'note': 'Could not measure'})


@app.route('/api/metric/fps')
def get_fps():
    device = request.args.get('device')
    package = request.args.get('package')
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
    fps = 60
    if total_frames > 0 and janky > 0:
        jank_pct = (janky / total_frames) * 100
        fps = max(1, int(60 * (1 - jank_pct / 100)))
    return jsonify({'value': fps, 'unit': 'fps', 'totalFrames': total_frames, 'janky': janky})


@app.route('/api/metric/network')
def get_network():
    device = request.args.get('device')
    output = run_adb(device, 'shell', 'ping', '-c', '3', '-W', '3', 'api-wms.delhivery.com')
    match = re.search(r'min/avg/max.*?=\s*([\d.]+)/([\d.]+)/([\d.]+)', output)
    if match:
        return jsonify({'value': float(match.group(2)), 'unit': 'ms'})
    match2 = re.search(r'(\d+\.\d+)/(\d+\.\d+)/(\d+\.\d+)', output)
    if match2:
        return jsonify({'value': float(match2.group(2)), 'unit': 'ms'})
    output2 = run_adb(device, 'shell', 'curl', '-o', '/dev/null', '-s', '-w', '%{time_connect}', 'https://api-wms.delhivery.com')
    try:
        seconds = float(output2.strip())
        return jsonify({'value': round(seconds * 1000, 1), 'unit': 'ms', 'note': 'Measured via curl connect time'})
    except:
        pass
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
    """Save a metric reading to the correct device sheet in the Excel file."""
    data = request.json
    wb = load_workbook(EXCEL_PATH)
    model = data.get('model', '')
    manufacturer = data.get('manufacturer', '')
    sheet_name = f"{manufacturer} {model}"
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
    return jsonify({'success': True, 'row': next_row, 'sheet': ws.title})


@app.route('/api/open-sheet')
def open_sheet():
    """Open the Excel file with the default application."""
    subprocess.Popen(['xdg-open', os.path.abspath(EXCEL_PATH)])
    return jsonify({'success': True})


if __name__ == '__main__':
    print("\n\U0001f527 HHD Performance Testing Tool")
    print(f"\U0001f4ca Excel: {os.path.abspath(EXCEL_PATH)}")
    print(f"\U0001f310 Open: http://localhost:5000\n")
    app.run(debug=True, port=5000)
