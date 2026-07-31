#!/usr/bin/env python3
"""
Generate the Excel sheet required for HHD performance data capture.

Usage:
    python3 generate_sheet.py

Output:
    Creates Godam_HHD_Performance_Testing.xlsx in the current directory.
    Each device gets its own sheet with pre-formatted headers.

Edit the `devices` list below to match your test devices.
"""
import os
import string
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


# ============================================================
# CONFIGURE YOUR DEVICES HERE
# Format: (sheet_name, manufacturer, model, serial, android_version)
# ============================================================
devices = [
    ("Newland NLS-MT93L", "Newland", "NLS-MT93L", "", "13"),
    ("Newland NLS-WD1", "Newland", "NLS-WD1", "", "13"),
    ("Chainway C66", "Chainway", "C66", "", "13"),
    ("Urovo CT58", "Urovo", "CT58", "", "12"),
    ("Urovo DT50S", "Urovo", "DT50S", "", "13"),
]

# Column headers for each device sheet
headers = [
    "S/N", "App Name", "Package", "CPU Usage (%)", "Memory Usage (MB)",
    "App Launch Time (ms)", "FPS", "Network Latency (ms)",
    "Crash Count", "ANR Count", "Rating (1-5)", "Remarks"
]

# ============================================================
# STYLES
# ============================================================
header_font = Font(bold=True, size=11, color="FFFFFF")
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
device_info_font = Font(bold=True, size=12, color="4FC3F7")
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
col_widths = [5, 20, 38, 14, 18, 20, 8, 18, 12, 12, 12, 35]


def generate():
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    for sheet_name, manufacturer, model, serial, android_ver in devices:
        ws = wb.create_sheet(title=sheet_name)

        # Device info header (rows 1-2)
        ws.cell(row=1, column=1, value="Manufacturer:").font = Font(bold=True)
        ws.cell(row=1, column=2, value=manufacturer).font = device_info_font
        ws.cell(row=1, column=3, value="Model:").font = Font(bold=True)
        ws.cell(row=1, column=4, value=model).font = device_info_font
        ws.cell(row=2, column=1, value="Serial:").font = Font(bold=True)
        ws.cell(row=2, column=2, value=serial if serial else "(fill in)").font = device_info_font
        ws.cell(row=2, column=3, value="Android:").font = Font(bold=True)
        ws.cell(row=2, column=4, value=android_ver).font = device_info_font

        # Column headers at row 4
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border

        # Set column widths
        for i, w in enumerate(col_widths):
            col_letter = string.ascii_uppercase[i] if i < 26 else 'A'
            ws.column_dimensions[col_letter].width = w

    # Summary sheet (first tab)
    ws_summary = wb.create_sheet(title="Summary", index=0)
    summary_headers = ["Device", "Manufacturer", "Model", "Android", "Avg Rating", "Total Apps Tested", "Overall Remarks"]
    ws_summary.append(summary_headers)
    for sheet_name, manufacturer, model, serial, android_ver in devices:
        ws_summary.append([sheet_name, manufacturer, model, android_ver, "", "", ""])
    for cell in ws_summary[1]:
        cell.font = header_font
        cell.fill = header_fill

    # Save
    output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Godam_HHD_Performance_Testing.xlsx")
    wb.save(output_path)
    print(f"Created: {output_path}")
    print(f"Sheets: {wb.sheetnames}")
    print(f"\nDone. Open the file and start collecting metrics with: python3 app.py")


if __name__ == "__main__":
    generate()
