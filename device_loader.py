import json
import os
from openpyxl import load_workbook

DEVICES_JSON = "devices.json"
DEVICES_XLSX = "devices.xlsx"


def load_devices():
    if os.path.exists(DEVICES_JSON) and os.path.getsize(DEVICES_JSON) > 0:
        try:
            with open(DEVICES_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except json.JSONDecodeError:
            pass
    return []


def load_devices_from_excel(xlsx_file=DEVICES_XLSX):
    devices = []

    if not os.path.exists(xlsx_file):
        return devices

    wb = load_workbook(xlsx_file)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):
        name, ip = row[:2]
        if name and ip:
            devices.append({
                "name": str(name),
                "ip": str(ip)
            })

    return devices


def save_devices(devices):
    with open(DEVICES_JSON, "w", encoding="utf-8") as f:
        json.dump(devices, f, indent=2, ensure_ascii=False)