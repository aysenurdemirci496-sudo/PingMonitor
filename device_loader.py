import json
import os
from datetime import datetime
from openpyxl import load_workbook

DEVICES_JSON = "devices.json"
DEVICES_XLSX = "devices.xlsx"


def load_devices():
    # JSON varsa ve doluysa oradan oku
    if os.path.exists(DEVICES_JSON) and os.path.getsize(DEVICES_JSON) > 0:
        try:
            with open(DEVICES_JSON, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except Exception:
            pass  # bozuksa Excel'e düş

    # Excel'den oku
    devices = []
    if os.path.exists(DEVICES_XLSX):
        wb = load_workbook(DEVICES_XLSX)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            name, ip = row[:2]
            if name and ip:
                devices.append({
                    "name": str(name),
                    "ip": str(ip),
                    "last_ping": None
                })

    save_devices(devices)
    return devices


def save_devices(devices):
    with open(DEVICES_JSON, "w", encoding="utf-8") as f:
        json.dump(devices, f, indent=2, ensure_ascii=False)


def update_last_ping(devices, ip):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for d in devices:
        if d["ip"] == ip:
            d["last_ping"] = now
    save_devices(devices)