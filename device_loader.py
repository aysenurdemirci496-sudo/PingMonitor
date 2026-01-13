import json
import os
from openpyxl import load_workbook

DEVICES_JSON = "devices.json"
DEVICES_XLSX = "devices.xlsx"


def load_devices():
    if os.path.exists(DEVICES_JSON):
        try:
            with open(DEVICES_JSON, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []


def load_devices_from_excel(xlsx_file=DEVICES_XLSX):
    devices = []

    if not os.path.exists(xlsx_file):
        return devices

    wb = load_workbook(xlsx_file)
    ws = wb.active

    for row in ws.iter_rows(min_row=2, values_only=True):
        (
            name,
            ip,
            device,
            model,
            mac,
            location,
            unit,
            description
        ) = row[:8]

        # 🔴 SADECE IP ZORUNLU
        if not ip:
            continue

        devices.append({
            # 🔹 Name yoksa IP’yi name olarak kullan
            "name": str(name) if name else str(ip),
            "ip": str(ip),
            "device": device or "",
            "model": model or "",
            "mac": mac or "",
            "location": location or "",
            "unit": unit or "",
            "description": description or ""
        })

    return devices


def save_devices(devices):
    with open(DEVICES_JSON, "w", encoding="utf-8") as f:
        json.dump(devices, f, indent=2, ensure_ascii=False)


def update_device_in_excel(old_ip, updated_device, xlsx_file=DEVICES_XLSX):
    if not os.path.exists(xlsx_file):
        return

    wb = load_workbook(xlsx_file)
    ws = wb.active

    for row in ws.iter_rows(min_row=2):
        if str(row[1].value) == str(old_ip):
            row[0].value = updated_device["name"]
            row[1].value = updated_device["ip"]
            row[2].value = updated_device.get("device", "")
            row[3].value = updated_device.get("model", "")
            row[4].value = updated_device.get("mac", "")
            row[5].value = updated_device.get("location", "")
            row[6].value = updated_device.get("unit", "")
            row[7].value = updated_device.get("description", "")
            break

    wb.save(xlsx_file)