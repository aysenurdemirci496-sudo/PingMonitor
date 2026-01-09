import json
import os
from openpyxl import load_workbook


def load_devices(config_path="config.json"):
    if os.path.exists("devices.json"):
        try:
            with open("devices.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, list):
                    return data
        except (json.JSONDecodeError, OSError):
            pass  # bozuk dosya → yeniden üret

    return _load_from_excel_and_create_json(config_path)


def merge_devices_from_excel(config_path="config.json"):
    old_devices = load_devices(config_path)
    old_map = {d["ip"]: d for d in old_devices}

    new_devices = _load_from_excel_raw(config_path)

    merged = []
    for d in new_devices:
        if d["ip"] in old_map:
            merged.append(old_map[d["ip"]])  # last_ping korunur
        else:
            merged.append(d)

    with open("devices.json", "w", encoding="utf-8") as f:
        json.dump(merged, f, indent=2, ensure_ascii=False)

    return merged


# ---------- helpers ----------

def _load_from_excel_and_create_json(config_path):
    devices = _load_from_excel_raw(config_path)
    with open("devices.json", "w", encoding="utf-8") as f:
        json.dump(devices, f, indent=2, ensure_ascii=False)
    return devices


def _load_from_excel_raw(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    wb = load_workbook(config["excel_file"])
    sheet = wb.active

    headers = {cell.value: i for i, cell in enumerate(sheet[1])}
    name_col = config["columns"]["name"]
    ip_col = config["columns"]["ip"]

    devices = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name = row[headers[name_col]]
        ip = row[headers[ip_col]]
        if name and ip:
            devices.append({
                "name": str(name),
                "ip": str(ip),
                "last_ping": None
            })

    return devices