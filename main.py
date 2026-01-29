import tkinter as tk
from tkinter import ttk
import subprocess
import threading
import queue
import platform
import re
from datetime import datetime
from tkinter import messagebox

from device_loader import load_devices, load_devices_from_excel, save_devices
from tkinter import filedialog
import json
import os



# ---------------- PLATFORM ----------------
IS_WINDOWS = platform.system().lower() == "windows"
FONT_NAME = "Segoe UI" if IS_WINDOWS else "Helvetica"

# ---------------- GLOBAL STATE ----------------
CONFIG_FILE = "config.json"
excel_path = None
excel_mapping = None
bulk_total = 0
bulk_done = 0
PAGE_SIZE = 100
current_page = 1
total_pages = 1
REQUIRED_FIELDS = [
    "ip",
    "device",
    "name",
    "model",
    "mac",
    "location",
    "unit",
    "description"
]
COLUMN_WIDTHS = {
    "Cihaz": 150,
    "IP": 130,
    "Ping (ms)": 100,
    "Son Ping": 170,
    "Device Name": 160,   # 👈 EKLE
    "Model": 120,
    "MAC": 160,
    "Location": 140,
    "Unit": 120,
    "Description": 200
}
current_sort_column = None
current_sort_reverse = False

def update_column_headers():
    for col in cols:
        text = col
        if col == current_sort_column:
            text += " ▼" if current_sort_reverse else " ▲"
        device_tree.heading(col, text=text)
devices = []
current_ip = None
is_running = False
ping_process = None
is_bulk_running = False
ping_thread = None
ui_queue = queue.Queue()
started_from_entry = False

def update_column_headers():
    for col, field in COLUMN_TO_FIELD.items():
        text = col

        # 🔍 Filtre aktif mi?
        if active_filters.get(field):
            text += " 🔍"

        # ⬆⬇ sıralama oku
        if col == current_sort_column:
            text += " ▼" if current_sort_reverse else " ▲"
        else:
            text += " ▼"

        device_tree.heading(col, text=text)

def update_tree_item_for_ip(ip):
    for item in device_tree.get_children():
        values = device_tree.item(item)["values"]
        if values and values[1] == ip:
            device = next((d for d in devices if d["ip"] == ip), None)
            if not device:
                return

            latency_txt = "-" if device.get("latency") is None else f"{device['latency']:.1f}"

            device_tree.item(
                item,
                values=(
                    device.get("device", ""),
                    device.get("ip", ""),
                    latency_txt,
                    device.get("last_ping") or "-",
                    device.get("name", ""),   # 👈 EKLE
                    device.get("model", ""),
                    device.get("mac", ""),
                    device.get("location", ""),
                    device.get("unit", ""),
                    device.get("description", "")
                ),
                tags=(device.get("status", "UNKNOWN"),)
            )
            return

# ---------------- PING HELPERS ----------------
def extract_ping_ms(text):
    text = text.lower()

    # time=14.2 ms | time<1ms | time=1ms
    match = re.search(r"time[=<]?\s*([\d\.]+)\s*ms", text)
    if match:
        return float(match.group(1))

    # bazı sistemlerde boşluk yok: time<1ms
    match = re.search(r"time[=<]?([\d\.]+)ms", text)
    if match:
        return float(match.group(1))

    return None

def single_ping(ip, timeout=2):
    try:
        if IS_WINDOWS:
            cmd = ["ping", "-n", "1", "-w", str(timeout * 1000), ip]
            startupinfo = subprocess.STARTUPINFO()
            startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW

            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                startupinfo=startupinfo,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
        else:
            cmd = ["ping", "-c", "1", ip]
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

        output, _ = proc.communicate(timeout=timeout + 1)
        return extract_ping_ms(output)

    except Exception:
        return None
    
from concurrent.futures import ThreadPoolExecutor

def bulk_ping_devices(devices_to_ping):
    def ping_one(device):
        ip = device["ip"]
        ms = single_ping(ip)
        ui_queue.put(("BULK", ip, ms))

    # aynı anda EN FAZLA 10 ping
    with ThreadPoolExecutor(max_workers=10) as executor:
        executor.map(ping_one, devices_to_ping)

def ip_exists(ip, exclude_device=None):
    for d in devices:
        if d["ip"] == ip:
            if exclude_device and d is exclude_device:
                continue
            return True
    return False

def bulk_ping_worker(devices_to_ping):
    global is_bulk_running

    def ping_one(device):
        ip = device["ip"]
        ms = single_ping(ip)
        ui_queue.put(("BULK", ip, ms))

    with ThreadPoolExecutor(max_workers=10) as executor:
        executor.map(ping_one, devices_to_ping)

    ui_queue.put(("BULK_DONE", None, None))

def device_matches_filters(device):
    # 1️⃣ Checkbox filtreleri
    for field, selected_values in active_filters.items():
        if selected_values:
            if device.get(field) not in selected_values:
                return False

    # 2️⃣ Global arama
    if search_text:
        haystack = " ".join(
            str(device.get(k, "")).lower()
            for k in [
                "device", "ip", "model",
                "mac", "location", "unit", "description"
            ]
        )
        if search_text.lower() not in haystack:
            return False

    return True

def status_by_latency(ms):
    if ms is None:
        return "DOWN"
    if ms < 50:
        return "FAST"
    elif ms < 100:
        return "NORMAL"
    elif ms < 200:
        return "SLOW"
    return "VERY_SLOW"


def sort_devices_by_column(col, reverse=False):
    global current_page
    current_page = 1
    global current_sort_column, current_sort_reverse

    if col == "IP":
        devices.sort(key=lambda d: ip_to_tuple(d.get("ip", "")), reverse=reverse)
    else:
        field = COLUMN_TO_FIELD.get(col)
        devices.sort(
            key=lambda d: (d.get(field) or "").lower(),
            reverse=reverse
        )

    current_sort_column = col
    current_sort_reverse = reverse

    refresh_device_list(keep_selection=True)
    update_column_headers()

def ping_command(ip):
    return ["ping", "-t", ip] if IS_WINDOWS else ["ping", ip]

def ip_to_tuple(ip):
    try:
        return tuple(int(x) for x in ip.split("."))
    except:
        return (0, 0, 0, 0)

def load_config():
    global excel_path
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            excel_path = data.get("excel_path")


def save_config():
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(
            {"excel_path": excel_path},
            f,
            ensure_ascii=False,
            indent=2
        )

# ---------------- FILTER STATE ----------------
FILTERABLE_FIELDS = {
    "device": "Cihaz",
    "ip": "IP",
    "name": "Device Name",   # 👈 EKLE
    "model": "Model",
    "mac": "MAC",
    "location": "Location",
    "unit": "Unit",
    "description": "Description"
}
COLUMN_TO_FIELD = {
    "Cihaz": "device",
    "IP": "ip",
    "Device Name": "name",   # 👈 EKLE
    "Model": "model",
    "MAC": "mac",
    "Location": "location",
    "Unit": "unit",
    "Description": "description"
}
# 🔹 Treeview
cols = (
    "Cihaz",
    "IP",
    "Ping (ms)",
    "Son Ping",
    "Device Name",   # 👈 YENİ
    "Model",
    "MAC",
    "Location",
    "Unit",
    "Description"
)

active_filters = {key: set() for key in FILTERABLE_FIELDS}
search_text = ""

# ---------------- PING LOOP ----------------
def ping_loop(ip):
    global ping_process

    flags = subprocess.CREATE_NO_WINDOW if IS_WINDOWS else 0

    ping_process = subprocess.Popen(
        ping_command(ip),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        creationflags=flags
    )

    for line in ping_process.stdout:
        if not is_running:
            break
        ui_queue.put(("SINGLE", ip, line))

    try:
        ping_process.terminate()
    except Exception:
        pass

# ---------------- UI QUEUE ----------------
def process_ui_queue():
    while not ui_queue.empty():
        item = ui_queue.get()

        item_type, ip, payload = item
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 🔵 TEKLİ PING
        if item_type == "SINGLE":
            line = payload

            output_box.config(state=tk.NORMAL)
            output_box.insert(tk.END, line)
            output_box.see(tk.END)
            output_box.config(state=tk.DISABLED)

            ms = extract_ping_ms(line)

            for d in devices:
                if d["ip"] == ip:
                    d["latency"] = ms
                    d["last_ping"] = now
                    d["status"] = status_by_latency(ms)
                    update_tree_item_for_ip(ip)
                    break

            save_devices(devices)

        # 🟢 TOPLU PING
        elif item_type == "BULK":
            global bulk_done

            bulk_done += 1   # 🔥 BAŞARILI / BAŞARISIZ FARK ETMEZ

            bulk_status_label.config(
                text=f"Toplu Ping: {bulk_done} / {bulk_total}"
            )

            ms = payload  # None olabilir, sorun değil

            for d in devices:
                if d["ip"] == ip:
                    d["latency"] = ms
                    d["last_ping"] = now
                    d["status"] = status_by_latency(ms)
                    update_tree_item_for_ip(ip)
                    break

          

        # 🔴 TOPLU PING BİTTİ
        elif item_type == "BULK_DONE":
            global is_bulk_running

            is_bulk_running = False
            save_devices(devices)
            start_btn.config(state=tk.NORMAL)
            refresh_btn.config(state=tk.NORMAL)
            add_btn.config(state=tk.NORMAL)
            bulk_status_label.config(
            text=f"Toplu Ping tamamlandı ({bulk_total} / {bulk_total})"
            )
              # ⏱ 3 saniye sonra temizle
            root.after(5000, lambda: bulk_status_label.config(text=""))
  

    root.after(30, process_ui_queue)

# ---------------- ACTIONS ----------------
def start_ping(event=None):
    global is_running, current_ip, ping_thread, started_from_entry
    bulk_status_label.config(text="")
    ip = ip_entry.get().strip()
    if not ip:
        return

    # 🔴 BU SATIR KRİTİK
    started_from_entry = True

    stop_ping()

    is_running = True
    current_ip = ip
    start_btn.config(text="Durdur")

    output_box.config(state=tk.NORMAL)
    output_box.delete("1.0", tk.END)
    output_box.config(state=tk.DISABLED)

    while not ui_queue.empty():
        ui_queue.get_nowait()

    ping_thread = threading.Thread(target=ping_loop, args=(ip,), daemon=True)
    ping_thread.start()

def start_ping_from_menu():
    global started_from_entry
    started_from_entry = False
    start_ping()

def stop_ping(event=None):
    global is_running, ping_process

    is_running = False
    start_btn.config(text="Başlat")
        # Ping durdu, artık entry'yi otomatik doldurabiliriz
    global started_from_entry

    if ping_process:
        try:
            ping_process.terminate()
        except Exception:
            pass
        ping_process = None


def toggle_ping():
    if is_running:
        stop_ping()
    else:
        start_ping()


def refresh_from_excel():
    if not excel_path:
        messagebox.showwarning("Excel", "Önce Excel dosyası seçin")
        return
    global devices
    bulk_status_label.config(text="")

    excel_devices = load_devices_from_excel(
        excel_path,
        mapping=None,          # ⬅️ burası mapping UI’yi tetikler
        parent=root            # ⬅️ Mac için şart
    )
    if not excel_mapping:
        messagebox.showwarning(
            "Excel Eşleştirme Yok",
            "Önce Excel Seç butonundan dosya ve kolon eşleştirmesi yapmalısın."
        )
        return

    excel_devices = load_devices_from_excel(excel_path, excel_mapping)
    merged = []

    for ex in excel_devices:
        old = next((d for d in devices if d["ip"] == ex["ip"]), None)
        if old:
            merged.append(old)
        else:
            merged.append({
                **ex,
                "latency": None,
                "last_ping": None,
                "status": "UNKNOWN"
            })

    devices = merged
    save_devices(devices)
    refresh_device_list(keep_selection=True)

def open_mapping_window(excel_headers, on_done):
    win = tk.Toplevel(root)
    win.title("Excel Kolon Eşleştirme")
    win.grab_set()

    tk.Label(
        win,
        text="Excel kolonlarını uygulama alanlarıyla eşleştir",
        font=(FONT_NAME, 11, "bold")
    ).pack(pady=10)

    mapping_vars = {}

    frame = tk.Frame(win)
    frame.pack(padx=10, pady=10)

    for header in excel_headers:
        row = tk.Frame(frame)
        row.pack(fill=tk.X, pady=3)

        tk.Label(row, text=header, width=25, anchor="w").pack(side=tk.LEFT)

        var = tk.StringVar(value="")
        combo = ttk.Combobox(
            row,
            textvariable=var,
            values=[""] + REQUIRED_FIELDS,
            state="readonly",
            width=20
        )
        combo.pack(side=tk.LEFT)

        mapping_vars[header] = var

    def apply():
        mapping = {}
        for header, var in mapping_vars.items():
            if var.get():
                mapping[var.get()] = header

        if "ip" not in mapping:
            messagebox.showerror("Hata", "IP alanı mutlaka eşleştirilmeli")
            return

        win.destroy()
        on_done(mapping)

    tk.Button(win, text="Tamam", command=apply).pack(pady=10)

def select_excel_file():
    global excel_path

    path = filedialog.askopenfilename(
        title="Excel dosyasını seç",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )

    if not path:
        return

    excel_path = path
    save_config()

    import pandas as pd
    df = pd.read_excel(excel_path)
    headers = list(df.columns)

    def on_mapping_done(mapping):
        global devices, excel_mapping
        excel_mapping = mapping

        devices = load_devices_from_excel(excel_path, excel_mapping)
        save_devices(devices)
        refresh_device_list()

    open_mapping_window(headers, on_mapping_done)

# ---------------- DEVICE LIST ----------------
def extend_selection(direction):
    items = device_tree.get_children()
    if not items:
        return "break"

    sel = device_tree.selection()

    # hiç seçim yoksa → ilk satırı seç
    if not sel:
        device_tree.selection_set(items[0])
        device_tree.focus(items[0])
        return "break"

    # son seçili satır
    last = sel[-1]
    index = items.index(last)

    new_index = index + direction
    if new_index < 0 or new_index >= len(items):
        return "break"

    new_item = items[new_index]

    # yeni satırı DAHİL ET
    device_tree.selection_add(new_item)
    device_tree.focus(new_item)
    device_tree.see(new_item)

    return "break"

def select_all_rows(event=None):
    items = device_tree.get_children()
    if not items:
        return "break"

    device_tree.selection_set(items)
    device_tree.focus(items[0])
    device_tree.see(items[0])
    return "break"

    
def get_selected_devices():
    selected = []
    for item in device_tree.selection():
        values = device_tree.item(item)["values"]
        if not values:
            continue

        ip = values[1]
        dev = next((d for d in devices if d["ip"] == ip), None)
        if dev:
            selected.append(dev)

    return selected

def on_tree_arrow(direction):
    move_selection(direction)
    return "break"   # 🔴 EN KRİTİK SATIR
    
def move_selection(direction):
    items = device_tree.get_children()
    if not items:
        return

    sel = device_tree.selection()
    if sel:
        index = items.index(sel[0])
    else:
        index = 0

    new_index = index + direction

    if new_index < 0 or new_index >= len(items):
        return

    device_tree.selection_set(items[new_index])
    device_tree.focus(items[new_index])
    device_tree.see(items[new_index])
    def write_ip_from_selection():
        global started_from_entry

        # 🔴 BİRDEN FAZLA SEÇİM VARSA HİÇBİR ŞEY YAPMA
        if len(device_tree.selection()) > 1:
            return

def get_paged_devices(filtered_devices):
    global total_pages

    total_pages = max(1, (len(filtered_devices) + PAGE_SIZE - 1) // PAGE_SIZE)

    start = (current_page - 1) * PAGE_SIZE
    end = start + PAGE_SIZE

    return filtered_devices[start:end]

def refresh_device_list(keep_selection=False):
    # mevcut seçimi hatırla
    selected_ip = None
    if keep_selection:
        sel = device_tree.selection()
        if sel:
            selected_ip = device_tree.item(sel[0])["values"][1]

    device_tree.delete(*device_tree.get_children())
    filtered = [d for d in devices if device_matches_filters(d)]
    paged_devices = get_paged_devices(filtered)

    for d in paged_devices:
        if not device_matches_filters(d):
            continue

        latency_txt = "-" if d.get("latency") is None else f"{d['latency']:.1f}"
        device_tree.insert(
            "",
            tk.END,
            values=(
                d.get("device", ""),
                d.get("ip", ""),
                latency_txt,
                d.get("last_ping") or "-",
                d.get("name", ""),     # 👈 DEVICE NAME
                d.get("model", ""),
                d.get("mac", ""),
                d.get("location", ""),
                d.get("unit", ""),
                d.get("description", "")
            ),
            tags=(d.get("status", "UNKNOWN"),)
        )

    # aynı IP varsa onu tekrar seç
    if selected_ip:
        for item in device_tree.get_children():
            if device_tree.item(item)["values"][1] == selected_ip:
                device_tree.selection_set(item)
                device_tree.focus(item)
                device_tree.see(item)
                write_ip_from_selection()
                return

    # yoksa ilk satırı seç
    items = device_tree.get_children()
    if items:
        device_tree.selection_set(items[0])
        device_tree.focus(items[0])
        device_tree.see(items[0])
        write_ip_from_selection()

    update_page_label()
    device_tree.update_idletasks()
    


def write_ip_from_selection():
    global started_from_entry

    if started_from_entry:
        return

    sel = device_tree.selection()
    if not sel:
        return

    ip_entry.delete(0, tk.END)
    ip_entry.insert(0, device_tree.item(sel[0])["values"][1])

def on_tree_select(event=None):
    global started_from_entry
    started_from_entry = False   # 👈 kilidi burada açıyoruz
    write_ip_from_selection()

def on_double_click(event):
    global started_from_entry
    row_id = device_tree.identify_row(event.y)
    if not row_id:
        return

    device_tree.selection_set(row_id)
    device_tree.focus(row_id)
    write_ip_from_selection()
    started_from_entry = False

    root.after(50, start_ping)

def move_focus_horizontal(direction):
    current = root.focus_get()

    if current not in top_controls:
        top_controls[0].focus_set()
        return

    index = top_controls.index(current)
    new_index = index + direction

    if 0 <= new_index < len(top_controls):
        top_controls[new_index].focus_set()

    refresh_device_list(keep_selection=True)

def show_sort_menu(event, col):
    menu = tk.Menu(root, tearoff=0)

    menu.add_command(
        label="A'dan Z'ye Sırala",
        command=lambda: sort_devices_by_column(col, reverse=False)
    )

    menu.add_command(
        label="Z'den A'ya Sırala",
        command=lambda: sort_devices_by_column(col, reverse=True)
    )

    menu.tk_popup(event.x_root, event.y_root)



# ---------------- CONTEXT MENU ----------------
def start_bulk_ping():
    global is_bulk_running, bulk_total, bulk_done

    if is_running:
        messagebox.showwarning(
            "Uyarı",
            "Tekli ping çalışırken toplu ping başlatılamaz."
        )
        return

    if is_bulk_running:
        return

    # ✅ SADECE SEÇİLENLER
    devices_to_ping = get_selected_devices()

    if not devices_to_ping:
        messagebox.showinfo(
            "Bilgi",
            "Lütfen en az bir cihaz seçin"
        )
        return

    is_bulk_running = True

    bulk_total = len(devices_to_ping)
    bulk_done = 0

    bulk_status_label.config(
        text=f"Toplu Ping: 0 / {bulk_total}"
    )

    start_btn.config(state=tk.DISABLED)
    refresh_btn.config(state=tk.DISABLED)
    add_btn.config(state=tk.DISABLED)

    threading.Thread(
        target=bulk_ping_worker,
        args=(devices_to_ping,),
        daemon=True
    ).start()

def start_bulk_ping_all_filtered():
    global is_bulk_running, bulk_total, bulk_done

    if is_running:
        messagebox.showwarning(
            "Uyarı",
            "Tekli ping çalışırken toplu ping başlatılamaz."
        )
        return

    if is_bulk_running:
        return

    # 🔴 SAYFAYA DEĞİL → TÜM FİLTRELİ CİHAZLAR
    filtered_devices = [
        d for d in devices
        if device_matches_filters(d)
    ]

    if not filtered_devices:
        messagebox.showinfo(
            "Bilgi",
            "Filtrelere uyan cihaz bulunamadı"
        )
        return

    # 🔥 KRİTİK SATIRLAR (SENDEN EKSİK OLANLAR)
    bulk_total = len(filtered_devices)
    bulk_done = 0

    bulk_status_label.config(
        text=f"Toplu Ping: 0 / {bulk_total}"
    )

    is_bulk_running = True

    # 🔒 UI kilidi
    start_btn.config(state=tk.DISABLED)
    refresh_btn.config(state=tk.DISABLED)
    add_btn.config(state=tk.DISABLED)

    threading.Thread(
        target=bulk_ping_worker,
        args=(filtered_devices,),
        daemon=True
    ).start()


def open_add_device_window():
    ip = ip_entry.get().strip()

    if not ip:
        messagebox.showwarning("Uyarı", "Önce IP adresi giriniz")
        return

    if ip_exists(ip):
        messagebox.showwarning("IP Çakışması", f"Bu IP zaten kayıtlı:\n{ip}")
        return

    win = tk.Toplevel(root)
    win.title("Yeni Cihaz Ekle")
    win.resizable(False, False)
    win.grab_set()

    fields = [
        ("Device Name", ""),
        ("IP Address", ip),
        ("Device", ""),
        ("Model", ""),
        ("MAC", ""),
        ("Location", ""),
        ("Unit", ""),
        ("Description", ""),
    ]

    entries = {}

    for i, (label, value) in enumerate(fields):
        tk.Label(win, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=4)

        ent = tk.Entry(win, width=40)
        ent.grid(row=i, column=1, padx=10, pady=4)
        ent.insert(0, value)

        entries[label] = ent
    def save_new_device():
        new_ip = entries["IP Address"].get().strip()

        if not new_ip:
            messagebox.showwarning("Hata", "IP Address boş olamaz")
            return

        if ip_exists(new_ip):
            messagebox.showwarning(
                "IP Çakışması",
                f"Bu IP zaten başka bir cihaza ait:\n{new_ip}"
            )
            return

        new_device = {
            "name": entries["Device Name"].get(),
            "ip": new_ip,
            "device": entries["Device"].get(),
            "model": entries["Model"].get(),
            "mac": entries["MAC"].get(),
            "location": entries["Location"].get(),
            "unit": entries["Unit"].get(),
            "description": entries["Description"].get(),
            "latency": None,
            "last_ping": None,
            "status": "UNKNOWN"
        }

        devices.append(new_device)

        from device_loader import add_device_to_excel
        add_device_to_excel(new_device)

        save_devices(devices)
        refresh_device_list(keep_selection=True)

        win.destroy()

    btns = tk.Frame(win)
    btns.grid(row=len(fields), column=0, columnspan=2, pady=10)

    tk.Button(btns, text="Kaydet", width=12, command=save_new_device).pack(side=tk.LEFT, padx=5)
    tk.Button(btns, text="İptal", width=12, command=win.destroy).pack(side=tk.LEFT, padx=5)
        
    
        
def show_device_details():
    sel = device_tree.selection()
    if not sel:
        return

    item = device_tree.item(sel[0])
    ip = item["values"][1]

    device = next((d for d in devices if d["ip"] == ip), None)
    if not device:
        return

    win = tk.Toplevel(root)
    win.title("Cihaz Detayları")
    win.resizable(False, False)
    win.grab_set()  # modal pencere

    fields = [
        ("Device Name", device.get("name"), True),
        ("IP Address", device.get("ip"), True),
        ("Device", device.get("device"), True),
        ("Model", device.get("model"), True),
        ("MAC", device.get("mac"), True),
        ("Location", device.get("location"), True),
        ("Unit", device.get("unit"), True),
        ("Description", device.get("description"), True),
    ]

    entries = {}

    for i, (label, value, editable) in enumerate(fields):
        tk.Label(win, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=4)

        ent = tk.Entry(win, width=40)
        ent.grid(row=i, column=1, padx=10, pady=4)
        ent.insert(0, value if value else "")

        if not editable:
            ent.config(state="disabled")

        entries[label] = ent
    def save_changes():
        new_ip = entries["IP Address"].get().strip()

        if not new_ip:
            tk.messagebox.showwarning("Hata", "IP Address boş olamaz")
            return

        if ip_exists(new_ip, exclude_device=device):
            tk.messagebox.showwarning(
                "IP Çakışması",
                f"Bu IP zaten başka bir cihaza ait:\n{new_ip}"
            )
            return

        old_ip = device["ip"]

        device["name"] = entries["Device Name"].get()
        device["ip"] = new_ip
        device["device"] = entries["Device"].get()
        device["model"] = entries["Model"].get()
        device["mac"] = entries["MAC"].get()
        device["location"] = entries["Location"].get()
        device["unit"] = entries["Unit"].get()
        device["description"] = entries["Description"].get()

        from device_loader import update_device_in_excel
        update_device_in_excel(old_ip, device)

        save_devices(devices)
        refresh_device_list(keep_selection=True)
        win.destroy()

    btns = tk.Frame(win)
    btns.grid(row=len(fields), column=0, columnspan=2, pady=10)

    tk.Button(btns, text="Kaydet", width=12, command=save_changes).pack(side=tk.LEFT, padx=5)
    tk.Button(btns, text="İptal", width=12, command=win.destroy).pack(side=tk.LEFT, padx=5)
def copy_selected_ip():
    sel = device_tree.selection()
    if not sel:
        return
    ip = device_tree.item(sel[0])["values"][1]
    root.clipboard_clear()
    root.clipboard_append(ip)

def open_filter_window(field):
    win = tk.Toplevel(root)
    win.title(f"{FILTERABLE_FIELDS[field]} Filtre")
    win.geometry("320x450")
    win.resizable(False, False)
    win.grab_set()

    # ================== ARAMA ==================
    search_var = tk.StringVar()
    search_entry = tk.Entry(win, textvariable=search_var)
    search_entry.pack(fill=tk.X, padx=10, pady=(10, 5))

    # ================== ÜST BUTONLAR ==================
    top_btns = tk.Frame(win)
    top_btns.pack(fill=tk.X, padx=10, pady=5)

    # bunlar aşağıda tanımlanacak ama butonlar burada duracak
    def select_all(): pass
    def clear_all(): pass

    tk.Button(top_btns, text="☑️ Tümünü Seç", command=lambda: select_all()).pack(side=tk.LEFT)
    tk.Button(top_btns, text="❌ Temizle", command=lambda: clear_all()).pack(side=tk.LEFT, padx=5)

    # ================== SCROLLABLE ALAN ==================
    list_container = tk.Frame(win)
    list_container.pack(fill=tk.BOTH, expand=True, padx=10)

    canvas = tk.Canvas(
    list_container,
    borderwidth=0,
    highlightthickness=0
    )
    scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)

    scroll_frame = tk.Frame(canvas)

    scroll_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.configure(yscrollcommand=scrollbar.set)
    scrollbar.configure(command=canvas.yview)

    window_id = canvas.create_window(
    (0, 0),
    window=scroll_frame,
    anchor="nw"
    )
    def _resize_canvas(event):
        canvas.itemconfig(window_id, width=event.width)

        canvas.bind("<Configure>", _resize_canvas)
        canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    # 🟢 Mouse scroll desteği (Windows + Mac)
    def _on_mousewheel(event):
        if not canvas.winfo_exists():
            return

        if event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

    win.bind("<MouseWheel>", _on_mousewheel)
    win.bind("<Button-4>", _on_mousewheel)
    win.bind("<Button-5>", _on_mousewheel)   

    def _on_mousewheel(event):
        if not canvas.winfo_exists():
            return

        if event.delta:
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        elif event.num == 4:
            canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            canvas.yview_scroll(1, "units")

    # ================== ALT OK BUTONU ==================
    bottom = tk.Frame(win)
    bottom.pack(fill=tk.X, pady=10)

    def apply_filters():
        global current_page
        current_page = 1
        active_filters[field].clear()

        for val, var in vars_map.items():
            if var.get():
                active_filters[field].add(val)

        refresh_device_list()
        update_column_headers()
        win.destroy()

    tk.Button(bottom, text="OK", width=10, command=apply_filters).pack(side=tk.RIGHT, padx=10)

    # ================== VERİLER ==================
    values = sorted(
        set(str(d.get(field)) for d in devices if d.get(field)),
        key=lambda x: tuple(int(p) for p in x.split(".")) if field == "ip" else x.lower()
    )

    vars_map = {}
    checkbuttons = {}

    def render_list():
        for chk in checkbuttons.values():
            chk.destroy()
        checkbuttons.clear()
        vars_map.clear()

        keyword = search_var.get().lower()

        for val in values:
            if keyword and keyword not in val.lower():
                continue

            var = tk.BooleanVar(value=val in active_filters[field])
            chk = tk.Checkbutton(scroll_frame, text=val, variable=var)
            chk.pack(anchor="w")

            vars_map[val] = var
            checkbuttons[val] = chk
        canvas.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def select_all():
        for var in vars_map.values():
            var.set(True)

    def clear_all():
        for var in vars_map.values():
            var.set(False)

    search_var.trace_add("write", lambda *args: render_list())
    render_list()
    
def clear_all_filters():
    global current_page

    # 1️⃣ tüm filtre setlerini boşalt
    for field in active_filters:
        active_filters[field].clear()

    # 2️⃣ kolon başlıklarını eski haline döndür
    for col in COLUMN_TO_FIELD:
        device_tree.heading(col, text=f"{col} ▼")

    # 3️⃣ sayfayı başa al
    current_page = 1

    # 4️⃣ listeyi yenile
    refresh_device_list()
    update_column_headers()

def clear_single_filter(field):
    global current_page

    active_filters[field].clear()
    current_page = 1
    refresh_device_list()
    update_column_headers()

def show_column_menu(event, col):
    menu = tk.Menu(root, tearoff=0)

    menu.add_command(
        label="A'dan Z'ye Sırala",
        command=lambda: sort_devices_by_column(col, reverse=False)
    )
    menu.add_command(
        label="Z'den A'ya Sırala",
        command=lambda: sort_devices_by_column(col, reverse=True)
    )

    if col in COLUMN_TO_FIELD:
        menu.add_separator()
        menu.add_command(
            label="Filtrele",
            command=lambda: open_filter_window(COLUMN_TO_FIELD[col])
        )
        menu.add_command(
            label="Filtreyi Temizle",
            command=lambda: clear_single_filter(COLUMN_TO_FIELD[col])
        )
   

    menu.tk_popup(event.x_root, event.y_root)
    menu.grab_release()

def show_context_menu(event):
    # Sağ tık yapılan satırı bul
    row_id = device_tree.identify_row(event.y)

    if not row_id:
        return

    # Eğer o satır zaten seçiliyse -> HİÇBİR ŞEY YAPMA
    if row_id not in device_tree.selection():
        # değilse sadece o satırı seç (tekli senaryo)
        device_tree.selection_set(row_id)
        device_tree.focus(row_id)
        write_ip_from_selection()

    context_menu.tk_popup(event.x_root, event.y_root)

# ---------------- UI ----------------
root = tk.Tk()
root.title("Ping Monitor")
root.geometry("1100x650")
root.minsize(1100, 650)

# 🔑 PLATFORM FONT STYLE (SADECE EKLENEN KISIM)
style = ttk.Style(root)
style.configure(
    "Treeview",
    font=(FONT_NAME, 10, "bold"),
    rowheight=26
)

style.configure(
    "Treeview.Heading",
    font=(FONT_NAME, 11, "bold")
)

top = tk.Frame(root)
top.pack(fill=tk.X, padx=10, pady=5)
left_controls = tk.Frame(top)
left_controls.pack(side=tk.LEFT)

right_controls = tk.Frame(top)
right_controls.pack(side=tk.RIGHT)

tk.Label(left_controls, text="IP:", font=(FONT_NAME, 11)).pack(side=tk.LEFT)

ip_entry = tk.Entry(left_controls, width=25, font=(FONT_NAME, 11))
ip_entry.pack(side=tk.LEFT, padx=5)

start_btn = tk.Button(left_controls, text="Başlat", width=10, command=toggle_ping)
start_btn.pack(side=tk.LEFT, padx=5)

refresh_btn = tk.Button(left_controls, text="Yenile", width=10, command=refresh_from_excel)
refresh_btn.pack(side=tk.LEFT)

add_btn = tk.Button(
    left_controls,
    text="➕ Ekle",
    width=18,
    command=open_add_device_window
)
add_btn.pack(side=tk.LEFT, padx=5)

excel_btn = tk.Button(
    left_controls,
    text="📂 Excel Seç",
    width=14,
    command=select_excel_file
)
excel_btn.pack(side=tk.LEFT, padx=5)

tk.Label(right_controls, text="Ara:", font=(FONT_NAME, 11)).pack(side=tk.LEFT, padx=(0, 5))

search_entry = tk.Entry(right_controls, width=25, font=(FONT_NAME, 11))
search_entry.pack(side=tk.LEFT)

def on_search_change(event=None):
    global search_text
    search_text = search_entry.get().strip()
    refresh_device_list()

search_entry.bind("<KeyRelease>", on_search_change)

add_btn.config(width=18)

top_controls = [
    ip_entry,
    start_btn,
    refresh_btn,
    add_btn
]

main = tk.PanedWindow(root, orient=tk.HORIZONTAL)
main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

output_box = tk.Text(main, state=tk.DISABLED, font=(FONT_NAME, 11))
main.add(output_box)

right = tk.Frame(main)
main.add(right)


# 🔹 Treeview + Scrollbar için container
tree_container = tk.Frame(right)
tree_container.pack(fill=tk.BOTH, expand=True)


# 🔹 Dikey scrollbar
tree_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

# 🔹 Yatay scrollbar (DİKKAT: burada device_tree YOK)
tree_scroll_x = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)

# 🔹 Treeview

device_tree = ttk.Treeview(
    tree_container,
    columns=cols,
    show="headings",
    selectmode="extended",
    yscrollcommand=tree_scroll.set,
    xscrollcommand=tree_scroll_x.set,
    takefocus=False
)

def on_heading_click(event):
    region = device_tree.identify_region(event.x, event.y)
    if region != "heading":
        return

    col_id = device_tree.identify_column(event.x)
    col_index = int(col_id.replace("#", "")) - 1
    col_name = cols[col_index]

    show_column_menu(event, col_name)
    return "break"  

# 🔑 Treeview başlıklarını TANIMLA
for c in cols:
    device_tree.heading(c, text=c)
    device_tree.column(c, width=COLUMN_WIDTHS[c], anchor="w")

for c in cols:
    if c in COLUMN_TO_FIELD:
        device_tree.heading(c, text=f"{c} ▼")
    else:
        device_tree.heading(c, text=c)

# 🔹 Scrollbar ↔ Treeview bağlantısı
tree_scroll.config(command=device_tree.yview)
tree_scroll_x.config(command=device_tree.xview)

# 🔹 Column genişlikleri
for c in cols:
    device_tree.column(
        c,
        width=COLUMN_WIDTHS[c],
        minwidth=COLUMN_WIDTHS[c],
        stretch=False
    )

device_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
device_tree.update_idletasks()

bulk_status_label = tk.Label(right, text="", font=(FONT_NAME, 10))
bulk_status_label.pack(fill=tk.X, padx=10, pady=(0, 5))
# ---------------- PAGINATION UI ----------------
pagination = tk.Frame(right)
pagination.pack(fill=tk.X, pady=5)

page_label = tk.Label(pagination, text="Sayfa 1 / 1")
page_label.pack(side=tk.LEFT, padx=10)

def update_page_label():
    page_label.config(text=f"Sayfa {current_page} / {total_pages}")

def prev_page():
    global current_page
    if current_page > 1:
        current_page -= 1
        refresh_device_list()
        update_page_label()

def next_page():
    global current_page
    if current_page < total_pages:
        current_page += 1
        refresh_device_list()
        update_page_label()

tk.Button(pagination, text="◀ Önceki", command=prev_page).pack(side=tk.LEFT)
tk.Button(pagination, text="Sonraki ▶", command=next_page).pack(side=tk.LEFT)

# BINDINGS
device_tree.bind("<Button-1>", on_heading_click)
device_tree.bind("<<TreeviewSelect>>", on_tree_select)
device_tree.bind("<Double-1>", on_double_click)
device_tree.bind("<Up>", lambda e: on_tree_arrow(-1))
device_tree.bind("<Down>", lambda e: on_tree_arrow(1))
device_tree.bind("<Shift-Up>", lambda e: extend_selection(-1))
device_tree.bind("<Shift-Down>", lambda e: extend_selection(1))
device_tree.bind("<Control-Up>", lambda e: extend_selection(-1))
device_tree.bind("<Control-Down>", lambda e: extend_selection(1))
device_tree.bind("<Control-a>", select_all_rows)
device_tree.bind("<Command-a>", select_all_rows)  # macOS için

def on_mousewheel(event):
    device_tree.yview_scroll(int(-1*(event.delta/120)), "units")

device_tree.bind("<MouseWheel>", on_mousewheel)        # Windows
device_tree.bind("<Button-4>", lambda e: device_tree.yview_scroll(-1, "units"))  # Mac
device_tree.bind("<Button-5>", lambda e: device_tree.yview_scroll(1, "units"))   # Mac


device_tree.bind("<Button-3>", show_context_menu)
device_tree.bind("<Button-2>", show_context_menu)
device_tree.bind("<Control-Button-1>", show_context_menu)
root.bind("<Shift-F10>", show_context_menu)


root.bind("<Return>", start_ping)
root.bind("<Escape>", stop_ping)
root.bind("<Left>", lambda e: move_focus_horizontal(-1))
root.bind("<Right>", lambda e: move_focus_horizontal(1))

# RENKLER
device_tree.tag_configure("UNKNOWN", foreground="#7f8c8d")
device_tree.tag_configure("FAST", foreground="#1e8449")
device_tree.tag_configure("NORMAL", foreground="#27ae60")
device_tree.tag_configure("SLOW", foreground="#b7950b")
device_tree.tag_configure("VERY_SLOW", foreground="#ca6f1e")
device_tree.tag_configure("DOWN", foreground="#c0392b")

# CONTEXT MENU
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="▶ Ping Başlat", command=start_ping_from_menu)
context_menu.add_command(label="⏹ Ping Durdur", command=stop_ping)
context_menu.add_separator()
context_menu.add_command(
    label="📡 Seçilenlere Toplu Ping",
    command=start_bulk_ping
)
context_menu.add_command(
    label="🌐 Filtrelenmiş TÜM Cihazlara Ping",
    command=start_bulk_ping_all_filtered
)
context_menu.add_separator()

# 🔴 YENİ EKLENEN
context_menu.add_command(label="📝 Cihaz Detayları", command=show_device_details)

context_menu.add_separator()
context_menu.add_command(label="📋 IP Kopyala", command=copy_selected_ip)
context_menu.add_separator()
context_menu.add_command(label="🔄 Excel'den Yenile", command=refresh_from_excel)
context_menu.add_separator()
context_menu.add_command(
    label="🧹 Tüm Filtreleri Temizle",
    command=clear_all_filters
)

# ---------------- START ----------------
load_config()

devices = load_devices()

if excel_path and excel_mapping:
    refresh_from_excel()
else:
    refresh_device_list()

root.after(100, process_ui_queue)
root.mainloop()