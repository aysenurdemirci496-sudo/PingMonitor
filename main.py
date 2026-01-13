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

# ---------------- PLATFORM ----------------
IS_WINDOWS = platform.system().lower() == "windows"
FONT_NAME = "Segoe UI" if IS_WINDOWS else "Helvetica"

# ---------------- GLOBAL STATE ----------------
devices = []
current_ip = None
is_running = False
ping_process = None
ping_thread = None
ui_queue = queue.Queue()
selected_index = 0   # OK tuşları için

# ---------------- PING HELPERS ----------------
def extract_ping_ms(text):
    match = re.search(r"time[=<]?\s*([\d\.]+)\s*ms", text.lower())
    return float(match.group(1)) if match else None

def ip_exists(ip, exclude_device=None):
    for d in devices:
        if d["ip"] == ip:
            if exclude_device and d is exclude_device:
                continue
            return True
    return False


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


def ping_command(ip):
    return ["ping", "-t", ip] if IS_WINDOWS else ["ping", ip]

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
        ui_queue.put(line)

    try:
        ping_process.terminate()
    except Exception:
        pass

# ---------------- UI QUEUE ----------------
def process_ui_queue():
    while not ui_queue.empty():
        line = ui_queue.get()

        output_box.config(state=tk.NORMAL)
        output_box.insert(tk.END, line)
        output_box.see(tk.END)
        output_box.config(state=tk.DISABLED)

        ms = extract_ping_ms(line)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for d in devices:
            if d["ip"] == current_ip:
                d["latency"] = ms
                d["last_ping"] = now
                d["status"] = status_by_latency(ms)
                break

        save_devices(devices)
        refresh_device_list(keep_selection=True)

    root.after(100, process_ui_queue)

# ---------------- ACTIONS ----------------
def start_ping(event=None):
    global is_running, current_ip, ping_thread

    ip = ip_entry.get().strip()
    if not ip:
        return

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


def stop_ping(event=None):
    global is_running, ping_process

    is_running = False
    start_btn.config(text="Başlat")

    if ping_process:
        try:
            ping_process.terminate()
        except Exception:
            pass
        ping_process = None

    # 🔴 TREEVIEW GERİ ODAK
    device_tree.focus_set()

    # 🔴 KALDIĞI SATIRI KORU
    items = device_tree.get_children()
    if items:
        device_tree.selection_set(items[selected_index])
        device_tree.focus(items[selected_index])
        device_tree.see(items[selected_index])  # 👈 scroll


def toggle_ping():
    if is_running:
        stop_ping()
    else:
        start_ping()


def refresh_from_excel():
    global devices
    excel_devices = load_devices_from_excel()
    merged = []
    seen_ips = set()

    for ex in excel_devices:
        if ex["ip"] in seen_ips:
            messagebox.showwarning("Uyarı", "Bu IP adresi zaten mevcut!")
            continue

        seen_ips.add(ex["ip"])

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

# ---------------- DEVICE LIST ----------------
def refresh_device_list(keep_selection=False):
    global selected_index

    items = device_tree.get_children()
    if keep_selection and items:
        sel = device_tree.selection()
        if sel:
            selected_index = items.index(sel[0])

    device_tree.delete(*device_tree.get_children())

    for d in devices:
        latency_txt = "-" if d.get("latency") is None else f"{d['latency']:.1f}"
        device_tree.insert(
            "",
            tk.END,
            values=(d["name"], d["ip"], latency_txt, d.get("last_ping") or "-"),
            tags=(d.get("status", "UNKNOWN"),)
        )

    items = device_tree.get_children()
    if items:
        selected_index = min(selected_index, len(items) - 1)
        device_tree.selection_set(items[selected_index])
        device_tree.focus(items[selected_index])
        device_tree.see(items[selected_index])
        write_ip_from_selection()

def write_ip_from_selection():
    sel = device_tree.selection()
    if not sel:
        return
    ip_entry.delete(0, tk.END)
    ip_entry.insert(0, device_tree.item(sel[0])["values"][1])

def on_single_click(event):
    write_ip_from_selection()

def on_double_click(event):
    write_ip_from_selection()

    # 🔴 Treeview focus'u SABİTLE
    device_tree.focus_set()

    # 🔴 selection index’i koru
    global selected_index
    items = device_tree.get_children()
    sel = device_tree.selection()
    if sel:
        selected_index = items.index(sel[0])

    # 🔴 Ping'i gecikmeli başlat (UI zinciri bitsin)
    root.after(150, start_ping)

def on_arrow_key(event):
    global selected_index

    items = device_tree.get_children()
    if not items:
        return "break"

    if event.keysym == "Down" and selected_index < len(items) - 1:
        selected_index += 1
    elif event.keysym == "Up" and selected_index > 0:
        selected_index -= 1
    else:
        return "break"

    device_tree.selection_set(items[selected_index])
    device_tree.focus(items[selected_index])
    device_tree.see(items[selected_index])  # 🔴 SCROLL
    write_ip_from_selection()
    return "break"

# ---------------- CONTEXT MENU ----------------
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
        refresh_from_excel()
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

def show_context_menu(event):
    global selected_index

    if hasattr(event, "y"):
        row_id = device_tree.identify_row(event.y)
    else:
        sel = device_tree.selection()
        row_id = sel[0] if sel else None

    if not row_id:
        return

    items = device_tree.get_children()
    selected_index = items.index(row_id)

    device_tree.selection_set(row_id)
    device_tree.focus(row_id)
    write_ip_from_selection()

    x = event.x_root if hasattr(event, "x_root") else root.winfo_pointerx()
    y = event.y_root if hasattr(event, "y_root") else root.winfo_pointery()
    context_menu.tk_popup(x, y)

# ---------------- UI ----------------
root = tk.Tk()
root.title("Ping Monitor")
root.geometry("1100x650")
root.minsize(1100, 650)

# 🔑 PLATFORM FONT STYLE (SADECE EKLENEN KISIM)
style = ttk.Style(root)
style.configure("Treeview", font=(FONT_NAME, 14), rowheight=36)
style.configure("Treeview.Heading", font=(FONT_NAME, 15, "bold"))

top = tk.Frame(root)
top.pack(fill=tk.X, padx=10, pady=5)

tk.Label(top, text="IP:", font=(FONT_NAME, 11)).pack(side=tk.LEFT)
ip_entry = tk.Entry(top, width=25, font=(FONT_NAME, 11))
ip_entry.pack(side=tk.LEFT, padx=5)

start_btn = tk.Button(top, text="Başlat", width=10, command=toggle_ping)
start_btn.pack(side=tk.LEFT, padx=5)

tk.Button(top, text="Yenile", width=10, command=refresh_from_excel).pack(side=tk.LEFT)

main = tk.PanedWindow(root, orient=tk.HORIZONTAL)
main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

output_box = tk.Text(main, state=tk.DISABLED, font=(FONT_NAME, 11))
main.add(output_box)

right = tk.Frame(main)
main.add(right)

cols = ("Cihaz", "IP", "Ping (ms)", "Son Ping")
device_tree = ttk.Treeview(right, columns=cols, show="headings")

for c in cols:
    device_tree.heading(c, text=c)
    device_tree.column(c, width=160 if c != "Son Ping" else 220)

device_tree.pack(fill=tk.BOTH, expand=True)

# BINDINGS
device_tree.bind("<<TreeviewSelect>>", on_single_click)
device_tree.bind("<Double-1>", on_double_click)
device_tree.bind("<Up>", lambda e: "break")
device_tree.bind("<Down>", lambda e: "break")

device_tree.bind("<Button-3>", show_context_menu)
device_tree.bind("<Button-2>", show_context_menu)
device_tree.bind("<Control-Button-1>", show_context_menu)
root.bind("<Shift-F10>", show_context_menu)

root.bind("<Up>", on_arrow_key)
root.bind("<Down>", on_arrow_key)
root.bind("<Return>", start_ping)
root.bind("<Escape>", stop_ping)

# RENKLER
device_tree.tag_configure("UNKNOWN", foreground="#7f8c8d")
device_tree.tag_configure("FAST", foreground="#1e8449")
device_tree.tag_configure("NORMAL", foreground="#27ae60")
device_tree.tag_configure("SLOW", foreground="#b7950b")
device_tree.tag_configure("VERY_SLOW", foreground="#ca6f1e")
device_tree.tag_configure("DOWN", foreground="#c0392b")

# CONTEXT MENU
context_menu = tk.Menu(root, tearoff=0)
context_menu.add_command(label="▶ Ping Başlat", command=start_ping)
context_menu.add_command(label="⏹ Ping Durdur", command=stop_ping)
context_menu.add_separator()

# 🔴 YENİ EKLENEN
context_menu.add_command(label="📝 Cihaz Detayları", command=show_device_details)

context_menu.add_separator()
context_menu.add_command(label="📋 IP Kopyala", command=copy_selected_ip)
context_menu.add_separator()
context_menu.add_command(label="🔄 Excel'den Yenile", command=refresh_from_excel)

# ---------------- START ----------------
devices = load_devices()
refresh_device_list()
root.after(100, process_ui_queue)
root.mainloop()