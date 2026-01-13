import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import queue
import platform
import re
from datetime import datetime

from device_loader import load_devices, load_devices_from_excel, save_devices

# ---------------- GLOBAL STATE ----------------
devices = []
current_ip = None
is_running = False
ping_process = None
ping_thread = None
ui_queue = queue.Queue()
selected_index = 0   # 🔴 OK TUŞLARI İÇİN KRİTİK

# ---------------- PING HELPERS ----------------
def extract_ping_ms(text):
    match = re.search(r"time[=<]?\s*([\d\.]+)\s*ms", text.lower())
    return float(match.group(1)) if match else None


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
    if platform.system().lower() == "windows":
        return ["ping", "-t", ip]
    return ["ping", ip]

# ---------------- PING LOOP ----------------
def ping_loop(ip):
    global ping_process

    flags = subprocess.CREATE_NO_WINDOW if platform.system().lower() == "windows" else 0

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

    stop_ping()  # 🔴 ESKİ PING %100 KAPAT

    is_running = True
    current_ip = ip
    start_btn.config(text="Durdur")

    output_box.config(state=tk.NORMAL)
    output_box.delete("1.0", tk.END)
    output_box.config(state=tk.DISABLED)

    while not ui_queue.empty():
        ui_queue.get_nowait()

    ping_thread = threading.Thread(
        target=ping_loop,
        args=(ip,),
        daemon=True
    )
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


def toggle_ping():
    if is_running:
        stop_ping()
    else:
        start_ping()


def refresh_from_excel():
    global devices
    excel_devices = load_devices_from_excel()
    merged = []

    for ex in excel_devices:
        old = next((d for d in devices if d["ip"] == ex["ip"]), None)
        if old:
            merged.append(old)
        else:
            merged.append({
                "name": ex["name"],
                "ip": ex["ip"],
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
    start_ping()

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
    write_ip_from_selection()
    return "break"

# ---------------- UI ----------------
root = tk.Tk()
root.title("Ping Monitor")
root.geometry("1100x650")
root.minsize(1100, 650)

top = tk.Frame(root)
top.pack(fill=tk.X, padx=10, pady=5)

tk.Label(top, text="IP:").pack(side=tk.LEFT)
ip_entry = tk.Entry(top, width=25)
ip_entry.pack(side=tk.LEFT, padx=5)

start_btn = tk.Button(top, text="Başlat", width=10, command=toggle_ping)
start_btn.pack(side=tk.LEFT, padx=5)

tk.Button(top, text="Yenile", width=10, command=refresh_from_excel).pack(side=tk.LEFT)

main = tk.PanedWindow(root, orient=tk.HORIZONTAL)
main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

output_box = tk.Text(main, state=tk.DISABLED)
main.add(output_box)

right = tk.Frame(main)
main.add(right)

cols = ("Cihaz", "IP", "Ping (ms)", "Son Ping")
device_tree = ttk.Treeview(right, columns=cols, show="headings")

for c in cols:
    device_tree.heading(c, text=c)
    device_tree.column(c, width=150 if c != "Son Ping" else 200)

device_tree.pack(fill=tk.BOTH, expand=True)

# BINDINGS
device_tree.bind("<<TreeviewSelect>>", on_single_click)
device_tree.bind("<Double-1>", on_double_click)

# 🔴 Treeview'in kendi ok davranışını iptal et
device_tree.bind("<Up>", lambda e: "break")
device_tree.bind("<Down>", lambda e: "break")

# 🔑 Ok tuşları root’tan yönetiliyor
root.bind("<Up>", on_arrow_key)
root.bind("<Down>", on_arrow_key)

root.bind("<Return>", start_ping)
root.bind("<Escape>", stop_ping)

# RENKLER (Windows uyumlu)
device_tree.tag_configure("UNKNOWN", foreground="#7f8c8d")
device_tree.tag_configure("FAST", foreground="#1e8449")
device_tree.tag_configure("NORMAL", foreground="#27ae60")
device_tree.tag_configure("SLOW", foreground="#b7950b")
device_tree.tag_configure("VERY_SLOW", foreground="#ca6f1e")
device_tree.tag_configure("DOWN", foreground="#c0392b")

# ---------------- START ----------------
devices = load_devices()
refresh_device_list()
root.after(100, process_ui_queue)
root.mainloop()