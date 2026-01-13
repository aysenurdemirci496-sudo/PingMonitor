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
ui_queue = queue.Queue()

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
    return ["ping", "-t", ip] if platform.system().lower() == "windows" else ["ping", ip]

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

        # TEXT BOX AÇ → YAZ → KİLİTLE
        output_box.config(state=tk.NORMAL)
        output_box.insert(tk.END, line)
        output_box.see(tk.END)
        output_box.config(state=tk.DISABLED)

        ms = extract_ping_ms(line)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for d in devices:
            if d["ip"] == current_ip:
                d["last_ping"] = now
                d["latency"] = ms
                d["status"] = status_by_latency(ms)
                break

        save_devices(devices)
        refresh_device_list()

    root.after(100, process_ui_queue)

# ---------------- ACTIONS ----------------
def start_ping(event=None):
    global is_running, current_ip

    ip = ip_entry.get().strip()
    if not ip or is_running:
        return

    is_running = True
    current_ip = ip
    start_btn.config(text="Durdur")

    output_box.config(state=tk.NORMAL)
    output_box.delete("1.0", tk.END)
    output_box.config(state=tk.DISABLED)

    threading.Thread(target=ping_loop, args=(ip,), daemon=True).start()


def stop_ping(event=None):
    global is_running
    is_running = False
    start_btn.config(text="Başlat")

    # OK TUŞLARI KAYBOLMASIN
    device_tree.focus_set()


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
                "last_ping": None,
                "latency": None,
                "status": "UNKNOWN"
            })

    devices = merged
    save_devices(devices)
    refresh_device_list()

# ---------------- DEVICE LIST ----------------
def refresh_device_list():
    device_tree.delete(*device_tree.get_children())

    for d in devices:
        d.setdefault("status", "UNKNOWN")
        d.setdefault("latency", None)
        d.setdefault("last_ping", None)

        latency_txt = "-" if d["latency"] is None else f"{d['latency']:.1f}"

        device_tree.insert(
            "",
            tk.END,
            values=(d["name"], d["ip"], latency_txt, d["last_ping"] or "-"),
            tags=(d["status"],)
        )

    # 🔴 KRİTİK: OK TUŞLARI İÇİN
    children = device_tree.get_children()
    if children:
        if not device_tree.selection():
            device_tree.selection_set(children[0])
        device_tree.focus(children[0])
        device_tree.focus_set()
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
    items = device_tree.get_children()
    if not items:
        return "break"

    selected = device_tree.selection()
    index = items.index(selected[0]) if selected else 0

    if event.keysym == "Down" and index < len(items) - 1:
        index += 1
    elif event.keysym == "Up" and index > 0:
        index -= 1
    else:
        return "break"

    device_tree.selection_set(items[index])
    device_tree.focus(items[index])
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

# SOL (READ-ONLY OUTPUT)
output_box = tk.Text(main, state=tk.DISABLED)
main.add(output_box)

# SAĞ (DEVICES)
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
device_tree.bind("<Up>", on_arrow_key)
device_tree.bind("<Down>", on_arrow_key)

root.bind("<Return>", start_ping)
root.bind("<Escape>", stop_ping)

# RENKLER
device_tree.tag_configure("UNKNOWN", background="#e0e0e0")
device_tree.tag_configure("FAST", background="#2ecc71")
device_tree.tag_configure("NORMAL", background="#a9dfbf")
device_tree.tag_configure("SLOW", background="#f9e79f")
device_tree.tag_configure("VERY_SLOW", background="#f5b041")
device_tree.tag_configure("DOWN", background="#f1948a")

# ---------------- START ----------------
devices = load_devices()
refresh_device_list()
root.after(100, process_ui_queue)
root.mainloop()