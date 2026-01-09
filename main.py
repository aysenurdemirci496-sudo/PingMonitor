import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import queue
import json
from datetime import datetime

from device_loader import load_devices, merge_devices_from_excel


# ======================
# STATE
# ======================
devices = []
current_device_index = None
is_running = False
ping_process = None
ui_queue = queue.Queue()


# ======================
# PING THREAD
# ======================
def ping_loop(ip):
    global ping_process, is_running

    ping_process = subprocess.Popen(
        ["ping", ip],
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )

    for line in ping_process.stdout:
        if not is_running:
            break

        ui_queue.put(line)

        if ("Reply from" in line) or ("bytes from" in line):
            if current_device_index is not None:
                devices[current_device_index]["last_ping"] = (
                    datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                )
                with open("devices.json", "w", encoding="utf-8") as f:
                    json.dump(devices, f, indent=2, ensure_ascii=False)

                ui_queue.put("__REFRESH__")

    ping_process.terminate()
    ping_process = None


# ======================
# UI QUEUE
# ======================
def process_ui_queue():
    try:
        while True:
            item = ui_queue.get_nowait()
            if item == "__REFRESH__":
                refresh_device_list()
            else:
                output_box.insert(tk.END, item)
                output_box.see(tk.END)
    except queue.Empty:
        pass

    root.after(30, process_ui_queue)


# ======================
# TOGGLE BUTTON
# ======================
def toggle_ping():
    global is_running, current_device_index

    if not is_running:
        ip = ip_entry.get().strip()
        if not ip:
            messagebox.showerror("Hata", "IP boş")
            return

        if current_device_index is None:
            for i, d in enumerate(devices):
                if d["ip"] == ip:
                    current_device_index = i
                    break

        if current_device_index is None:
            messagebox.showerror("Hata", "IP listede yok")
            return

        is_running = True
        start_button.config(text="Durdur")
        threading.Thread(target=ping_loop, args=(ip,), daemon=True).start()

    else:
        is_running = False
        start_button.config(text="Başlat")


# ======================
# DEVICE SELECT
# ======================
def on_device_select(event):
    global current_device_index

    sel = device_tree.selection()
    if not sel:
        return

    values = device_tree.item(sel[0], "values")
    ip = values[1]

    for i, d in enumerate(devices):
        if d["ip"] == ip:
            current_device_index = i
            break

    ip_entry.delete(0, tk.END)
    ip_entry.insert(0, ip)

    if is_running:
        toggle_ping()
        root.after(100, toggle_ping)


# ======================
# REFRESH
# ======================
def refresh_from_excel():
    global devices, current_device_index, is_running

    if is_running:
        is_running = False
        start_button.config(text="Başlat")

    devices = merge_devices_from_excel()
    current_device_index = None
    ip_entry.delete(0, tk.END)

    for sel in device_tree.selection():
        device_tree.selection_remove(sel)

    refresh_device_list()
    messagebox.showinfo("Bilgi", "Cihaz listesi güncellendi")


# ======================
# TREEVIEW
# ======================
def refresh_device_list():
    device_tree.delete(*device_tree.get_children())
    for d in devices:
        device_tree.insert("", tk.END, values=(
            d["name"],
            d["ip"],
            d["last_ping"] or "-"
        ))


# ======================
# UI
# ======================
root = tk.Tk()
root.title("Ping Monitor")

# Önce minimum boyut
root.minsize(1000, 600)

# Layout hesaplat
root.update_idletasks()

# Sonra kesin boyut ver
root.geometry("1100x650")

top = tk.Frame(root)
top.pack(fill=tk.X, padx=10, pady=5)

tk.Label(top, text="IP:").pack(side=tk.LEFT)
ip_entry = tk.Entry(top, width=20)
ip_entry.pack(side=tk.LEFT, padx=5)

start_button = tk.Button(top, text="Başlat", width=10, command=toggle_ping)
start_button.pack(side=tk.LEFT, padx=5)

tk.Button(top, text="Yenile", command=refresh_from_excel).pack(side=tk.LEFT, padx=5)

middle = tk.Frame(root)
middle.pack(fill=tk.BOTH, expand=True, padx=10)

output_box = tk.Text(middle)
output_box.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

right = tk.Frame(middle, width=450)
right.pack(side=tk.RIGHT, fill=tk.Y)
right.pack_propagate(False)

tk.Label(right, text="Cihazlar").pack()

device_tree = ttk.Treeview(
    right,
    columns=("name", "ip", "last_ping"),
    show="headings"
)
device_tree.heading("name", text="Cihaz")
device_tree.heading("ip", text="IP")
device_tree.heading("last_ping", text="Son Ping")

device_tree.column("name", width=120)
device_tree.column("ip", width=110)
device_tree.column("last_ping", width=200)

device_tree.pack(fill=tk.BOTH, expand=True)
device_tree.bind("<<TreeviewSelect>>", on_device_select)


# ======================
# START
# ======================
devices = load_devices()
refresh_device_list()
process_ui_queue()
root.mainloop()