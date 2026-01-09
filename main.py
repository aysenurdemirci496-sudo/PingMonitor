import tkinter as tk
from tkinter import ttk, messagebox
import threading
import subprocess
import queue
import platform
import time

from device_loader import load_devices, update_last_ping

# ---------------- GLOBAL ----------------
is_running = False
ping_process = None
output_queue = queue.Queue()
devices = []


# ---------------- PING ----------------
def get_ping_command(ip):
    if platform.system().lower() == "windows":
        return ["ping", "-t", ip]
    else:
        return ["ping", ip]


def ping_loop(ip):
    global ping_process

    try:
        cmd = get_ping_command(ip)

        creationflags = 0
        if platform.system().lower() == "windows":
            creationflags = subprocess.CREATE_NO_WINDOW

        ping_process = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=creationflags
        )

        for line in ping_process.stdout:
            if not is_running:
                break
            output_queue.put(line)

    except Exception as e:
        output_queue.put(f"HATA: {e}\n")


def stop_ping():
    global ping_process
    if ping_process:
        try:
            ping_process.terminate()
        except Exception:
            pass
        ping_process = None


# ---------------- UI CALLBACKS ----------------
def toggle_ping():
    global is_running

    ip = ip_entry.get().strip()
    if not ip:
        messagebox.showwarning("Uyarı", "IP adresi gir")
        return

    if not is_running:
        is_running = True
        start_button.config(text="Durdur")
        output_box.delete("1.0", tk.END)

        threading.Thread(
            target=ping_loop,
            args=(ip,),
            daemon=True
        ).start()

        update_last_ping(devices, ip)
        refresh_device_list()

    else:
        is_running = False
        stop_ping()
        start_button.config(text="Başlat")


def process_output():
    while not output_queue.empty():
        line = output_queue.get()
        output_box.insert(tk.END, line)
        output_box.see(tk.END)

    root.after(100, process_output)


def refresh_device_list():
    device_tree.delete(*device_tree.get_children())
    for d in devices:
        last = d["last_ping"] if d["last_ping"] else "-"
        device_tree.insert("", tk.END, values=(d["name"], d["ip"], last))


def on_device_select(event):
    selected = device_tree.selection()
    if not selected:
        return
    item = device_tree.item(selected[0])
    ip_entry.delete(0, tk.END)
    ip_entry.insert(0, item["values"][1])


# ---------------- UI ----------------
root = tk.Tk()
root.title("Ping Monitor")
root.minsize(1100, 650)
root.update_idletasks()
root.geometry("1100x650")

top = tk.Frame(root)
top.pack(fill=tk.X, padx=10, pady=5)

tk.Label(top, text="IP:").pack(side=tk.LEFT)
ip_entry = tk.Entry(top, width=25)
ip_entry.pack(side=tk.LEFT, padx=5)

start_button = tk.Button(top, text="Başlat", width=10, command=toggle_ping)
start_button.pack(side=tk.LEFT, padx=5)

refresh_button = tk.Button(top, text="Yenile", width=10)
refresh_button.pack(side=tk.LEFT, padx=5)

main = tk.PanedWindow(root, sashrelief=tk.RAISED)
main.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

# Output
output_box = tk.Text(main)
main.add(output_box)

# Devices
right = tk.Frame(main)
main.add(right)

tk.Label(right, text="Cihazlar").pack(anchor="w")

cols = ("Cihaz", "IP", "Son Ping")
device_tree = ttk.Treeview(right, columns=cols, show="headings", height=25)
for c in cols:
    device_tree.heading(c, text=c)
    device_tree.column(c, width=180 if c != "Son Ping" else 220)

device_tree.pack(fill=tk.BOTH, expand=True)
device_tree.bind("<<TreeviewSelect>>", on_device_select)

# ---------------- INIT ----------------
devices = load_devices()
refresh_device_list()
root.after(100, process_output)

root.mainloop()