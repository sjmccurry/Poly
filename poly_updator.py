import os


try:
    import tkinter as tk
except:
    os.system('pip3 install tkinter')
    import tkinter as tk
try:
    import requests
except:
    os.system('pip3 install requests')
import subprocess
import sys

import tkinter as tk
import time


VERSION_URL = "https://raw.githubusercontent.com/sjmccurry/Poly/refs/heads/main/version.txt"
SCRIPT_URL = "https://raw.githubusercontent.com/sjmccurry/Poly/refs/heads/main/poly_gui.py"

LOCAL_SCRIPT = "poly_gui.py"
LOCAL_VERSION_FILE = "version.txt"

# Create basic GUI window
root = tk.Tk()
root.title("Poly Updater")
root.geometry("400x120")
root.configure(bg="#f8f9fa")
label = tk.Label(root, text="Checking for updates...", font=("Segoe UI", 11), bg="#f8f9fa")
label.pack(pady=35)
root.update()

def show_message(msg):
    label.config(text=msg)
    root.update()

def show_error_and_exit(msg):
    label.config(text=msg, fg="red")
    root.update()
    time.sleep(3)
    root.destroy()
    sys.exit(1)

def get_local_version():
    if not os.path.exists(LOCAL_VERSION_FILE):
        return "0.0.0"
    with open(LOCAL_VERSION_FILE, "r") as f:
        return f.read().strip()

def get_remote_version():
    try:
        r = requests.get(VERSION_URL)
        r.raise_for_status()
        return r.text.strip()
    except:
        return None

def is_newer(remote, local):
    return tuple(map(int, remote.split("."))) > tuple(map(int, local.split(".")))

def download_and_overwrite_script():
    try:
        for i in range(5):
            try:
                if os.path.exists(LOCAL_SCRIPT):
                    os.remove(LOCAL_SCRIPT)
                break
            except PermissionError:
                time.sleep(1)
        else:
            raise PermissionError(f"Could not delete {LOCAL_SCRIPT}")

        r = requests.get(SCRIPT_URL)
        r.raise_for_status()
        with open(LOCAL_SCRIPT, "w", encoding="utf-8") as f:
            f.write(r.text)
    except Exception as e:
        show_error_and_exit(f"Update failed:\n{str(e)}")

def write_local_version(version):
    with open(LOCAL_VERSION_FILE, "w") as f:
        f.write(version)

def launch_poly():
    show_message("Launching Poly...")
    root.update()
    subprocess.Popen([sys.executable, LOCAL_SCRIPT], creationflags=subprocess.CREATE_NO_WINDOW)
    root.after(500, root.destroy)

def main():
    show_message("Checking for updates...")
    local_version = get_local_version()
    remote_version = get_remote_version()

    if remote_version and is_newer(remote_version, local_version):
        show_message(f"Updating to {remote_version}...")
        download_and_overwrite_script()
        write_local_version(remote_version)
    else:
        show_message("Already up-to-date.")

    launch_poly()

root.after(100, main)
root.mainloop()
