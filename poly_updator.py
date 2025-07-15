import requests
import subprocess
import sys
import os

VERSION_URL = "https://raw.githubusercontent.com/sjmccurry/Poly/refs/heads/main/version"
SCRIPT_URL = "https://raw.githubusercontent.com/sjmccurry/Poly/refs/heads/main/poly_gui.py"
LOCAL_SCRIPT = "poly_gui.py"
LOCAL_VERSION_FILE = "version.txt"

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
        r = requests.get(SCRIPT_URL)
        r.raise_for_status()
        with open(LOCAL_SCRIPT, "w", encoding="utf-8") as f:
            f.write(r.text)
        print(f"[Poly Updater] Downloaded and updated '{LOCAL_SCRIPT}'.")
    except Exception as e:
        print(f"[Poly Updater] Failed to download script: {e}")
        sys.exit(1)

def write_local_version(version):
    with open(LOCAL_VERSION_FILE, "w") as f:
        f.write(version)

def launch_poly():
    print("[Poly Updater] Launching updated Poly...")
    subprocess.Popen([sys.executable, LOCAL_SCRIPT])
    sys.exit()

def main():
    print("[Poly Updater] Checking for updates...")
    local_version = get_local_version()
    remote_version = get_remote_version()

    if remote_version and is_newer(remote_version, local_version):
        print(f"[Poly Updater] New version available: {remote_version} (current: {local_version})")
        download_and_overwrite_script()
        write_local_version(remote_version)
    else:
        print("[Poly Updater] Already up-to-date.")

    launch_poly()

if __name__ == "__main__":
    main()
