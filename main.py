import ctypes
import json
import re
import subprocess
import winreg
from pathlib import Path

import pythoncom
import win32com.client
import tkinter as tk

# =========================
# Constants
# =========================

REG_PATH = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings"
CONFIG_DIR = Path.home() / ".proxy"
CONFIG_FILE = CONFIG_DIR / "config.json"

INTERNET_OPTION_SETTINGS_CHANGED = 39
INTERNET_OPTION_REFRESH = 37
CREATE_NO_WINDOW = 0x08000000  # subprocess.CREATE_NO_WINDOW (Py3.8 safe)


# =========================
# Proxy management
# =========================

def _refresh_proxy():
    """Notify Windows that proxy settings changed."""
    ctypes.windll.Wininet.InternetSetOptionW(0, INTERNET_OPTION_SETTINGS_CHANGED, 0, 0)
    ctypes.windll.Wininet.InternetSetOptionW(0, INTERNET_OPTION_REFRESH, 0, 0)


def activate_proxy(proxy_ip_port: str):
    """Enable system proxy (WinINET)."""
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        REG_PATH,
        0,
        winreg.KEY_SET_VALUE
    ) as key:
        winreg.SetValueEx(key, "ProxyEnable", 0, winreg.REG_DWORD, 1)
        winreg.SetValueEx(key, "ProxyServer", 0, winreg.REG_SZ, proxy_ip_port)

    _refresh_proxy()


def deactivate_proxy():
    """Disable system proxy (WinINET)."""
    with winreg.OpenKey(
        winreg.HKEY_CURRENT_USER,
        REG_PATH,
        0,
        winreg.KEY_SET_VALUE
    ) as key:
        winreg.SetValueEx(key, "ProxyEnable", 0, winreg.REG_DWORD, 0)

    _refresh_proxy()


# =========================
# Wi-Fi detection
# =========================

def get_wifi_ids():
    """
    Returns (ssid, bssid) or None if not connected to Wi-Fi.
    """
    output = subprocess.check_output(
        ["netsh", "wlan", "show", "interfaces"],
        encoding="utf-8",
        errors="ignore",
        creationflags=CREATE_NO_WINDOW
    )

    ssid = None
    bssid = None

    m = re.search(r"^\s*SSID\s*:\s*(.*)$", output, re.MULTILINE)
    if m:
        ssid = m.group(1).strip() or None

    m = re.search(r"^\s*BSSID\s*:\s*(.+)$", output, re.MULTILINE)
    if m:
        bssid = m.group(1).strip()

    return (ssid, bssid) if bssid else None


# =========================
# Config management
# =========================

def load_config():
    """Load config.json as dict."""
    if not CONFIG_FILE.exists():
        return {}

    with CONFIG_FILE.open("r", encoding="utf-8") as f:
        return json.load(f)


def save_config(config: dict):
    """Save dict to config.json."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)

    with CONFIG_FILE.open("w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)


# =========================
# UI
# =========================

def popup_input(title: str, message: str):
    """
    Shows a modal popup with text + input + Save button.
    Returns entered value or None.
    """
    result:dict[str,str] = {"value": None}

    def on_save():
        result["value"] = entry.get().strip() or None
        root.destroy()

    root = tk.Tk()
    root.title(title)
    root.resizable(False, False)

    width, height = 400, 120
    x = (root.winfo_screenwidth() - width) // 2
    y = (root.winfo_screenheight() - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")

    tk.Label(root, text=message).pack(pady=(10, 5))

    entry = tk.Entry(root, width=40)
    entry.pack(pady=5)
    entry.focus()

    tk.Button(root, text="Save", command=on_save).pack(pady=(5, 10))

    root.mainloop()
    return result["value"]


# =========================
# Wi-Fi change handler
# =========================

def on_wifi_change():
    config = load_config()
    wifi = get_wifi_ids()

    if not wifi:
        return

    ssid, bssid = wifi
    proxy = config.get(bssid, "null")

    if proxy != "null":
        if proxy:
            activate_proxy(proxy)
        else:
            deactivate_proxy()
        return

    # New Wi-Fi detected
    ans = popup_input(
        "New Wi-Fi detected",
        f"Enter proxy for {ssid or 'this network'} (ip:port), leave blank for no proxy:"
    )

    config[bssid] = ans  # None = no proxy

    if ans:
        activate_proxy(ans)
    else:
        deactivate_proxy()

    save_config(config)


# =========================
# Event listener
# =========================

def listen_wifi_changes():
    pythoncom.CoInitialize()

    locator = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    svc = locator.ConnectServer(".", "root\\cimv2")

    query = (
        "SELECT * FROM __InstanceModificationEvent "
        "WITHIN 1 "
        "WHERE TargetInstance ISA 'Win32_NetworkAdapter' "
        "AND TargetInstance.NetConnectionStatus IS NOT NULL"
    )

    watcher = svc.ExecNotificationQuery(query)

    # print("Listening for Wi-Fi changes...")

    while True:
        watcher.NextEvent()
        on_wifi_change()


# =========================
# Entry point
# =========================

if __name__ == "__main__":
    listen_wifi_changes()
