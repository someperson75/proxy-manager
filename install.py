import subprocess
from pathlib import Path
import sys
import urllib.request

URL = "https://raw.githubusercontent.com/someperson75/proxy-manager/refs/heads/main/main.py"

DEST_DIR = Path.home() / ".proxy"
SCRIPT = DEST_DIR / "main.py"

PYTHONW = Path(sys.executable).with_name("pythonw.exe")

TASK_NAME = "Proxy WiFi Switch"


def create_task():
    cmd = [
        "schtasks",
        "/Create",
        "/F",
        "/SC", "ONLOGON",
        "/RL", "LIMITED",
        "/TN", TASK_NAME,
        "/TR", f'"{PYTHONW}" "{SCRIPT}"'
    ]

    subprocess.run(" ".join(cmd), shell=True, check=True)
    print("Task created successfully")


def install():
    DEST_DIR.mkdir(parents=True, exist_ok=True)

    print("Downloading main.py...")
    with urllib.request.urlopen(URL) as response:
        content = response.read().decode("utf-8")

    SCRIPT.write_text(content, encoding="utf-8")
    print(f"Installed to {SCRIPT}")


if __name__ == "__main__":
    install()
    create_task()
