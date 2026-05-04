import subprocess
import sys
import time
import os
from datetime import datetime

SCRIPT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "main.py")
LOG    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "launcher.log")
PYTHON = sys.executable.replace("pythonw.exe", "python.exe")

def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

restart_count = 0
BACKOFF = [5, 10, 30, 60, 120]

while True:
    try:
        log(f"Starting bot (restart #{restart_count})...")
        env = os.environ.copy()
        env["PYTHONIOENCODING"] = "utf-8"
        proc = subprocess.Popen(
            [PYTHON, SCRIPT],
            cwd=os.path.dirname(SCRIPT),
            stdout=open(os.path.join(os.path.dirname(SCRIPT), "bot.log"), "a", encoding="utf-8"),
            stderr=open(os.path.join(os.path.dirname(SCRIPT), "bot_err.log"), "a", encoding="utf-8"),
            env=env,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        proc.wait()
        exit_code = proc.returncode
        log(f"Bot exited with code {exit_code}.")
    except Exception as e:
        log(f"Launcher error: {e}")

    restart_count += 1
    delay = BACKOFF[min(restart_count - 1, len(BACKOFF) - 1)]
    log(f"Restarting in {delay}s...")
    time.sleep(delay)
