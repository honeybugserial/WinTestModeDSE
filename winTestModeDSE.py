#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Usage:
    script.exe                           (default: enable test mode + disable DSE)
    script.exe --mode enable
    script.exe --mode disable
    script.exe --auto-accept
    script.exe --auto-reboot

Modes:
    enable  → testsigning=on,  nointegritychecks=on
    disable → testsigning=off, nointegritychecks=off

--auto-accept:
    Skips user prompts at start and end.

--auto-reboot:
    Automatically reboots the system after changes are applied.
    
Pyinstaller:
    pyinstaller --onefile --icon=configuration.ico  --clean --uac-admin winTestMode+DSE.py
"""

import os
import sys
import time
import shutil
import subprocess
import ctypes
import shlex
import ctypes
import ctypes.wintypes as wt
from pathlib import Path

# ------------------------------
# Rich TUI Setup
# ------------------------------
from rich.console import Console
from rich.panel import Panel
from rich.text import Text
from rich.traceback import install
from rich.rule import Rule

console = Console()

WIN_DIR = Path(os.environ.get("WINDIR", r"C:\Windows"))
BCDEDIT = WIN_DIR / "System32" / "bcdedit.exe"
SHUTDOWN = WIN_DIR / "System32" / "shutdown.exe"
MSIEXEC = WIN_DIR / "System32" / "msiexec.exe"

# ------------------------------
# Arch detection for Wow64 redirection
# ------------------------------
IS_64BIT = (os.environ.get("PROCESSOR_ARCHITECTURE", "").lower() == "amd64")

# ------------------------------
# Flag parsing
# ------------------------------
VALID_FLAGS = {"--mode", "--auto-accept", "--auto-reboot"}

MODE = "enable"
AUTO_ACCEPT = False
AUTO_REBOOT = False

args = sys.argv[1:]
i = 0

while i < len(args):
    arg = args[i]

    if arg not in VALID_FLAGS and not arg.startswith("--mode="):
        console.print(f"[bold red]Invalid flag:[/bold red] {arg}")
        sys.exit(1)

    if arg == "--auto-accept":
        AUTO_ACCEPT = True

    elif arg == "--auto-reboot":
        AUTO_REBOOT = True

    elif arg == "--mode":
        if i + 1 >= len(args):
            console.print("[bold red]--mode requires: enable | disable[/bold red]")
            sys.exit(1)
        MODE = args[i + 1].strip().lower()
        i += 1

    elif arg.startswith("--mode="):
        MODE = arg.split("=", 1)[1].strip().lower()

    i += 1

if MODE not in ("enable", "disable"):
    console.print("[bold red]--mode must be: enable | disable[/bold red]")
    sys.exit(1)


def info(msg):
    console.print(f"[bold cyan][ # ] [/bold cyan] {msg}")
    time.sleep(0.15)

def ok(msg):
    console.print(f"[bold green][ + ] [/bold green] {msg}")
    time.sleep(0.15)

def warn(msg):
    console.print(f"[bold yellow][WARN] [/bold yellow] {msg}")
    time.sleep(0.15)

def error(msg):
    console.print(f"[bold red][ERROR] [/bold red] {msg}")
    time.sleep(0.15)

def fatal(msg, code=1):
    error(msg)
    sys.exit(code)


# Working directory logic
if getattr(sys, "frozen", False):
    WORKDIR = Path(sys.executable).parent
else:
    WORKDIR = Path.cwd()

LOG = WORKDIR / "etka-post-setup.log"


# ------------------------------
# Logging to file
# ------------------------------
def filelog(msg):
    try:
        LOG.parent.mkdir(parents=True, exist_ok=True)
        ts = time.strftime("%m-%d-%Y#%H:%M")
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(f"{ts}$: {msg}\n")
    except Exception:
        pass

def log(msg):
    filelog(msg)
    info(msg)


# ------------------------------
# Admin / Elevation
# ------------------------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def relaunch_elevated():
    params = " ".join([shlex.quote(a) for a in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, params, None, 1
    )


# ------------------------------
# WOW64 Redirection (32/64 aware)
# ------------------------------
class Wow64DisableRedirection:
    def __enter__(self):
        if not IS_64BIT:
            self._active = False
            return self

        self._old = ctypes.c_void_p()
        self._disable = getattr(ctypes.windll.kernel32, "Wow64DisableWow64FsRedirection", None)
        self._revert = getattr(ctypes.windll.kernel32, "Wow64RevertWow64FsRedirection", None)

        if self._disable and self._disable(ctypes.byref(self._old)):
            self._active = True
        else:
            self._active = False

        return self

    def __exit__(self, exc_type, exc, tb):
        if IS_64BIT and self._active:
            self._revert(self._old)


# ------------------------------
# Command Execution
# ------------------------------
def run(cmd, check=False, show=False):
    if show:
        console.print(f"[magenta]$ {cmd}[/magenta]")
        time.sleep(0.10)

    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.run(
        cmd,
        shell=True,
        text=True,
        capture_output=True,
        creationflags=creationflags,
    )

    if proc.stdout:
        for line in proc.stdout.splitlines():
            console.print(f"[grey50]{line}[/grey50]")
        time.sleep(0.10)

    if proc.stderr:
        for line in proc.stderr.splitlines():
            console.print(f"[red]{line}[/red]")
        time.sleep(0.10)

    if check and proc.returncode != 0:
        fatal(f"Command failed: {cmd}\nRC={proc.returncode}")

    return proc.returncode, proc.stdout, proc.stderr


def run_sys32(cmd, check=False, show=False):
    with Wow64DisableRedirection():
        return run(cmd, check=check, show=show)


# ------------------------------
# BCD / Test Mode / DSE
# ------------------------------
VALID_ON = ("on", "yes", "true", "1", "enabled")

def bcd_query():
    for cmd in (f'"{BCDEDIT}" /enum {{current}}', f'"{BCDEDIT}" /enum'):
        rc, out, err = run_sys32(cmd)
        if rc == 0:
            break

    if rc != 0:
        fatal("Unable to query BCD state.")

    low = out.lower()

    def get_state(keyword):
        for line in low.splitlines():
            if keyword in line:
                val = line.split()[-1].strip().lower()
                return val in VALID_ON
        return False

    return get_state("testsigning"), get_state("nointegritychecks")


def set_flag(flag, state):
    state_str = "on" if state else "off"
    cmd = f'"{BCDEDIT}" /set {flag} {state_str}'
    rc, _, _ = run_sys32(cmd, show=True)
    if rc != 0:
        error(f"Failed to set BCD flag: {flag} -> {state_str}")
    else:
        ok(f"BCD flag updated: {flag} = {state_str}")


# ------------------------------
# Ensuring Test Mode / DSE (Mode-Aware)
# ------------------------------
def ensure_testmode_and_dse(mode):
    """
    mode = enable  → testsigning=on,  nointegritychecks=on
    mode = disable → testsigning=off, nointegritychecks=off
    """

    desired_state = (mode == "enable")

    console.line()
    console.print(Rule(f"[yellow]--- Verifying Test Mode + DSE state ({mode}) ---[/yellow]"))
    time.sleep(0.6)

    info("Checking BCD flags")

    before_t, before_n = bcd_query()
    console.print(f"[cyan] - testsigning:[/cyan] {before_t}")
    console.print(f"[cyan] - nointegritychecks:[/cyan] {before_n}")
    time.sleep(0.3)

    changed = False

    if before_t != desired_state:
        warn(f"Testsigning is {'OFF' if desired_state else 'ON'} — adjusting")
        time.sleep(0.2)
        set_flag("testsigning", desired_state)
        changed = True
        time.sleep(0.3)
    else:
        ok("Testsigning already correct.")
        time.sleep(0.2)

    if before_n != desired_state:
        warn(f"NoIntegrityChecks is {'OFF' if desired_state else 'ON'} — adjusting")
        time.sleep(0.2)
        set_flag("nointegritychecks", desired_state)
        changed = True
        time.sleep(0.3)
    else:
        ok("NoIntegrityChecks already correct.")
        time.sleep(0.2)

    info("Rechecking BCD flags")
    after_t, after_n = bcd_query()

    console.print(f"[cyan] - testsigning:[/cyan] {after_t}")
    console.print(f"[cyan] - nointegritychecks:[/cyan] {after_n}")
    time.sleep(0.3)

    if after_t != desired_state or after_n != desired_state:
        fatal("Windows rejected Test Mode or DSE changes.")

    ok("Test Mode + DSE state verified.")
    time.sleep(0.3)

    return changed


# ------------------------------
# Main
# ------------------------------
def main():

    if not is_admin():
        relaunch_elevated()
        return
        
    console.print(f" _ _ _  _         ___             _      __ __         _         ___  ___  ___    ___            _  ")
    console.print(f"| | | |[_] _ _   |_ _| ___  ___ _| |_   |  \  \ ___  _| | ___   | . \/ __]| __]  |_ _| ___  ___ | | ")
    console.print(f"| | | || || ' |   | | / ._][_-[  | |    |     |/ . \/ . |/ ._]  | | |\__ \| _]    | | / . \/ . \| | ")
    console.print(f"|__/_/ |_||_|_|   |_| \___./__/  |_|    |_|_|_|\___/\___|\___.  |___/[___/|___]   |_| \___/\___/|_| ")

    console.print(f"----------------------------------------------------------------------------------------------------")
    console.line()
    
    if not AUTO_ACCEPT:
        choice = input(f"Apply Mode '{MODE}'? (Y/N): ").strip().lower()
        if choice != "y":
            warn("User abortion. Exiting.")
            return

    changed = ensure_testmode_and_dse(MODE)

    if changed:
        warn("System must reboot to apply changes.")

        if AUTO_REBOOT:
            ok("Rebooting in 5 seconds...")
            run_sys32(f'"{SHUTDOWN}" /r /t 5 /f', show=True)
            return

    console.line()
    console.print(Panel("[bold green]--- Completed Successfully! ---[/bold green]", border_style="green"))

    if not AUTO_ACCEPT:
        console.print("\n[bold cyan]Press ENTER to exit…[/bold cyan]")
        input()


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        fatal(str(e))
