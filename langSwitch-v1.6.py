#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import platform
import sys
import subprocess
from urllib.parse import unquote
import urllib.request
import urllib.error
import time
import random

import questionary
from questionary import Style

from rich.console import Console
from rich.text import Text
from rich.panel import Panel
from rich.progress import (
    Progress,
    BarColumn,
    TextColumn,
    TimeRemainingColumn,
    DownloadColumn,
)
from rich.table import Table

COMMON_LANGS = [
    "en-us",
    "es-es",
    "fr-fr",
    "de-de",
    "it-it",
    "pt-br",
    "ru-ru",
    "zh-cn",
    "zh-tw",
    "ja-jp",
    "ko-kr",
]

LANG_LABELS = {
    "en-us": "English (United States)",
    "es-es": "Spanish (Spain)",
    "fr-fr": "French (France)",
    "de-de": "German (Germany)",
    "it-it": "Italian (Italy)",
    "pt-br": "Portuguese (Brazil)",
    "ru-ru": "Russian (Russia)",
    "zh-cn": "Chinese (Simplified)",
    "zh-tw": "Chinese (Traditional)",
    "ja-jp": "Japanese",
    "ko-kr": "Korean",
}

MENU_STYLE = Style([
    ("qmark", "fg:cyan bold"),
    ("question", "bold"),
    ("answer", "fg:green bold"),
    ("pointer", "fg:yellow bold"),        # the arrow itself
    ("highlighted", "fg:yellow bold"),    # highlighted row text
    ("selected", "fg:green"),
    ("separator", "fg:ansiwhite"),
])

# ============================================================
# Path handling
# ============================================================

def get_exe_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


EXE_DIR = get_exe_dir()

# ============================================================
# Console / Logger
# ============================================================

console = Console()

class ConsoleLogger:
    LEVEL_STYLES = {
        "SUCCESS": "bold green",
        "FAIL": "bold red",
        "WARN": "bold yellow",
        "INFO": "cyan",
        "DEBUG": "dim",
    }

    DELAY_RANGE = (0.1, 0.25)

    def __init__(self, console: Console):
        self.console = console

    def _delay(self):
        time.sleep(random.uniform(*self.DELAY_RANGE))

    def log(self, level: str, message: str):
        style = self.LEVEL_STYLES.get(level.upper(), "white")
        tag = f"[{level.upper()}]".ljust(10)
        text = Text.assemble((tag, style), (" ", None), (message, None))
        self.console.print(text)
        self._delay()

    def panel(self, text: str, *, title=None, color="cyan"):
        self.console.print(Panel(text, title=title, border_style=color))
        self._delay()

logger = ConsoleLogger(console)
log = logger.log
panel = logger.panel

# ============================================================
# Exit handling (never auto-close)
# ============================================================

def pause_exit(code=0):
    console.print("\n[bold cyan]Press Enter to exit...[/bold cyan]")
    input()
    sys.exit(code)

def fatal(msg, code=1):
    log("FAIL", msg)
    pause_exit(code)

# ============================================================
# Args
# ============================================================

DRY_RUN = "--dry-run" in sys.argv

# ============================================================
# Input Files / Paths
# ============================================================

INPUT_FILES = [
    EXE_DIR / "win10LangExpPacks.dat",
    EXE_DIR / "win10FoD.dat",
    EXE_DIR / "win10LangOpts.dat",
]

DOWNLOAD_ROOT = EXE_DIR / "downloads"
APPX_DIR = DOWNLOAD_ROOT / "appx"
CAB_DIR = DOWNLOAD_ROOT / "cab"

FOD_FEATURE_KEYWORDS = [
    "languagefeatures-basic-",
    "languagefeatures-ocr-",
    "languagefeatures-texttospeech-",
    "languagefeatures-fonts-",
]

# ============================================================
# Internet check
# ============================================================

def check_internet():
    log("INFO", "Checking internet connectivity")
    try:
        req = urllib.request.Request(
            "https://archive.org",
            method="HEAD",
            headers={"User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=5):
            pass
        log("SUCCESS", "Internet connectivity OK")
    except Exception:
        fatal("No internet connection detected")

# ============================================================
# Helpers
# ============================================================

def load_all_lines(paths):
    lines = []
    for p in paths:
        if not p.exists():
            fatal(f"Missing file: {p}")
        lines.extend(
            l.strip() for l in p.read_text(encoding="utf-8").splitlines() if l.strip()
        )
    return list(dict.fromkeys(lines))

def detect_arch():
    m = platform.machine().lower()
    if m in ("amd64", "x86_64"):
        return "x64", "amd64"
    if m in ("x86", "i386", "i686"):
        return "x86", "x86"
    if m in ("arm64", "aarch64"):
        return "arm64", "arm64"
    fatal(f"Unsupported architecture: {m}")

def extract_languages(lines):
    langs = set()
    for l in lines:
        u = unquote(l).lower()
        if "/localexperiencepack/" in u:
            parts = u.split("/localexperiencepack/")[1].split("/")
            if len(parts) > 1:
                langs.add(parts[0])
    return sorted(langs)

def is_winpe(l):
    return "winpe" in l or "windows preinstallation environment" in l

def run(cmd):
    log("INFO", f"Running: {' '.join(cmd)}")

    p = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
    )

    for line in p.stdout:
        line = line.rstrip()
        if line:
            console.print(f"[dim]{line}[/dim]")

    ret = p.wait()
    if ret != 0:
        fatal(f"Command failed with exit code {ret}")


# ============================================================
# UI Helpers
# ============================================================
    
def show_language_table(langs):
    table = Table(title="Available Languages")
    table.add_column("Index", style="cyan", justify="right", no_wrap=True)
    table.add_column("Language", style="bold", no_wrap=True)

    for i, lang in enumerate(langs, 1):
        table.add_row(str(i), lang)

    console.print(table)


def select_language(langs):
    langs = sorted(set(langs))
    common = [l for l in COMMON_LANGS if l in langs]
    others = [l for l in langs if l not in common]

    # build label → value mapping
    def label(l):
        return f"{l} — {LANG_LABELS.get(l, 'Unknown')}"

    common_choices = {label(l): l for l in common}
    all_choices = {label(l): l for l in langs}

    while True:
        choice = questionary.select(
            "Select language to install:",
            choices=[
                *common_choices.keys(),
                questionary.Separator(),
                "Show all languages",
                "Exit",
            ],
            style=MENU_STYLE,
        ).ask()

        if choice in common_choices:
            return common_choices[choice]

        if choice == "Show all languages":
            sub = questionary.select(
                "All available languages:",
                choices=[*all_choices.keys(), "Back"],
                style=MENU_STYLE,
            ).ask()

            if sub and sub != "Back":
                return all_choices[sub]

        if choice in (None, "Exit"):
            fatal("Cancelled by user")


# ============================================================
# Download with Progress
# ============================================================

def download_file(url, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)

    if dest.exists():
        log("DEBUG", f"Already exists: {dest.name}")
        return

    log("INFO", f"Downloading {dest.name}")

    with urllib.request.urlopen(url) as r:
        length = r.headers.get("Content-Length")
        total = int(length) if length and length.isdigit() else None

        with Progress(
            TextColumn("{task.description}"),
            BarColumn(),
            DownloadColumn(),
            TimeRemainingColumn(),
            transient=True,
        ) as prog:

            task = prog.add_task(dest.name, total=total)

            with open(dest, "wb") as f:
                while True:
                    chunk = r.read(1024 * 128)
                    if not chunk:
                        break
                    f.write(chunk)
                    prog.update(task, advance=len(chunk))

    #log("SUCCESS", f"Downloaded {dest.name}")
    log("SUCCESS", f"Download complete.")

def prompt_reboot():
    try:
        ans = console.input("\nReboot now to apply changes? ([Y]/n): ").strip().lower()
    except Exception:
        ans = "n"

    if ans in ("", "y", "yes"):
        log("INFO", "Rebooting system now")
        subprocess.run(["shutdown", "/r", "/t", "0"], check=False)
    else:
        log("WARN", "Reboot skipped. Please reboot manually to apply language changes.")
        pause_exit(0)

# ============================================================
# Main
# ============================================================

def main():
    console.clear()
    panel("Windows Language Pack Deployment", title="deployLangs", color="yellow")

    panel(
        f"Base dir: {EXE_DIR}\n"
        f"Inputs:\n"
        f"  - {INPUT_FILES[0].name}\n"
        f"  - {INPUT_FILES[1].name}\n"
        f"  - {INPUT_FILES[2].name}\n"
        f"Downloads: {DOWNLOAD_ROOT}",
        title="Paths",
        color="cyan",
    )

    check_internet()

    log("INFO", "Loading input URL lists")
    lines = load_all_lines(INPUT_FILES)

    arch_path, arch_token = detect_arch()
    log("INFO", f"Detected architecture: {arch_path}")

    time.sleep(1.2)
    console.clear()
    console.print("\n")

    console.print(
"""  _                                        __  __    _                                      ___            
 | |   __ _ _ _  __ _ _  _ __ _ __ _ ___  |  \\/  |__| |   __ _ _ _  __ _ _  _ __ _ __ _ ___| __|_ _ __ ___ 
 | |__/ _` | ' \\/ _` | || / _` / _` / -_) | |\\/| / _| |__/ _` | ' \\/ _` | || / _` / _` / -_) _/ _` / _/ -_)
 |____\\__,_|_||_\\__, |\\_,_\\__,_\\__, \\___| |_|  |_\\__|____\\__,_|_||_\\__, |\\_,_\\__,_\\__, \\___|_|\\__,_\\__\\___|
                |___/          |___/                               |___/          |___/                    """,
        style="yellow",
    )

    console.rule(" iam5 ")
    console.print("\n")

    langs = extract_languages(lines)
    if not langs:
        fatal("No languages found")

    lang = select_language(langs)

    panel(
        f"Language: {lang}\nArchitecture: {arch_path}\nMode: {'DRY-RUN' if DRY_RUN else 'INSTALL'}",
        title="Selection Summary",
        color="green",
    )

    try:
        preserve = console.input(
            "\nPreserve existing keyboard layouts? ([Y]/n): "
        ).strip().lower() not in ("n", "no")

        extend = console.input(
            "Add language instead of replacing existing ones? ([Y]/n): "
        ).strip().lower() not in ("n", "no")
    except Exception:
        fatal("Invalid input")

    log(
        "INFO",
        f"Language list mode: {'EXTEND' if extend else 'REPLACE'}, "
        f"Keyboards preserved: {'YES' if preserve else 'NO'}",
    )

    panel(
        f"Language: {lang}\n"
        f"Architecture: {arch_path}\n"
        f"Mode: {'DRY-RUN' if DRY_RUN else 'INSTALL'}\n"
        f"Language list: {'Extend' if extend else 'Replace'}\n"
        f"Preserve keyboards: {'Yes' if preserve else 'No'}",
        title="Selection Summary",
        color="green",
    )

    found = []

    for line in lines:
        u = unquote(line)
        l = u.lower()

        if is_winpe(l):
            continue

        if f"/localexperiencepack/{lang}/" in l:
            if l.endswith(".appx") or l.endswith("license.xml"):
                found.append(line)
            continue

        if f"microsoft-windows-client-language-pack_{arch_path}_{lang}.cab" in l:
            found.append(line)
            continue

        if (
            "microsoft-windows-languagefeatures-" in l
            and f"-{lang}-package" in l
            and f"~{arch_token}~~.cab" in l
        ):
            for k in FOD_FEATURE_KEYWORDS:
                if k in l:
                    found.append(line)
                    break

    if not found:
        fatal("No matching files found")

    file_list = "\n".join(f"• {Path(unquote(f)).name}" for f in found)
    panel(file_list, title="Files to download", color="cyan")

    appx_files = []
    cab_files = []

    for url in found:
        name = Path(unquote(url)).name
        if name.lower().endswith(".appx") or name.lower() == "license.xml":
            dest = APPX_DIR / name
            download_file(url, dest)
            appx_files.append(dest)
        else:
            dest = CAB_DIR / name
            download_file(url, dest)
            cab_files.append(dest)

    if DRY_RUN:
        log("WARN", "Dry-run enabled: installation skipped")
        pause_exit(0)

    panel("Installing CAB packages", color="yellow")

    for i, cab in enumerate(cab_files, 1):
        log("INFO", f"Installing CAB {i}/{len(cab_files)}: {cab.name}")
        run([
            "dism",
            "/Online",
            "/Add-Package",
            f"/PackagePath:{cab}",
            "/NoRestart",
        ])
        log("SUCCESS", f"Installed {cab.name}")

    lep = next((x for x in appx_files if "languageexperiencepack" in x.name.lower()), None)
    lic = next((x for x in appx_files if x.name.lower() == "license.xml"), None)

    if lep:
        log("INFO", "Installing Language Experience Pack")
        if lic:
            run([
                "powershell",
                "-NoProfile",
                "-Command",
                f"Add-AppxProvisionedPackage -Online -PackagePath '{lep}' -LicensePath '{lic}'"
            ])
        else:
            run([
                "powershell",
                "-NoProfile",
                "-Command",
                f"Add-AppxProvisionedPackage -Online -PackagePath '{lep}' -SkipLicense"
            ])
    else:
        log("WARN", "No Language Experience Pack found")

    # ------------------------------------------------------------
    # Apply language (this is REQUIRED for actual UI change)
    # ------------------------------------------------------------

    log("INFO", "Applying language settings to current user and system")

    if extend:
        if preserve:
            ps_cmd = (
                f"$list = Get-WinUserLanguageList; "
                f"if ($list.LanguageTag -notcontains '{lang}') {{ "
                f"$list.Add('{lang}') }}; "
                f"Set-WinUserLanguageList $list -Force; "
                f"Set-WinUILanguageOverride -Language {lang}; "
                f"Set-WinSystemLocale -SystemLocale {lang}"
            )
        else:
            ps_cmd = (
                f"$list = Get-WinUserLanguageList | "
                f"Where-Object {{ $_.LanguageTag -ne '{lang}' }}; "
                f"$new = New-WinUserLanguageList '{lang}'; "
                f"$new.AddRange($list); "
                f"Set-WinUserLanguageList $new -Force; "
                f"Set-WinUILanguageOverride -Language {lang}; "
                f"Set-WinSystemLocale -SystemLocale {lang}"
            )
    else:
        ps_cmd = (
            f"Set-WinUILanguageOverride -Language {lang}; "
            f"$list = New-WinUserLanguageList '{lang}'; "
            f"Set-WinUserLanguageList $list -Force; "
            f"Set-WinSystemLocale -SystemLocale {lang}"
        )

    run([
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-Command", ps_cmd,
    ])

    log("SUCCESS", f"Language '{lang}' set as system UI language")

    # ------------------------------------------------------------
    # Apply language to system / login screen (WELCOME / LOGON UI)
    # ------------------------------------------------------------

    log("INFO", "Applying language to system login / welcome screen")

    ps_cmd_login = f"""
# Apply language to system (login screen / welcome screen) via registry
$lang = '{lang}'

# Set system locale and UI language override
Set-WinSystemLocale -SystemLocale $lang
Set-WinUILanguageOverride -Language $lang

# Write to MUI settings registry
$dst = 'HKLM:\\SYSTEM\\CurrentControlSet\\Control\\MUI\\Settings'
try {{
    New-ItemProperty -Path $dst -Name 'PreferredUILanguages' -Value @($lang) -PropertyType MultiString -Force | Out-Null
    New-ItemProperty -Path $dst -Name 'PreferredUILanguagesPending' -Value @($lang) -PropertyType MultiString -Force | Out-Null
}} catch {{
    Write-Warning "Could not write MUI settings: $_"
}}

# Copy to Welcome Screen and new user accounts
try {{
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUserAccounts $true
    Write-Host "Copied language settings to Welcome Screen and new user accounts"
}} catch {{
    Write-Warning "Copy-UserInternationalSettingsToSystem not available, falling back to intl.cpl"
    $xml = @"
<gs:GlobalizationServices xmlns:gs="urn:longhornGlobalizationUnattend">
  <gs:UserList>
    <gs:User UserID="Current"/>
    <gs:User UserID="System"/>
  </gs:UserList>
  <gs:UILanguagePreferences>
    <gs:UILanguage Value="{lang}"/>
  </gs:UILanguagePreferences>
</gs:GlobalizationServices>
"@
    $$path = "$$env:TEMP\\intl.xml"
    $$xml | Out-File -Encoding UTF8 $$path
    control.exe "intl.cpl,,/f:$$path"
}}
"""


    run([
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-Command", ps_cmd_login,
    ])
    console.print("\n")
    console.rule(" FINISHED ")
    console.print("\n")
    panel(
        "Installation complete.\n\n"
        "Note:\n"
        "• DISM may report 'RestartNeeded : False' for individual packages.\n"
        "• This is normal and does NOT indicate completion of language activation.\n\n"
        "A reboot IS REQUIRED to activate the new display language.",
        title="Success",
        color="green",
    )

    prompt_reboot()


if __name__ == "__main__":
    main()
