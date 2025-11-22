# Windows Test Mode + Driver Signature Enforcement Controller

This tool enables or disables Windows Test Mode and Driver Signature Enforcement (DSE) using BCDEdit, with optional automation flags, WOW64-aware System32 execution, and full administrator elevation support.
A PyInstaller build can be generated with automatic UAC elevation.

---

## Features

- Enable or disable:
  - testsigning
  - nointegritychecks
- Architecture-aware:
  - Disables WOW64 filesystem redirection only on 64-bit Windows
- Automatic elevation handling
- Rich-formatted console output
- Logging to etka-post-setup.log
- Optional:
  - Skip prompts (--auto-accept)
  - Auto-reboot (--auto-reboot)
- Clean PyInstaller build support using --uac-admin

---

## Usage

script.exe                            (default: enable test mode + disable DSE)
script.exe --mode enable
script.exe --mode disable
script.exe --auto-accept
script.exe --auto-reboot

---

## Modes

enable  → testsigning=on,  nointegritychecks=on

disable → testsigning=off, nointegritychecks=off

---

## Optional Flags

Skips start and exit prompts.
--auto-accept

Automatically reboots after changing BCD state.
--auto-reboot

---

## PyInstaller Build

To build a standalone EXE that automatically runs with administrator privileges:

pyinstaller --onefile --icon=configuration.ico --clean --uac-admin winTestMode+DSE.py

If using a .spec file, set:

uac_admin=True

inside the EXE() section.

---

## Notes

- This tool modifies the Windows boot configuration (BCD).
- A reboot is required after changing Test Mode or DSE settings.
- When --auto-reboot is used, the system will reboot automatically using shutdown.exe /r /t 5 /f.
