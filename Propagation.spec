# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller build spec for the Propagation toolkit.
#
# Produces a windowed one-FOLDER app named "Propagation" (no console window).
# The runnable exe is dist/Propagation/Propagation.exe.
#
# One-folder (not one-file) is used on purpose:
#   * Main_GUI.py relaunches this exe as "Propagation.exe --run <module>" for
#     every sub-tool. One-file would re-extract the whole ~185 MB bundle to a
#     temp dir on each launch; one-folder starts instantly.
#   * One-file PyInstaller exes are frequently false-flagged & quarantined by
#     Windows Defender; the one-folder layout is far less likely to be.
#
# Every sub-tool launched by string name must be listed in `hiddenimports`
# below -- PyInstaller cannot discover them because they are never imported
# directly in source.
#
# Build with:   pyinstaller Propagation.spec
# Output:       dist/Propagation/Propagation.exe

from PyInstaller.utils.hooks import collect_submodules

# Sub-tool modules launched via `Main_GUI.script_cmd()` + runpy.run_module.
sub_tools = [
    "format_files",
    "terrain",
    "demographics",
    "switch_group",
    "switch_dhb",
    "move_layer",
    "show_hide",
    "rename_finds",
    "export_transmitters",
    "measurement_analysis",
    "presentation",
    "screenshots",
    "assumptions",
    "add_project_files",
    "faa_tool",
    "zipper",
    # directly imported helpers, listed for safety
    "new_request",
    "Data_Import",
    "move_data",
    "settings",
    "edx_open",
]

# Packages with lazy / dynamic imports that PyInstaller can miss.
extra_hidden = [
    "openpyxl",                         # lazy import in Data_Import.py
    "webdriver_manager.microsoft",      # lazy import in faa_tool.py
    "pythoncom",
    "pywintypes",
    "win32timezone",
]

hiddenimports = sub_tools + extra_hidden
hiddenimports += collect_submodules("pywinauto")

a = Analysis(
    ["Main_GUI.py"],
    pathex=[],
    binaries=[],
    datas=[
        ("Logos", "Logos"),                      # bundle icons next to the exe
        ("drivers\\msedgedriver.exe", "."),       # Edge WebDriver for faa_tool.py
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,             # one-folder: binaries collected below
    name="Propagation",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,                     # windowed: no black console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="Logos\\Propagation.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Propagation",
)
