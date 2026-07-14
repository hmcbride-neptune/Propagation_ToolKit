# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['Main_GUI.py'],
    pathex=[],
    binaries=[],
    datas=[('Logos', 'Logos')],
    hiddenimports=['format_files', 'terrain', 'demographics', 'switch_group', 'switch_dhb', 'move_layer', 'show_hide', 'rename_finds', 'export_transmitters', 'measurement_analysis', 'presentation', 'screenshots', 'assumptions', 'add_project_files', 'faa_tool', 'zipper', 'new_request', 'Data_Import', 'move_data', 'settings', 'edx_open', 'openpyxl', 'webdriver_manager.microsoft'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='Propagation_debug',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='Propagation_debug',
)
