# -*- mode: python ; coding: utf-8 -*-
# Build:  pyinstaller --noconfirm AgroApp.spec
# PyInstaller only builds for the system it runs on:
#   on Windows -> dist/AgroApp.exe        on macOS -> dist/AgroApp.app
import sys
from PyInstaller.utils.hooks import collect_data_files

datas = [
    # Starting chart colors; copied to the user's app-data folder on first run (see Utils.py)
    ('series_colors.json', '.'),
]
# CustomTkinter's themes (green.json etc.) and fonts aren't Python files, so include them explicitly
datas += collect_data_files('customtkinter')

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Not used by the app (only optional plotting in statsmodels). Leaving it out keeps the app smaller
    # and avoids matplotlib 3.9.1's broken Windows DLL, which made PyInstaller fail on Windows.
    excludes=['matplotlib'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

if sys.platform == "darwin":
    # macOS: a normal double-clickable app, dist/AgroApp.app (a bare executable would open in Terminal)
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name='AgroApp',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
    coll = COLLECT(exe, a.binaries, a.datas, strip=False, upx=True, name='AgroApp')
    app = BUNDLE(
        coll,
        name='AgroApp.app',
        bundle_identifier='com.agrunomproject.agroapp',
        info_plist={'NSHighResolutionCapable': True},
    )
else:
    # Windows: a single file, dist/AgroApp.exe
    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        name='AgroApp',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        upx_exclude=[],
        runtime_tmpdir=None,
        console=False,  # no black console window next to the app
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
    )
