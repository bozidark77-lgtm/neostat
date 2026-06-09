# Neostat_Analiza.spec — reproducible, cross-platform PyInstaller build
# Build:  pyinstaller --noconfirm --clean Neostat_Analiza.spec
# Output: dist/Neostat_Analiza.exe          on Windows  (single-file, windowed)
#         dist/Neostat_Analiza.app          on macOS    (windowed .app bundle)
#         dist/Neostat_Analiza              on Linux     (single-file ELF binary)
#
# NOTE: PyInstaller bundles the native interpreter of the OS it runs on. It does
#       NOT cross-compile — run this on Windows to get a .exe, on macOS to get a
#       .app. The CI workflow builds each target on its matching runner.

import os
import sys

_is_mac = sys.platform == "darwin"
_is_win = sys.platform == "win32"

# Per-platform application icon. Windows uses .ico, macOS uses .icns. Each is
# optional — if the file is absent the build still succeeds with a default icon.
_ico = "assets/neostat.ico"
_icns = "assets/neostat.icns"
if _is_win and os.path.exists(_ico):
    _icon = _ico
elif _is_mac and os.path.exists(_icns):
    _icon = _icns
else:
    _icon = None

# Bundle any icon/logo assets that exist, so they are reachable at runtime via
# sys._MEIPASS (see ConvertLauncherApp._resource_path in src/app.py).
_datas = []
for _f in (_ico, _icns, "assets/neostat_logo.png"):
    if os.path.exists(_f):
        _datas.append((_f, "."))   # extracted to sys._MEIPASS root at runtime

a = Analysis(
    ["src/app.py"],
    pathex=["src"],                # so `import analyze` / `import convert` resolve
    binaries=[],
    datas=_datas,
    hiddenimports=[
        "openpyxl",                # pandas imports the excel engine lazily by string
        "openpyxl.cell._writer",
    ],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="Neostat_Analiza",
    debug=False,
    strip=False,
    upx=False,
    console=False,                 # windowed GUI app (no console window)
    icon=_icon,
)

# On macOS, wrap the single-file executable in a double-clickable .app bundle.
if _is_mac:
    app = BUNDLE(
        exe,
        name="Neostat_Analiza.app",
        icon=_icon,
        bundle_identifier="me.slobodan.neostat.analiza",
        info_plist={
            "CFBundleName": "Neostat Analiza",
            "CFBundleDisplayName": "AMS – Analiza izveštaja",
            "NSHighResolutionCapable": True,
        },
    )
