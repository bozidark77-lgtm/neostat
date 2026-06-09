# assets/

Optional branding for the app window and the packaged builds. **All of these are
optional** — the spec only bundles files that exist and `app.py` loads the icon
inside a `try/except`, so the build is green with this folder empty.

| File                | Used by            | Purpose                                  |
|---------------------|--------------------|------------------------------------------|
| `neostat.ico`       | Windows build      | `.exe` + window icon                     |
| `neostat.icns`      | macOS build        | `.app` bundle icon                       |
| `neostat_logo.png`  | runtime (all OSes) | `iconphoto` window logo                  |

Drop the files here with exactly these names and the next build picks them up.
The original `.ico` / `.png` can be pulled from the full PyInstaller extraction
of `Neostat_Analiza.exe` (see `SUMMARY.md`). To make a macOS `.icns` from a PNG:

```bash
# 1024×1024 source.png -> neostat.icns
mkdir neostat.iconset
for s in 16 32 64 128 256 512; do
  sips -z $s   $s   source.png --out neostat.iconset/icon_${s}x${s}.png
  sips -z $((s*2)) $((s*2)) source.png --out neostat.iconset/icon_${s}x${s}@2x.png
done
iconutil -c icns neostat.iconset -o neostat.icns
```
