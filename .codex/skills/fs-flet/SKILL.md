---
name: fs-flet
description: Diagnose and fix Flet desktop applications packaged with PyInstaller, especially missing window icons, unavailable assets, incorrect writable-data paths, and content hidden by the desktop title bar. Use when reviewing a Flet main.py, .spec file, or packaged Windows executable.
---

# Flet PyInstaller 排查

## 工作流

1. Read the application entry point and the PyInstaller `.spec` together. Enumerate every `ft.Image(src=...)`, `page.window.icon`, `assets_dir`, and file-system write.
2. Separate paths into two roots:
   - Bundle/read-only root: `Path(sys._MEIPASS)` when frozen, otherwise the project directory. Use it for `assets` and the window icon.
   - Application/data root: `Path(sys.executable).resolve().parent` when frozen, otherwise the source directory. Use it for databases, uploads, temporary files, and user exports.
3. Centralize asset lookup. Do not derive bundled assets from `sys.executable` in one-file mode; that points beside the executable while PyInstaller extracts bundled data under `_MEIPASS`.
4. Keep Flet relative image sources and `ft.run(assets_dir=...)` on the same bundle asset directory. Ensure the spec contains `datas=[('assets', 'assets')]` (or an equivalent validated collection).
5. Configure the native title bar explicitly (`page.window.title_bar_hidden = False`) and wrap the root layout in `ft.SafeArea`, with a small `minimum_padding` if the first row is still visually clipped.
6. Rebuild, launch the actual executable, and verify: title-bar icon, navigation SVGs/logo, first content row, database creation beside the executable, and attachment read/write behavior.

## Recommended code pattern

```python
@staticmethod
def get_bundle_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parent

@staticmethod
def get_asset_path(filename: str) -> Path:
    return Services.get_bundle_dir() / "assets" / filename
```

Keep `get_app_dir()` for writable data and pass `str(Services.get_asset_path("").resolve())` to `ft.run`. For the root page, use `ft.SafeArea(content=root, minimum_padding=ft.Padding(0, 8, 0, 0))` only when needed; do not hide the title bar to compensate for a layout offset.

## 常见错误

- Using `Path(sys.executable).parent / "assets"` for bundled assets in a one-file build.
- Assuming the `icon=` argument in `EXE` also supplies the runtime Flet window icon. It only sets the Windows executable icon; `page.window.icon` is a separate runtime path.
- Using absolute source-tree paths in `ft.Image`; use relative names with the configured `assets_dir`.
- Moving the database into `_MEIPASS`; that directory is temporary and not writable/persistent.
