"""Сборка портативного CRF.exe в корень проекта."""

import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ICONS = ROOT / "icons"

SIZE_PNG_RE = re.compile(r'^(\d+)[xX](\d+)\.png$', re.IGNORECASE)
CRITICAL_SIZES = {16, 32, 48, 256}
RECOMMENDED_SIZES = {16, 24, 32, 48, 64, 128, 256}


def _ensure_pillow():
    try:
        from PIL import Image
        return Image
    except ImportError:
        print("Для конвертации PNG -> ICO нужен Pillow.\n"
              "Установите: python -m pip install Pillow")
        sys.exit(1)


def collect_size_pngs() -> list[tuple[int, Path]]:
    found = []
    if not ICONS.exists():
        return found
    for path in ICONS.iterdir():
        if not path.is_file():
            continue
        m = SIZE_PNG_RE.match(path.name)
        if m:
            w, h = int(m.group(1)), int(m.group(2))
            if w != h:
                print(f"  [warn] {path.name}: стороны разные ({w}x{h}), пропускаю")
                continue
            found.append((w, path))
    found.sort(key=lambda x: x[0], reverse=True)
    return found


def _validate_sizes(present: set[int]) -> None:
    missing_critical = CRITICAL_SIZES - present
    missing_extra = RECOMMENDED_SIZES - present - CRITICAL_SIZES
    if missing_critical:
        print(f"  [warn] Нет критических размеров: "
              f"{', '.join(f'{s}x{s}' for s in sorted(missing_critical))}")
    if missing_extra:
        print(f"  [info] Нет опциональных размеров: "
              f"{', '.join(f'{s}x{s}' for s in sorted(missing_extra))}")


def build_app_ico() -> None:
    ico_path = ICONS / "app.ico"

    if ico_path.exists():
        print(f"  [ok]   {ico_path.name} найден - использую как есть")
        return

    pngs = collect_size_pngs()
    if not pngs:
        print("  [skip] Нет ни app.ico, ни PNG с разрешениями - exe без иконки")
        return

    Image = _ensure_pillow()

    images = []
    for size, path in pngs:
        img = Image.open(path).convert("RGBA")

        if img.size != (size, size):
            img = img.resize((size, size), Image.Resampling.LANCZOS)

        images.append(img)
        print(f"  [img]  {path.name}  ({size}x{size})")

    _validate_sizes({img.width for img in images})

    base = images[0]  # обычно 256x256, т.к. collect_size_pngs сортирует по убыванию

    base.save(
        ico_path,
        format="ICO",
        append_images=images[1:],
        sizes=[(img.width, img.height) for img in images],
    )

    print(
        f"  [ok]   {ico_path.name}  <- {len(images)} размер(ов): "
        f"{', '.join(f'{img.width}x{img.height}' for img in images)}"
    )


def run_pyinstaller() -> None:
    print("\nPyInstaller:")
    cmd = [sys.executable, "-m", "PyInstaller", "--clean", "build.spec"]
    result = subprocess.run(cmd, cwd=ROOT)
    if result.returncode != 0:
        print(f"\nPyInstaller завершился с кодом {result.returncode}")
        sys.exit(result.returncode)


def move_exe_to_root() -> Path:
    exe_name = "CRF.exe" if os.name == "nt" else "CRF"
    src = ROOT / "dist" / exe_name
    dst = ROOT / exe_name
    if not src.exists():
        print(f"Не найден {src}")
        sys.exit(1)
    if dst.exists():
        dst.unlink()
    shutil.move(str(src), str(dst))
    try:
        shutil.rmtree(ROOT / "dist")
    except OSError:
        pass
    print(f"\nГотово: {dst}")
    return dst


def main() -> None:
    print("Иконки:")
    build_app_ico()
    run_pyinstaller()
    move_exe_to_root()


if __name__ == "__main__":
    main()
