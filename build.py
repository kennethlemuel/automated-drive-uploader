#!/usr/bin/env python3
import os
import platform
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


ROOT = Path(__file__).resolve().parent
APP_NAME = "Efficient Audit Uploader"
ICON = ROOT / "assets" / ("ESLogo.ico" if sys.platform == "win32" else "ESLogo.icns")
ARCH = "arm64" if platform.machine().lower() in {"arm64", "aarch64"} else "x64"
PLATFORM = "Windows" if sys.platform == "win32" else "macOS"
PACKAGE = ROOT / "dist" / f"Efficient-Audit-Uploader-{PLATFORM}-{ARCH}.zip"


with tempfile.TemporaryDirectory(prefix="efficient-audit-uploader-") as temporary:
    temporary = Path(temporary)
    subprocess.run(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--windowed",
            *(["--onefile"] if sys.platform == "win32" else []),
            "--name",
            APP_NAME,
            "--icon",
            str(ICON),
            "--add-data",
            f"{ROOT / 'web' / 'index.html'}{os.pathsep}web",
            "--specpath",
            str(temporary),
            "--workpath",
            str(temporary / "build"),
            "--distpath",
            str(temporary / "dist"),
            str(ROOT / "desktop_app.py"),
        ],
        cwd=ROOT,
        check=True,
    )

    PACKAGE.parent.mkdir(exist_ok=True)
    if sys.platform == "win32":
        executable = temporary / "dist" / f"{APP_NAME}.exe"
        with zipfile.ZipFile(PACKAGE, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.write(executable, executable.name)
    else:
        app = temporary / "dist" / f"{APP_NAME}.app"
        subprocess.run(["codesign", "--verify", "--deep", "--strict", str(app)], check=True)
        subprocess.run(["ditto", "-c", "-k", "--sequesterRsrc", "--keepParent", str(app), str(PACKAGE)], check=True)

print(PACKAGE)
