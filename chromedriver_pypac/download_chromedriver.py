#!/usr/bin/env python3
"""Download ChromeDriver for a specific Chrome version.

Same behaviour as the ``chromedriver`` util, but every HTTP request goes
through a :class:`pypac.PACSession` so that a system/enterprise PAC file
(Proxy Auto-Config) is honoured. When no PAC file is discovered pypac falls
back to requests' normal proxy handling (``HTTP_PROXY`` / ``HTTPS_PROXY``).
"""

import platform
import sys
import zipfile
from pathlib import Path

from pypac import PACSession

# Chrome for Testing JSON endpoints (Chrome >= 115)
CFT_KNOWN_GOOD = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"
CFT_LATEST_PATCH = "https://googlechromelabs.github.io/chrome-for-testing/latest-patch-versions-per-build-with-downloads.json"

# Legacy endpoint (Chrome < 115)
LEGACY_URL = "https://chromedriver.storage.googleapis.com"

# One PAC-aware session reused for every request in this run.
_SESSION = PACSession()


def detect_platform() -> str:
    system = platform.system().lower()
    arch = platform.machine().lower()
    if system == "darwin":
        return "mac-arm64" if arch == "arm64" else "mac-x64"
    elif system == "windows":
        return "win64" if "64" in arch else "win32"
    return "linux64"


def download_chromedriver(chrome_version: str, output_dir: str = ".") -> Path:
    major = int(chrome_version.split(".")[0])
    plat = detect_platform()
    output = Path(output_dir)
    output.mkdir(parents=True, exist_ok=True)

    if major >= 115:
        url = _get_cft_url(chrome_version, plat)
    else:
        url = _get_legacy_url(chrome_version, plat)

    print(f"Downloading: {url}")
    zip_path = output / "chromedriver.zip"
    _download_file(url, zip_path)

    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            for member in zf.infolist():
                # Strip top-level folder (e.g. chromedriver-win64/) so files
                # land directly in output_dir. Skip path traversal attempts.
                parts = Path(member.filename).parts
                rel = Path(*parts[1:]) if len(parts) > 1 else Path(parts[0])
                if ".." in rel.parts:
                    continue
                dest = output / rel
                if member.is_dir():
                    dest.mkdir(parents=True, exist_ok=True)
                else:
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    dest.write_bytes(zf.read(member.filename))
    finally:
        zip_path.unlink(missing_ok=True)

    print(f"ChromeDriver extracted to: {output}")
    return output


def _download_file(url: str, dest: Path) -> None:
    with _SESSION.get(url, stream=True) as resp:
        resp.raise_for_status()
        with open(dest, "wb") as fh:
            for chunk in resp.iter_content(chunk_size=65536):
                fh.write(chunk)


def _get_json(url: str) -> dict:
    resp = _SESSION.get(url)
    resp.raise_for_status()
    return resp.json()


def _get_cft_url(version: str, plat: str) -> str:
    parts = version.split(".")

    # Try exact match via latest-patch-versions-per-build
    build_key = ".".join(parts[:3])  # e.g. 120.0.6099
    print(f"Looking up build: {build_key}")

    data = _get_json(CFT_LATEST_PATCH)
    if build_key in data["builds"]:
        entry = data["builds"][build_key]
        for dl in entry["downloads"].get("chromedriver", []):
            if dl["platform"] == plat:
                return dl["url"]

    # Fallback: find closest match by major version
    print(f"Exact build not found, searching by major version {parts[0]}...")
    data = _get_json(CFT_KNOWN_GOOD)
    matches = [
        v for v in data["versions"]
        if v["version"].startswith(parts[0] + ".")
        and "chromedriver" in v.get("downloads", {})
    ]
    if not matches:
        raise RuntimeError(f"No ChromeDriver found for Chrome {version}")

    best = matches[-1]  # latest patch for that major
    print(f"Using version: {best['version']}")
    for dl in best["downloads"]["chromedriver"]:
        if dl["platform"] == plat:
            return dl["url"]

    raise RuntimeError(f"No download for platform {plat}")


def _get_legacy_url(version: str, plat: str) -> str:
    build = ".".join(version.split(".")[:3])
    ver_url = f"{LEGACY_URL}/LATEST_RELEASE_{build}"
    resp = _SESSION.get(ver_url)
    resp.raise_for_status()
    driver_version = resp.text.strip()
    print(f"Matched driver version: {driver_version}")

    plat_map = {"linux64": "linux64", "mac-x64": "mac64", "mac-arm64": "mac_arm64",
                "win32": "win32", "win64": "win32"}
    return f"{LEGACY_URL}/{driver_version}/chromedriver_{plat_map[plat]}.zip"


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python download_chromedriver.py <chrome_version> [output_dir]")
        print("Example: python download_chromedriver.py 120.0.6099.109")
        sys.exit(1)

    ver = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "."
    download_chromedriver(ver, out)
