from __future__ import annotations

import argparse
import contextlib
import os
import re
import shutil
import socket
import ssl
import subprocess
import sys
import time
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.request import urlopen


TOOL_DIR = Path(__file__).resolve().parent
TEMPLATE_DIR = TOOL_DIR / "template_htmlinppt"
OUTPUT_DIR = TOOL_DIR / "outputs"
WEB_DIR_NAME = "pd_soi_fbe_v3_local_package"
MANIFEST_NAME = "pd_soi_fbe_powerpoint_manifest.xml"
OLD_PPT_NAME = "PD_SOI_FBE_V3_Direct_In_PowerPoint.pptx"
FALLBACK_NAME = "soi_dynamic_webpage.png"
PORT = 8765


def log(message: str) -> None:
    print(message, flush=True)


def safe_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "_", name).strip(" ._")
    cleaned = re.sub(r"\\s+", "_", cleaned)
    return cleaned[:48] or "html_page"


def copytree_clean(src: Path, dst: Path) -> None:
    if dst.exists():
        shutil.rmtree(dst)
    shutil.copytree(src, dst)


def read_text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="ignore")


def write_text(path: Path, text: str) -> None:
    path.write_text(text, encoding="utf-8", newline="")


def choose_python() -> str | None:
    candidates = [
        ["python", "-c", "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)"],
        ["python3", "-c", "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)"],
        ["py", "-3", "-c", "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)"],
    ]
    for cmd in candidates:
        try:
            result = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except FileNotFoundError:
            continue
        if result.returncode == 0:
            return cmd[0] if cmd[0] != "py" else "py -3"
    return None


def python_command_for_server() -> list[str]:
    candidates = [
        ["python", "-c", "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)"],
        ["python3", "-c", "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)"],
        ["py", "-3", "-c", "import sys; raise SystemExit(0 if sys.version_info[0] >= 3 else 1)"],
    ]
    for cmd in candidates:
        try:
            result = subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except FileNotFoundError:
            continue
        if result.returncode == 0:
            return cmd[:1] if cmd[0] != "py" else ["py", "-3"]
    raise RuntimeError("Python 3 could not start.")


def find_browser() -> Path | None:
    candidates = [
        Path(os.environ.get("ProgramFiles(x86)", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("ProgramFiles", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("LocalAppData", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("ProgramFiles", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("ProgramFiles(x86)", "")) / "Google/Chrome/Application/chrome.exe",
        Path(os.environ.get("LocalAppData", "")) / "Google/Chrome/Application/chrome.exe",
    ]
    for path in candidates:
        if path and path.exists():
            return path
    for name in ("msedge.exe", "chrome.exe"):
        found = shutil.which(name)
        if found:
            return Path(found)
    return None


def stop_python_on_port(port: int) -> None:
    ps = (
        "$ownerPids = Get-NetTCPConnection -LocalAddress 127.0.0.1 "
        f"-LocalPort {port} -State Listen -ErrorAction SilentlyContinue | "
        "Select-Object -ExpandProperty OwningProcess -Unique; "
        "foreach($ownerPid in $ownerPids){ "
        "try { $proc = Get-Process -Id $ownerPid -ErrorAction Stop; "
        "if($proc.ProcessName -match 'python|py'){ Stop-Process -Id $ownerPid -Force } } catch {} }"
    )
    with contextlib.suppress(Exception):
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=10)


def wait_for_url(url: str, timeout: float = 10.0) -> bool:
    deadline = time.time() + timeout
    context = ssl._create_unverified_context()
    while time.time() < deadline:
        try:
            with urlopen(url, context=context, timeout=2) as response:
                return 200 <= response.status < 400
        except Exception:
            time.sleep(0.25)
    return False


def make_screenshot(package_dir: Path, web_dir: Path, screenshot_path: Path, url_token: str) -> bool:
    browser = find_browser()
    if not browser:
        log("No Edge/Chrome found. Skip fallback screenshot.")
        return False

    stop_python_on_port(PORT)
    cmd = python_command_for_server() + ["server_https.py"]
    creationflags = subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
    server = subprocess.Popen(
        cmd,
        cwd=str(web_dir),
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=creationflags,
    )
    try:
        url = f"https://127.0.0.1:{PORT}/index.html?v={url_token}&shot={int(time.time())}"
        if not wait_for_url(url):
            log("Local server did not respond in time. Skip fallback screenshot.")
            return False
        args = [
            str(browser),
            "--headless",
            "--disable-gpu",
            "--hide-scrollbars",
            "--ignore-certificate-errors",
            "--allow-insecure-localhost",
            "--window-size=1680,945",
            f"--screenshot={screenshot_path}",
            url,
        ]
        result = subprocess.run(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=30)
        return result.returncode == 0 and screenshot_path.exists() and screenshot_path.stat().st_size > 1000
    except Exception as exc:
        log(f"Screenshot failed: {exc}")
        return False
    finally:
        with contextlib.suppress(Exception):
            server.terminate()
            server.wait(timeout=3)
        with contextlib.suppress(Exception):
            server.kill()
        stop_python_on_port(PORT)


def zip_replace_text(zip_path: Path, inner_name: str, new_text: str) -> None:
    temp_zip = zip_path.with_suffix(zip_path.suffix + ".tmp")
    with zipfile.ZipFile(zip_path, "r") as src, zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as dst:
        replaced = False
        for item in src.infolist():
            if item.filename == inner_name:
                dst.writestr(item, new_text.encode("utf-8"))
                replaced = True
            else:
                dst.writestr(item, src.read(item.filename))
        if not replaced:
            raise RuntimeError(f"{inner_name} not found in {zip_path}")
    temp_zip.replace(zip_path)


def zip_replace_file(zip_path: Path, inner_name: str, file_path: Path) -> None:
    temp_zip = zip_path.with_suffix(zip_path.suffix + ".tmp")
    with zipfile.ZipFile(zip_path, "r") as src, zipfile.ZipFile(temp_zip, "w", zipfile.ZIP_DEFLATED) as dst:
        replaced = False
        for item in src.infolist():
            if item.filename == inner_name:
                dst.writestr(item, file_path.read_bytes())
                replaced = True
            else:
                dst.writestr(item, src.read(item.filename))
        if not replaced:
            raise RuntimeError(f"{inner_name} not found in {zip_path}")
    temp_zip.replace(zip_path)


def update_manifest(manifest: Path, addin_id: str, version: str, url_token: str) -> None:
    text = read_text(manifest)
    text = re.sub(r"<Id>[^<]+</Id>", f"<Id>{addin_id}</Id>", text, count=1)
    text = re.sub(r"<Version>[^<]+</Version>", f"<Version>{version}</Version>", text, count=1)
    source = f'<SourceLocation DefaultValue="https://127.0.0.1:{PORT}/index.html?v={url_token}" />'
    text = re.sub(r'<SourceLocation\s+DefaultValue="[^"]+"\s*/>', source, text, count=1)
    write_text(manifest, text)


def write_setup_script(web_dir: Path) -> None:
    script = r"""$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$cert = Join-Path $scriptDir 'localhost_selfsigned.crt'

if (-not (Test-Path -LiteralPath $cert)) { throw "Certificate not found: $cert" }

Write-Host 'Trusting localhost certificate for current Windows user...'
& certutil -user -addstore Root $cert | Out-Host

$manifests = @(Get-ChildItem -LiteralPath $scriptDir -Filter '*.xml' | Where-Object {
  $_.Name -like '*manifest*.xml'
})

if ($manifests.Count -eq 0) { throw "No manifest XML files found in: $scriptDir" }

Write-Host 'Registering PowerPoint content add-in manifests for current user...'
foreach ($manifestFile in $manifests) {
  $text = Get-Content -LiteralPath $manifestFile.FullName -Raw -Encoding UTF8
  if ($text -notmatch '<Id>([^<]+)</Id>') {
    throw "Cannot find add-in Id in manifest: $($manifestFile.FullName)"
  }
  $addinId = $Matches[1]
  Write-Host "Manifest: $($manifestFile.FullName)"
  reg add 'HKCU\Software\Microsoft\Office\16.0\WEF\Developer' /v $addinId /t REG_SZ /d $manifestFile.FullName /f | Out-Host
}

Write-Host 'Done. Restart PowerPoint if it was already open.'
"""
    write_text(web_dir / "setup_powerpoint_addin.ps1", script)


def update_start_script(script: Path, ppt_name: str) -> None:
    text = read_text(script)
    text = re.sub(r'set "PPT=%~dp0[^"]+\.pptx"', f'set "PPT=%~dp0{ppt_name}"', text, count=1)
    write_text(script, text)


def update_ppt(ppt: Path, addin_id: str, version: str, fallback_png: Path | None) -> None:
    with zipfile.ZipFile(ppt, "r") as z:
        xml = z.read("ppt/webextensions/webextension1.xml").decode("utf-8")
    xml = re.sub(r'<we:reference\s+id="[^"]+"\s+version="[^"]+"',
                 f'<we:reference id="{addin_id}" version="{version}"', xml, count=1)
    zip_replace_text(ppt, "ppt/webextensions/webextension1.xml", xml)
    if fallback_png and fallback_png.exists():
        zip_replace_file(ppt, "ppt/media/soi_dynamic_webpage.png", fallback_png)


def write_readme(package_dir: Path, title: str) -> None:
    readme = f"""关闭所有 PowerPoint 后，双击 START_HERE.bat。
演示时保持弹出的服务器窗口运行。
需要 Microsoft PowerPoint 桌面版 + Python 3；WPS 不支持交互。
"""
    write_text(package_dir / "README_使用说明.txt", readme)


def package_zip(package_dir: Path) -> Path:
    zip_path = package_dir.with_suffix(".zip")
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for path in package_dir.rglob("*"):
            if path.is_file():
                z.write(path, path.relative_to(package_dir.parent).as_posix())
    return zip_path


def build_package(html_path: Path, make_zip: bool = False) -> tuple[Path, Path | None]:
    if not TEMPLATE_DIR.exists():
        raise RuntimeError(f"Template folder not found: {TEMPLATE_DIR}")
    if not html_path.exists() or html_path.suffix.lower() not in (".html", ".htm"):
        raise RuntimeError("Please provide one .html file.")

    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = safe_name(html_path.stem)
    package_name = f"htmlinppt_{stem}_{stamp}"
    package_dir = OUTPUT_DIR / package_name
    web_dir = package_dir / WEB_DIR_NAME
    ppt_name = f"{stem}_interactive_in_ppt.pptx"
    ppt_path = package_dir / ppt_name
    version = f"1.0.{int(time.time()) % 9000}.0"
    addin_id = str(uuid.uuid4())
    url_token = f"{stem}_{stamp}"

    log(f"Creating package: {package_dir}")
    OUTPUT_DIR.mkdir(exist_ok=True)
    copytree_clean(TEMPLATE_DIR, package_dir)

    old_ppt = package_dir / OLD_PPT_NAME
    if old_ppt.exists():
        old_ppt.rename(ppt_path)
    else:
        raise RuntimeError(f"Template PPT not found: {old_ppt}")

    shutil.copy2(html_path, web_dir / "index.html")

    # Keep server/support files, remove the old demo entry pages so the package is less confusing.
    for old_name in ("pd_soi_fbe_simulator_v3.html", "pd_soi_fbe_ppt_content.html", "pd_soi_fbe_visual_config.json"):
        with contextlib.suppress(FileNotFoundError):
            (web_dir / old_name).unlink()

    update_manifest(web_dir / MANIFEST_NAME, addin_id, version, url_token)
    write_setup_script(web_dir)
    update_start_script(package_dir / "02_start_server_and_open_ppt.bat", ppt_name)
    write_readme(package_dir, html_path.stem)
    with contextlib.suppress(FileNotFoundError):
        (package_dir / "01_setup_current_user.bat").unlink()

    fallback = web_dir / FALLBACK_NAME
    screenshot_ok = make_screenshot(package_dir, web_dir, fallback, url_token)
    if screenshot_ok:
        log("Fallback screenshot updated.")
    else:
        log("Fallback screenshot was not updated. The interactive PPT package is still created.")

    update_ppt(ppt_path, addin_id, version, fallback if fallback.exists() else None)
    zip_path = package_zip(package_dir) if make_zip else None
    return package_dir, zip_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert one local HTML file into a PowerPoint interactive folder.")
    parser.add_argument("html", help="Path to one .html file")
    parser.add_argument("--zip", action="store_true", help="Also create a zip file next to the output folder.")
    args = parser.parse_args()

    try:
        html_path = Path(args.html).expanduser().resolve()
        package_dir, zip_path = build_package(html_path, make_zip=args.zip)
    except Exception as exc:
        log("")
        log(f"ERROR: {exc}")
        return 1

    log("")
    log("Done.")
    log(f"Package folder: {package_dir}")
    if zip_path:
        log(f"Zip file: {zip_path}")
    log("")
    log("Open the folder and run START_HERE.bat.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
