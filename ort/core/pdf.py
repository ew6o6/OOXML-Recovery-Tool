"""PDF preview of the rebuilt document, made by an office application that is already installed.

Order: LibreOffice (soffice, any platform), then Microsoft Office through COM automation (Windows). If neither
is available, nothing is converted and the reason is recorded. The manifest field always says what happened
("written", "skipped", "failed", or "disabled" with --no-pdf) and, unless a PDF was written, how to obtain one.
Each conversion has a time limit; a timed-out Office instance started for automation is stopped.
"""

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

TIMEOUT_SECONDS = 180
_OFFICE_APPS = {".docx": "Microsoft Word", ".xlsx": "Microsoft Excel", ".pptx": "Microsoft PowerPoint"}
_SOFFICE_PATHS = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]
# Exit codes: 0 written, 2 conversion failed, 3 application not installed.
_OFFICE_SCRIPT = r"""
param([string]$In, [string]$Out)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$ext = [IO.Path]::GetExtension($In).ToLower()
$app = $null
try {
    switch ($ext) {
        '.docx' { try { $app = New-Object -ComObject Word.Application } catch { exit 3 }
                  $app.Visible = $false; $app.DisplayAlerts = 0
                  $doc = $app.Documents.Open($In, $false, $true); $doc.SaveAs2($Out, 17); $doc.Close($false) }
        '.xlsx' { try { $app = New-Object -ComObject Excel.Application } catch { exit 3 }
                  $app.Visible = $false; $app.DisplayAlerts = $false
                  $book = $app.Workbooks.Open($In, 0, $true); $book.ExportAsFixedFormat(0, $Out); $book.Close($false) }
        '.pptx' { try { $app = New-Object -ComObject PowerPoint.Application } catch { exit 3 }
                  $pres = $app.Presentations.Open($In, -1, 0, 0); $pres.SaveAs($Out, 32); $pres.Close() }
    }
} catch { Write-Output $_.Exception.Message; exit 2 }
finally { if ($app) { try { $app.Quit() } catch {} } }
exit 0
"""


def how_to(rebuilt_name, disabled=False):
    steps = ["Run ORT again without --no-pdf." if disabled else None,
             "Install LibreOffice (with 'soffice' on PATH) and run ORT again, or",
             "on Windows, install Microsoft Office and run ORT again, or",
             f"open {rebuilt_name} in Word, Excel, PowerPoint, or LibreOffice and save it as PDF."]
    return [step for step in steps if step]


def find_soffice():
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    return next((path for path in _SOFFICE_PATHS if Path(path).exists()), None)


def _with_soffice(soffice, source):
    with tempfile.TemporaryDirectory() as profile:  # a private profile allows parallel conversions
        command = [soffice, f"-env:UserInstallation={Path(profile).as_uri()}", "--headless", "--norestore",
                   "--convert-to", "pdf", "--outdir", str(source.parent), str(source)]
        process = subprocess.run(command, capture_output=True, timeout=TIMEOUT_SECONDS)
    produced = source.with_suffix(".pdf")
    if process.returncode == 0 and produced.exists():
        return "written", None
    return "failed", (process.stdout + process.stderr).decode("utf-8", "replace").strip()[-300:] or "no PDF produced"


def _stop_automation_instances():
    subprocess.run(["powershell", "-NoProfile", "-Command",
                    "Get-CimInstance Win32_Process -Filter \"Name='WINWORD.EXE' OR Name='EXCEL.EXE' OR "
                    "Name='POWERPNT.EXE'\" | Where-Object { $_.CommandLine -match '(?i)automation|embedding' } | "
                    "ForEach-Object { Stop-Process -Id $_.ProcessId -Force }"], capture_output=True)


def _with_office(source):
    """("written" | "failed" | "unavailable", message)."""
    with tempfile.TemporaryDirectory() as folder:
        script = Path(folder) / "ort_pdf.ps1"
        script.write_text(_OFFICE_SCRIPT, encoding="utf-8-sig")
        try:
            process = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(script),
                                      str(source), str(source.with_suffix(".pdf"))],
                                     capture_output=True, timeout=TIMEOUT_SECONDS)
        except subprocess.TimeoutExpired:
            _stop_automation_instances()
            return "failed", f"timed out after {TIMEOUT_SECONDS} s"
    if process.returncode == 3:
        return "unavailable", None
    if process.returncode == 0 and source.with_suffix(".pdf").exists():
        return "written", None
    return "failed", (process.stdout + process.stderr).decode("utf-8", "replace").strip()[-300:] or "no PDF produced"


def make_preview(rebuilt_path, enabled=True):
    """Convert the rebuilt document to PDF next to it; return the manifest field preview_pdf."""
    name = rebuilt_path.name if rebuilt_path else "the rebuilt document"
    field = {"status": None, "converter": None, "path": None, "reason": None, "how_to": None}
    if not enabled:
        return {**field, "status": "disabled", "reason": "disabled by --no-pdf", "how_to": how_to(name, disabled=True)}
    if rebuilt_path is None or not rebuilt_path.exists():
        return {**field, "status": "skipped", "reason": "no rebuilt document to convert"}
    errors = []
    soffice = find_soffice()
    if soffice:
        try:
            status, message = _with_soffice(soffice, rebuilt_path)
        except subprocess.TimeoutExpired:
            status, message = "failed", f"timed out after {TIMEOUT_SECONDS} s"
        if status == "written":
            return {**field, "status": "written", "converter": "LibreOffice", "path": rebuilt_path.with_suffix(".pdf").name}
        errors.append(f"LibreOffice: {message}")
    if sys.platform == "win32" and shutil.which("powershell"):
        status, message = _with_office(rebuilt_path)
        if status == "written":
            return {**field, "status": "written", "converter": _OFFICE_APPS[rebuilt_path.suffix.lower()],
                    "path": rebuilt_path.with_suffix(".pdf").name}
        if status == "failed":
            errors.append(f"{_OFFICE_APPS[rebuilt_path.suffix.lower()]}: {message}")
    if errors:
        return {**field, "status": "failed", "reason": "; ".join(errors), "how_to": how_to(name)}
    return {**field, "status": "skipped",
            "reason": "no PDF converter found (LibreOffice; Microsoft Office on Windows)", "how_to": how_to(name)}
