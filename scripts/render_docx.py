#!/usr/bin/env python3
"""Render a DOCX to page PNGs using LibreOffice and pdftoppm when available."""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Tuple


def find_executable(env_name: str, candidates: list[str]) -> Optional[str]:
    configured = os.environ.get(env_name)
    if configured and Path(configured).exists():
        return configured
    for candidate in candidates:
        found = shutil.which(candidate)
        if found:
            return found
    return None


def convert_with_word(source: Path, pdf: Path) -> Tuple[bool, str]:
    if os.name != "nt":
        return False, "Microsoft Word COM fallback is available only on Windows"
    powershell = find_executable("POWERSHELL_BIN", ["pwsh", "powershell"])
    if not powershell:
        return False, "PowerShell not found for Microsoft Word COM fallback"
    script = r"""
$docPath=$env:ADMIN_DOCX_SOURCE
$pdfPath=$env:ADMIN_DOCX_PDF
$word=$null
$opened=$null
try {
  $word=New-Object -ComObject Word.Application
  $word.Visible=$false
  $word.DisplayAlerts=0
  $opened=$word.Documents.Open($docPath,$false,$true)
  $opened.ExportAsFixedFormat($pdfPath,17)
} finally {
  if($opened){ $opened.Close(0) }
  if($word){ $word.Quit() }
  if($opened){ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($opened) }
  if($word){ [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}
"""
    env = os.environ.copy()
    env["ADMIN_DOCX_SOURCE"] = str(source)
    env["ADMIN_DOCX_PDF"] = str(pdf)
    proc = subprocess.run(
        [powershell, "-NoProfile", "-NonInteractive", "-Command", script],
        capture_output=True,
        text=True,
        env=env,
    )
    message = "\n".join(part for part in (proc.stdout, proc.stderr) if part)
    return proc.returncode == 0 and pdf.exists(), message


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("docx")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--emit-pdf", action="store_true")
    args = parser.parse_args()

    source = Path(args.docx).resolve()
    output_dir = Path(args.output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    soffice = find_executable("SOFFICE_BIN", ["soffice", "libreoffice"])
    pdftoppm = find_executable("PDFTOPPM_BIN", ["pdftoppm"])
    if not pdftoppm:
        print("ERROR: pdftoppm not found; set PDFTOPPM_BIN", file=sys.stderr)
        return 2

    pdf = output_dir / f"{source.stem}.pdf"
    if soffice:
        with tempfile.TemporaryDirectory(prefix="admin_docx_lo_") as profile:
            profile_uri = Path(profile).as_uri()
            cmd = [
                soffice,
                "--headless",
                f"-env:UserInstallation={profile_uri}",
                "--convert-to", "pdf",
                "--outdir", str(output_dir),
                str(source),
            ]
            proc = subprocess.run(cmd, capture_output=True, text=True)
            if proc.returncode != 0 or not pdf.exists():
                print(proc.stdout, file=sys.stderr)
                print(proc.stderr, file=sys.stderr)
                print("ERROR: DOCX to PDF conversion failed", file=sys.stderr)
                return 1
    else:
        ok, message = convert_with_word(source, pdf)
        if not ok:
            print(message, file=sys.stderr)
            print("ERROR: neither LibreOffice nor Microsoft Word conversion is available", file=sys.stderr)
            return 2

    prefix = output_dir / "page"
    proc = subprocess.run([pdftoppm, "-png", "-r", "150", str(pdf), str(prefix)], capture_output=True, text=True)
    if proc.returncode != 0:
        print(proc.stderr, file=sys.stderr)
        return 1
    if not args.emit_pdf:
        pdf.unlink(missing_ok=True)
    for image in sorted(output_dir.glob("page-*.png")):
        print(image)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
