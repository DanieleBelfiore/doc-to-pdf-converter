# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Is

Windows-only PowerShell script that recursively converts `.doc`/`.docx` files to PDF using Microsoft Word COM automation. No external dependencies — requires Word installed on the machine.

## Running

**Double-click** `converter.bat` — launches `convert.ps1` with `-ExecutionPolicy Bypass`.

Or directly:
```powershell
PowerShell.exe -ExecutionPolicy Bypass -File convert.ps1
```

## How It Works

1. GUI folder picker (`System.Windows.Forms.FolderBrowserDialog`) selects input directory
2. Creates invisible `Word.Application` COM object
3. Recursively finds `*.doc`/`*.docx`, skipping temp files (`~$*`) and already-converted PDFs
4. Opens each file read-only, saves as `wdFormatPDF` (17) alongside source, closes without saving
5. Quits Word and releases COM object

Output PDFs land next to source files, same directory, same name with `.pdf` extension.
