# ppt2pdf

```
██████╗ ██████╗ ████████╗██████╗ ██████╗ ██████╗ ███████╗
██╔══██╗██╔══██╗╚══██╔══╝╚════██╗██╔══██╗██╔══██╗██╔════╝
██████╔╝██████╔╝   ██║    █████╔╝██████╔╝██║  ██║█████╗  
██╔═══╝ ██╔═══╝    ██║   ██╔═══╝ ██╔═══╝ ██║  ██║██╔══╝  
██║     ██║        ██║   ███████╗██║     ██████╔╝██║     
╚═╝     ╚═╝        ╚═╝   ╚══════╝╚═╝     ╚═════╝ ╚═╝     
                                                         
ppt2pdf — download, convert, merge, compress right from your terminal!
```

ppt2pdf is a command-line utility that downloads presentations (PPT, PPTX, Google Slides) and PDFs from a Google Drive folder, converts presentations to PDF as needed, and merges selected PDFs into a single file.

This document describes how to install, configure, and use the tool. 

## Features

- Fetch presentations and PDFs from a Google Drive folder
- Preserve a user-specified order (supports individual indices and ranges)
- Convert PPT/PPTX/Slides to PDF using LibreOffice headless
- Include existing PDFs without conversion
- Merge PDFs with `pypdf` and optionally compress the result with Ghostscript
- Clean up intermediate files automatically (safe default)

## Requirements

- Python 3.10+ (3.11 recommended)
- LibreOffice (for PPT/PPTX -> PDF conversion)
- Ghostscript (optional, for PDF compression)

Platform install notes:

- macOS (Homebrew):
	- LibreOffice: `brew install --cask libreoffice`
	- Ghostscript: `brew install ghostscript`
	- Python 3.11: `brew install python@3.11`

- Ubuntu / Debian:
	- LibreOffice: `sudo apt update && sudo apt install -y libreoffice`
	- Ghostscript: `sudo apt install -y ghostscript`

- Windows (recommended: WSL):
	- Install WSL2, then follow the Linux instructions inside the WSL environment.
	- Alternatively, use Chocolatey (`choco install libreoffice ghostscript python`) but behavior may vary.

Create and activate a virtual environment and install Python dependencies:

```bash
python3.11 -m venv .venv || python3 -m venv .venv
source .venv/bin/activate
.venv/bin/pip install -r requirements.txt
```

## Google Drive API setup (one-time)

1. In the Google Cloud Console enable the *Google Drive API* for a project.
2. Create OAuth 2.0 credentials (Application type: *Desktop app*).
3. Download the credentials JSON and save it as `credentials.json` in the project root.

On first run the tool opens a browser to authorize access and caches a `token.json` file.

## Usage

Run the tool and follow the interactive prompts:

```bash
source .venv/bin/activate
python main.py
```
### Scripts & helpers

The repository includes convenience helpers:

- `scripts/bootstrap.sh` — creates `.venv` (prefers `python3.11` if available) and installs dependencies. Make it executable if needed: `chmod +x scripts/bootstrap.sh`.
- `Makefile` targets:
	- `make setup` — runs the bootstrap script
	- `make run` — runs the CLI using the virtualenv python
	- `make clean` — removes `.venv` and `downloads/`

Examples:

```bash
./scripts/bootstrap.sh
# or
make setup

make run
# or
source .venv/bin/activate && python main.py
```

## Interactive prompts:

- Paste a Google Drive folder URL or raw folder ID (both accepted)
- Choose ordering: press Enter or type `d` for default order, or `m` to enter manual order
- When manual: enter indices and ranges (examples: `3,1,2`, `3-8,1-2,9-11`, `5,2-4`); ranges may be ascending or descending
- Choose whether to compress the merged PDF and the output filename/location

Command-line options (use `python main.py --help`):

- `--sort` — initial sort applied before numbering (`name` by default)
- `--quality` — Ghostscript compression preset (`screen`, `ebook`, `printer`, `prepress`)
- `--work-dir` — specify an alternate directory for intermediates; when omitted a `downloads/` folder inside the repo is used and removed after completion

## Ordering rules

- Indices are 1-based and refer to the order displayed after listing files.
- Manual input supports single indices and ranges (e.g. `3-6`).
- The user may specify any subset of files — the tool will download and merge only the selected items in the provided order.

## Output & cleanup

- The merged PDF is written to the user-specified location.
- By default intermediate PPT/PPTX and per-file PDFs are stored in a local `downloads/` folder and removed after completion or on abort. If a custom `--work-dir` is provided, that directory is preserved by default.

## Compression options

The tool uses Ghostscript presets for compression. Available presets:

- `screen` — lowest quality, smallest size (suitable for preview)
- `ebook` — balanced quality/size (default)
- `printer` — higher quality for printing
- `prepress` — minimal compression, preserves quality

## Troubleshooting

- 403 / Access Denied: confirm the Drive folder is shared with the Google account used to authorize the tool; ensure the Drive API is enabled in the Cloud Console and, if the OAuth consent screen is in testing mode, add the Google account as a test user.
- LibreOffice not found: install with Homebrew: `brew install --cask libreoffice`.
- Compression unavailable: install Ghostscript: `brew install ghostscript`.

## License & contribution

- **License:** This project is licensed under the MIT License - see the `LICENSE` file.

---
