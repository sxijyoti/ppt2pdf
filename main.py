#!/usr/bin/env python3
"""
ppt2pdf - Download PPT/PPTX from Google Drive, convert to PDF, and merge.

Setup (one-time):
  1. Go to https://console.cloud.google.com/
  2. Create a project → enable "Google Drive API"
  3. Create credentials → OAuth 2.0 Client ID → Desktop app
  4. Download the JSON → save as credentials.json in this directory
  5. Run: python main.py <folder_url_or_id>
"""

import io
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional

import click
import logging
import warnings
import atexit
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from pypdf import PdfWriter
from rich.console import Console
from rich.progress import (
    BarColumn,
    DownloadColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TransferSpeedColumn,
)
from rich.prompt import Confirm, Prompt
from rich.table import Table

console = Console()

# Small ASCII banner printed at startup (keeps output compact)
BANNER = r"""

 ██████╗ ██████╗ ████████╗██████╗ ██████╗ ██████╗ ███████╗
██╔══██╗██╔══██╗╚══██╔══╝╚════██╗██╔══██╗██╔══██╗██╔════╝
██████╔╝██████╔╝   ██║    █████╔╝██████╔╝██║  ██║█████╗  
██╔═══╝ ██╔═══╝    ██║   ██╔═══╝ ██╔═══╝ ██║  ██║██╔══╝  
██║     ██║        ██║   ███████╗██║     ██████╔╝██║     
╚═╝     ╚═╝        ╚═╝   ╚══════╝╚═╝     ╚═════╝ ╚═╝     
                                                         
ppt2pdf — download, convert, merge, compress right from your terminal!
"""

VERSION_FILE = Path(__file__).parent / "VERSION"
try:
    VERSION = VERSION_FILE.read_text(encoding="utf-8").strip()
except Exception:
    VERSION = None

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
SUPPORTED_MIME = {
    "application/vnd.ms-powerpoint": ".ppt",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
    "application/vnd.google-apps.presentation": ".pptx",  # Google Slides → export as pptx
    "application/pdf": ".pdf",
}
CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"
TOKEN_FILE = Path(__file__).parent / "token.json"


# ─── Auth ─────────────────────────────────────────────────────────────────────

def authenticate() -> object:
    """Authenticate with Google Drive and return a service client."""
    if not CREDENTIALS_FILE.exists():
        console.print(
            "[bold red]Error:[/] credentials.json not found.\n"
            "Create OAuth 2.0 credentials at https://console.cloud.google.com/\n"
            "and save them as [bold]credentials.json[/] in:\n"
            f"  {CREDENTIALS_FILE.parent}"
        )
        sys.exit(1)

    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())

    return build("drive", "v3", credentials=creds)


# ─── Drive helpers ─────────────────────────────────────────────────────────────

def extract_folder_id(url_or_id: str) -> str:
    """Extract folder ID from a Google Drive URL or return as-is if already an ID."""
    patterns = [
        r"drive\.google\.com/drive/folders/([a-zA-Z0-9_-]+)",
        r"drive\.google\.com/drive/u/\d+/folders/([a-zA-Z0-9_-]+)",
        r"id=([a-zA-Z0-9_-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, url_or_id)
        if match:
            return match.group(1)
    # assume raw ID if no URL pattern matched
    return url_or_id.strip()


def list_presentation_files(service, folder_id: str) -> list[dict]:
    """Return all PPT/PPTX/Google Slides files in a Drive folder, sorted by name."""
    mime_query = " or ".join(
        f"mimeType='{mime}'" for mime in SUPPORTED_MIME
    )
    query = f"'{folder_id}' in parents and ({mime_query}) and trashed=false"

    results = []
    page_token = None
    try:
        while True:
            resp = (
                service.files()
                .list(
                    q=query,
                    spaces="drive",
                    fields="nextPageToken, files(id, name, mimeType, size)",
                    pageToken=page_token,
                    orderBy="name",
                )
                .execute()
            )
            results.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
    except HttpError as e:
        _handle_http_error(e, folder_id)

    return results


def _handle_http_error(e: HttpError, context: str = "") -> None:
    """Print a helpful message for common Google API HTTP errors and exit."""
    status = e.resp.status
    if status == 403:
        console.print(
            "\n[bold red]403 Forbidden[/] — access denied.\n"
            "\nCommon causes and fixes:\n"
            "  [yellow]1.[/] [bold]Drive API not enabled[/] on your project\n"
            "       → console.cloud.google.com → APIs & Services → Library → \"Google Drive API\" → Enable\n"
            "  [yellow]2.[/] [bold]OAuth app in Testing mode[/] and your account isn't a test user\n"
            "       → APIs & Services → OAuth consent screen → Test users → add your Google email\n"
            "  [yellow]3.[/] [bold]Folder not shared[/] with the account you signed in as\n"
            "       → Open the Drive folder → Share → add your Google email with Viewer access\n"
            "  [yellow]4.[/] Stale token — delete [bold]token.json[/] and re-run to re-authenticate\n"
        )
    elif status == 404:
        console.print(
            f"\n[bold red]404 Not Found[/] — folder [bold]{context}[/] doesn't exist or\n"
            "isn't accessible. Double-check the URL/ID.\n"
        )
    elif status == 401:
        console.print(
            "\n[bold red]401 Unauthorized[/] — token is invalid.\n"
            "Delete [bold]token.json[/] and re-run to sign in again.\n"
        )
    else:
            # ── Manual reordering ──────────────────────────────────────────────
            def prompt_order(files):
                console.print("\n[bold]You can rearrange the order before downloading.[/]")
                console.print("Enter a comma-separated list of numbers (e.g. 3,1,2,4-6) for your desired order, or press Enter to keep as listed.")
                idxs = Prompt.ask("Order", default=",").replace(" ", "")
                if not idxs or idxs == ",":
                    return files
                try:
                    order = [int(i)-1 for i in idxs.split(",") if i]
                    if sorted(order) != list(range(len(files))):
                        raise ValueError
                    return [files[i] for i in order]
                except Exception:
                    console.print("[red]Invalid order. Using default.[/]")
                    return files

    sys.exit(1)


def download_file(service, file_info: dict, dest_path: Path) -> None:
    mime = file_info["mimeType"]

    if mime == "application/vnd.google-apps.presentation":
        # Export Google Slides as PPTX
        request = service.files().export_media(
            fileId=file_info["id"],
            mimeType="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    else:
        request = service.files().get_media(fileId=file_info["id"])

    try:
        with open(dest_path, "wb") as fh:
            downloader = MediaIoBaseDownload(fh, request, chunksize=4 * 1024 * 1024)
            done = False
            while not done:
                _, done = downloader.next_chunk()
    except HttpError as e:
        _handle_http_error(e, file_info.get("name", ""))


# ─── Conversion ───────────────────────────────────────────────────────────────

def find_libreoffice() -> Optional[str]:
    """Find the LibreOffice binary."""
    candidates = [
        "libreoffice",
        "soffice",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/lib/libreoffice/program/soffice",
    ]
    for candidate in candidates:
        path = shutil.which(candidate) or (candidate if os.path.isfile(candidate) else None)
        if path:
            return path
    return None

def convert_to_pdf(soffice: str, pptx_path: Path, output_dir: Path) -> Path:
    """Convert a PPTX/PPT file to PDF using LibreOffice. Returns the PDF path."""
    result = subprocess.run(
        [
            soffice,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(pptx_path),
        ],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(
            f"LibreOffice conversion failed for {pptx_path.name}:\n{result.stderr}"
        )
    # LibreOffice names the output <stem>.pdf
    pdf_path = output_dir / (pptx_path.stem + ".pdf")
    if not pdf_path.exists():
        raise RuntimeError(f"Expected PDF not found: {pdf_path}")
    return pdf_path


# ─── PDF merge & compress ─────────────────────────────────────────────────────

def merge_pdfs(pdf_paths: list[Path], output_path: Path) -> None:
    """Merge multiple PDFs into one."""
    # Some PDF libraries emit verbose annotation-size warnings to stderr;
    # temporarily redirect stderr to suppress them during merge.
    writer = PdfWriter()
    import sys
    import os
    orig_stderr = sys.stderr
    try:
        sys.stderr = open(os.devnull, "w")
        for pdf in pdf_paths:
            writer.append(str(pdf))
        with open(output_path, "wb") as f:
            writer.write(f)
    finally:
        try:
            sys.stderr.close()
        except Exception:
            pass
        sys.stderr = orig_stderr


def compress_pdf_gs(input_path: Path, output_path: Path, quality: str = "ebook") -> bool:
    """
    Compress PDF with Ghostscript.
    quality: screen | ebook | printer | prepress (increasing size/quality)
    Returns True on success.
    """
    gs_bin = shutil.which("gs") or shutil.which("ghostscript")
    if not gs_bin:
        return False

    result = subprocess.run(
        [
            gs_bin,
            "-sDEVICE=pdfwrite",
            "-dCompatibilityLevel=1.4",
            f"-dPDFSETTINGS=/{quality}",
            "-dNOPAUSE",
            "-dQUIET",
            "-dBATCH",
            f"-sOutputFile={output_path}",
            str(input_path),
        ],
        capture_output=True,
        text=True,
    )
    return result.returncode == 0 and output_path.exists()


def human_bytes(size: int) -> str:
    for unit in ["B", "KB", "MB", "GB"]:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size = float(size) / 1024 # type: ignore
    return f"{size:.1f} TB"


# ─── CLI ──────────────────────────────────────────────────────────────────────

@click.version_option(version=VERSION or "0.0.0", prog_name="ppt2pdf")
@click.command(context_settings={"help_option_names": ["-h", "--help"]})
@click.argument("folder", metavar="FOLDER_URL_OR_ID", required=False)
@click.option(
    "--sort",
    "sort_by",
    type=click.Choice(["name", "created", "modified"], case_sensitive=False),
    default="name",
    show_default=True,
    help="Order files by this field before numbering.",
)
@click.option(
    "--quality",
    type=click.Choice(["screen", "ebook", "printer", "prepress"], case_sensitive=False),
    default="ebook",
    show_default=True,
    help="Ghostscript compression quality (if compression is chosen).",
)
@click.option(
    "--work-dir",
    type=click.Path(file_okay=False, writable=True),
    default=None,
    help="Temporary working directory (default: system temp).",
)
def cli(folder: str, sort_by: str, quality: str, work_dir: Optional[str]):
    """
    Download PPT/PPTX from a Google Drive folder, convert to PDF, and merge.

    \b
    FOLDER_URL_OR_ID can be:
      • A full Google Drive folder URL
      • A raw folder ID

    \b
    One-time setup:
      1. https://console.cloud.google.com/ → enable Drive API
      2. Create OAuth 2.0 credentials (Desktop app)
      3. Download JSON → save as credentials.json beside this script
    """
    # Prompt for folder if not provided; accept full URL or raw ID
    folder_input = folder if folder else Prompt.ask("Paste Google Drive folder link or ID (URL or ID)")
    # strip surrounding quotes and whitespace
    folder_input = folder_input.strip().strip('"').strip("'")
    folder_id = extract_folder_id(folder_input)
    console.rule("[bold blue]ppt2pdf[/]")

    # ── 1. Auth ────────────────────────────────────────────────────────────────
    with console.status("[bold]Authenticating with Google Drive…"):
        service = authenticate()
    console.print("[green]✓[/] Authenticated")

    # ── 2. List files ──────────────────────────────────────────────────────────
    with console.status("[bold]Fetching file list…"):
        files = list_presentation_files(service, folder_id)

    if not files:
        console.print("[yellow]No PPT/PPTX/Google Slides files found in that folder.[/]")
        sys.exit(0)

    # Apply sort order
    if sort_by == "name":
        files.sort(key=lambda f: f["name"].lower())
    # (Drive API already sorts by name by default; created/modified would require
    #  extra fields — for now name sort is most useful)

    table = Table(title=f"Found {len(files)} file(s)", show_lines=False)
    table.add_column("#", style="dim", width=4)
    table.add_column("Name", style="bold")
    table.add_column("Type", style="cyan")
    for i, f in enumerate(files, 1):
        ext = SUPPORTED_MIME.get(f["mimeType"], "?")
        table.add_row(str(i), f["name"], ext)
    console.print(table)

    # Let user choose default or manual ordering; accept shorthand (d/D or enter) and m/M
    raw = Prompt.ask("Choose order — (D)efault / (M)anual [Enter=default]", default="").strip()
    choice = raw.lower()
    if choice in ("m", "manual"):
        console.print("Enter a comma-separated list of numbers (e.g. 3,1,2,4-6) for your desired order.")
        idxs = Prompt.ask("Order", default=",").replace(" ", "")
        try:
            n = len(files)
            tokens = [t for t in idxs.split(",") if t]
            order_indices: list[int] = []
            seen = set()
            for t in tokens:
                if "-" in t:
                    a_s, b_s = t.split("-", 1)
                    a = int(a_s)
                    b = int(b_s)
                    if a < 1 or b < 1 or a > n or b > n:
                        raise ValueError("range out of bounds")
                    if a <= b:
                        seq = list(range(a - 1, b))
                    else:
                        seq = list(range(a - 1, b - 2, -1))
                    for idx in seq:
                        if idx in seen:
                            raise ValueError("duplicate index")
                        order_indices.append(idx)
                        seen.add(idx)
                else:
                    idx = int(t) - 1
                    if idx < 0 or idx >= n:
                        raise ValueError("index out of bounds")
                    if idx in seen:
                        raise ValueError("duplicate index")
                    order_indices.append(idx)
                    seen.add(idx)

            # allow subsets: user may skip some files intentionally
            if len(order_indices) > 0 and len(seen) == len(order_indices):
                files = [files[i] for i in order_indices]
            else:
                console.print("[red]Order invalid — keeping default order.[/]")
        except Exception:
            console.print("[red]Invalid input — keeping default order.[/]")

    if not Confirm.ask("Proceed with download and conversion?", default=True):
        sys.exit(0)

    # ── 3. Work directory ──────────────────────────────────────────────────────
    repo_dir = Path(__file__).parent.resolve()
    if work_dir:
        tmp_root = Path(work_dir)
    else:
        tmp_root = repo_dir / "downloads"
    if tmp_root.exists():
        shutil.rmtree(tmp_root, ignore_errors=True)
    tmp_root.mkdir(parents=True, exist_ok=True)

    # ensure downloads are removed on any exit (including Ctrl+C)
    def _cleanup_tmp():
        try:
            if tmp_root and not work_dir and tmp_root.exists():
                shutil.rmtree(tmp_root, ignore_errors=True)
        except Exception:
            pass

    atexit.register(_cleanup_tmp)

    pptx_dir = tmp_root / "pptx"
    pdf_dir = tmp_root / "pdf"
    pptx_dir.mkdir(exist_ok=True)
    pdf_dir.mkdir(exist_ok=True)

    # ── 4. Download ────────────────────────────────────────────────────────────
    console.rule("[bold]Downloading")
    pad = len(str(len(files)))  # zero-pad width
    downloaded: list[Path] = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[bold]{task.description}"),
        BarColumn(),
        TextColumn("{task.percentage:>3.0f}%"),
        console=console,
    ) as progress:
        task = progress.add_task("Downloading…", total=len(files))
        for i, f in enumerate(files, 1):
            ext = SUPPORTED_MIME.get(f["mimeType"], ".pptx")
            seq_name = f"{str(i).zfill(pad)}_{f['name']}"
            # Ensure the extension is present
            if not seq_name.lower().endswith(ext):
                seq_name += ext
            dest = pptx_dir / seq_name
            progress.update(task, description=f"[{i}/{len(files)}] {f['name'][:50]}")
            download_file(service, f, dest)
            downloaded.append(dest)
            progress.advance(task)

    console.print(f"[green]✓[/] Downloaded {len(downloaded)} file(s) to {pptx_dir}")

    # ── 5. Convert ─────────────────────────────────────────────────────────────
    console.rule("[bold]Converting to PDF")
    soffice = find_libreoffice()
    if not soffice:
        console.print(
            "[bold red]Error:[/] LibreOffice not found.\n"
            "Install with: [bold]brew install --cask libreoffice[/]"
        )
        shutil.rmtree(tmp_root, ignore_errors=True)
        sys.exit(1)

    pdf_paths: list[Path] = []
    errors: list[str] = []

    with Progress(
        SpinnerColumn(),
        TextColumn("[bold]{task.description}"),
        BarColumn(),
        TextColumn("{task.percentage:>3.0f}%"),
        console=console,
    ) as progress:
        task = progress.add_task("Converting…", total=len(downloaded))
        for i, src in enumerate(sorted(downloaded), 1):
            progress.update(task, description=f"[{i}/{len(downloaded)}] {src.name[:50]}")
            try:
                if src.suffix.lower() == ".pdf":
                    # already a PDF — no conversion needed
                    dest_pdf = pdf_dir / src.name
                    # copy or move the downloaded pdf into pdf_dir if not already there
                    if src.resolve() != dest_pdf.resolve():
                        shutil.copy2(src, dest_pdf)
                    pdf_paths.append(dest_pdf)
                else:
                    pdf = convert_to_pdf(soffice, src, pdf_dir)
                    pdf_paths.append(pdf)
            except RuntimeError as e:
                console.print(f"[yellow]⚠ Skipped {src.name}: {e}[/]")
                errors.append(src.name)
            progress.advance(task)

    if not pdf_paths:
        console.print("[bold red]No PDFs were produced. Aborting.[/]")
        shutil.rmtree(tmp_root, ignore_errors=True)
        sys.exit(1)

    console.print(f"[green]✓[/] Converted {len(pdf_paths)} file(s)")
    if errors:
        console.print(f"[yellow]  ⚠ {len(errors)} file(s) failed conversion[/]")

    # PDFs are named after their source PPTX (e.g., 01_Lecture.pdf), so sort by name
    pdf_paths.sort(key=lambda p: p.name.lower())

    # ── 6. Merge ───────────────────────────────────────────────────────────────
    console.rule("[bold]Merging PDFs")
    merged_tmp = tmp_root / "merged_raw.pdf"
    with console.status(f"Merging {len(pdf_paths)} PDF(s)…"):
        merge_pdfs(pdf_paths, merged_tmp)
    raw_size = merged_tmp.stat().st_size
    console.print(f"[green]✓[/] Merged → {human_bytes(raw_size)}")

    # ── 7. Compression ─────────────────────────────────────────────────────────
    final_pdf = merged_tmp
    compressed_tmp = tmp_root / "merged_compressed.pdf"

    gs_available = bool(shutil.which("gs") or shutil.which("ghostscript"))
    if not gs_available:
        console.print(
            "[dim]Ghostscript not found — compression unavailable. "
            "Install with: brew install ghostscript[/]"
        )
        do_compress = False
    else:
        do_compress = Confirm.ask(
            f"Compress the PDF? (current size: {human_bytes(raw_size)})", default=True
        )

    if do_compress:
        with console.status(f"Compressing (quality=[bold]{quality}[/])…"):
            ok = compress_pdf_gs(merged_tmp, compressed_tmp, quality=quality)
        if ok:
            comp_size = compressed_tmp.stat().st_size
            pct = (1 - comp_size / raw_size) * 100
            console.print(
                f"[green]✓[/] Compressed: {human_bytes(raw_size)} → "
                f"{human_bytes(comp_size)} ([bold green]{pct:.1f}% smaller[/])"
            )
            final_pdf = compressed_tmp
        else:
            console.print("[yellow]⚠ Compression failed, keeping uncompressed version.[/]")

    # ── 8. Name & destination ──────────────────────────────────────────────────
    console.rule("[bold]Save merged PDF")
    default_name = "merged"
    output_name = Prompt.ask("Name for the merged PDF (without .pdf)", default=default_name)
    output_name = output_name.strip().rstrip(".pdf").rstrip()
    if not output_name:
        output_name = default_name
    output_filename = output_name + ".pdf"

    default_dest = str(Path.cwd())
    dest_dir_str = Prompt.ask("Save to directory", default=default_dest)
    dest_dir = Path(dest_dir_str).expanduser().resolve()
    dest_dir.mkdir(parents=True, exist_ok=True)

    final_dest = dest_dir / output_filename

    # Avoid overwrite collision
    counter = 1
    while final_dest.exists():
        final_dest = dest_dir / f"{output_name}_{counter}.pdf"
        counter += 1

    shutil.copy2(final_pdf, final_dest)
    console.print(f"[green]✓[/] Saved to [bold]{final_dest}[/]")

    # ── 9. Cleanup ─────────────────────────────────────────────────────────────
    with console.status("Cleaning up temporary files…"):
        shutil.rmtree(tmp_root, ignore_errors=True)
    console.print("[green]✓[/] Temporary files deleted")

    console.rule("[bold green]Done![/]")
    console.print(f"\n[bold]Output:[/] {final_dest}  ({human_bytes(final_dest.stat().st_size)})")


if __name__ == "__main__":
    try:
        console.print(BANNER, style="#B7BDF7 bold")
        try:
            ver = VERSION
        except NameError:
            ver = None
        if ver:
            console.print(f"[dim]version {ver}[/]\n")
    except Exception:
        pass
    cli()
