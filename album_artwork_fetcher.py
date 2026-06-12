#!/usr/bin/env python3
"""
album_artwork_fetcher.py
========================

Pull high-resolution album artwork from Apple Music (via the public iTunes
Search API) for a list of albums stored in an Excel file, with a lightweight
tkinter "Approve / Deny" review UI and resume-from-where-you-left-off support.

How the "high resolution" trick works
--------------------------------------
The iTunes Search API returns an ``artworkUrl100`` that looks like::

    https://is1-ssl.mzstatic.com/.../source/100x100bb.jpg

The image is served by Apple's CDN, and the ``100x100bb`` (or ``600x600bb``)
segment near the end of the URL is just a *resize instruction*. If we swap that
segment for an absurdly large value such as ``10000x10000bb``, Apple's CDN
clamps the request down to the largest original master it actually has on file
(commonly 1400x1400, 3000x3000, or larger). This is exactly the technique used
by Ben Dodson's "Apple Music Artwork Finder".

See the README for setup / usage instructions.
"""

import io
import json
import os
import re
import sys
import time
from dataclasses import dataclass, asdict
from typing import Optional

import requests

try:
    import openpyxl
except ImportError:  # pragma: no cover - friendly guidance
    sys.exit("Missing dependency 'openpyxl'. Install with: pip install openpyxl")

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:  # pragma: no cover
    sys.exit("tkinter is not available in this Python build. On Debian/Ubuntu: "
             "sudo apt-get install python3-tk")

try:
    from PIL import Image, ImageTk
except ImportError:  # pragma: no cover
    sys.exit("Missing dependency 'Pillow'. Install with: pip install Pillow")


# --------------------------------------------------------------------------- #
# Configuration                                                               #
# --------------------------------------------------------------------------- #

EXCEL_FILE = "albums.xlsx"          # input spreadsheet (columns: Artist, Album)
ARTIST_COLUMN = "Artist"            # column header for the artist name
ALBUM_COLUMN = "Album"              # column header for the album title
OUTPUT_FOLDER = "HighRes_Covers"    # where approved covers are saved
PROGRESS_FILE = "progress.json"     # tracks per-row status so we can resume

ITUNES_ENDPOINT = "https://itunes.apple.com/search"
REQUEST_TIMEOUT = 15                # seconds for any single network request
MAX_PREVIEW_SIZE = 700             # max width/height (px) for the on-screen preview
TARGET_RESOLUTION = "10000x10000bb"  # request size; CDN downscales to the real max

# Matches the "<width>x<height>bb" size token that precedes the file extension,
# e.g. "100x100bb", "600x600bb.jpg", "1200x1200bb-60.jpg". We only rewrite the
# final size token so query strings / other path parts are left untouched.
SIZE_TOKEN_RE = re.compile(r"\d+x\d+bb")


# --------------------------------------------------------------------------- #
# State management                                                            #
# --------------------------------------------------------------------------- #

@dataclass
class AlbumRecord:
    """A single album plus its review status."""
    artist: str
    album: str
    status: str = "pending"          # pending | approved | denied | not_found | error
    saved_path: Optional[str] = None
    note: Optional[str] = None        # error message / extra info


class ProgressStore:
    """Loads and persists per-album status to a small JSON file.

    The key for each album is "Artist|||Album" so the file stays human
    readable and we can match rows back to the spreadsheet even if their
    order changes.
    """

    def __init__(self, path: str):
        self.path = path
        self._data: dict = {}
        self._load()

    @staticmethod
    def _key(artist: str, album: str) -> str:
        return f"{artist.strip()}|||{album.strip()}"

    def _load(self) -> None:
        if os.path.exists(self.path):
            try:
                with open(self.path, "r", encoding="utf-8") as fh:
                    self._data = json.load(fh)
            except (json.JSONDecodeError, OSError) as exc:
                print(f"[warn] Could not read {self.path} ({exc}); starting fresh.")
                self._data = {}

    def save(self) -> None:
        # Write to a temp file then rename, so an interrupted write can't
        # corrupt existing progress.
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as fh:
            json.dump(self._data, fh, indent=2, ensure_ascii=False)
        os.replace(tmp, self.path)

    def get_status(self, artist: str, album: str) -> Optional[dict]:
        return self._data.get(self._key(artist, album))

    def is_done(self, artist: str, album: str) -> bool:
        """Has this album already been reviewed (any terminal status)?"""
        rec = self.get_status(artist, album)
        # "error" is NOT terminal: we want to retry transient failures on resume.
        return bool(rec) and rec.get("status") in {"approved", "denied", "not_found"}

    def update(self, record: AlbumRecord) -> None:
        self._data[self._key(record.artist, record.album)] = asdict(record)
        self.save()


# --------------------------------------------------------------------------- #
# Apple Music / iTunes lookup                                                 #
# --------------------------------------------------------------------------- #

def build_highres_url(artwork_url: str, target: str = TARGET_RESOLUTION) -> str:
    """Rewrite an iTunes artwork URL to request the maximum resolution.

    The last "<n>x<n>bb" token in the path is replaced with ``target``.
    If no size token is found the URL is returned unchanged.
    """
    matches = list(SIZE_TOKEN_RE.finditer(artwork_url))
    if not matches:
        return artwork_url
    last = matches[-1]
    return artwork_url[:last.start()] + target + artwork_url[last.end():]


def lookup_album_artwork(artist: str, album: str) -> Optional[str]:
    """Query the iTunes Search API and return a high-res artwork URL, or None.

    Raises requests.RequestException on network failure so the caller can
    distinguish "not found" (return None) from "network error" (exception).
    """
    query = f"{artist} {album}".strip()
    params = {
        "term": query,
        "entity": "album",
        "limit": 5,        # grab a few; the first hit is usually correct
        "media": "music",
    }
    resp = requests.get(ITUNES_ENDPOINT, params=params, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    if not results:
        return None

    # Prefer an exact-ish album-title match; fall back to the first result.
    best = None
    album_norm = album.strip().lower()
    for item in results:
        if item.get("collectionName", "").strip().lower() == album_norm:
            best = item
            break
    if best is None:
        best = results[0]

    artwork = best.get("artworkUrl100") or best.get("artworkUrl60")
    if not artwork:
        return None
    return build_highres_url(artwork)


def download_image_bytes(url: str) -> bytes:
    """Download raw image bytes (used both for preview and for saving)."""
    resp = requests.get(url, timeout=REQUEST_TIMEOUT)
    resp.raise_for_status()
    return resp.content


def safe_filename(artist: str, album: str) -> str:
    """Build a filesystem-safe filename from artist + album."""
    raw = f"{artist} - {album}"
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", raw).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return f"{cleaned}.jpg"


# --------------------------------------------------------------------------- #
# Review UI                                                                   #
# --------------------------------------------------------------------------- #

class ReviewApp:
    """Tkinter app that walks through pending albums one at a time."""

    def __init__(self, albums: list[AlbumRecord], store: ProgressStore):
        self.albums = albums
        self.store = store
        self.index = 0
        # Full-resolution bytes of the currently displayed image (for saving).
        self._current_bytes: Optional[bytes] = None
        self._current_url: Optional[str] = None
        self._tk_image: Optional[ImageTk.PhotoImage] = None  # keep a reference!

        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        self.root = tk.Tk()
        self.root.title("Apple Music High-Res Artwork Reviewer")
        self.root.configure(bg="#1e1e1e")

        self._build_widgets()
        self._bind_keys()

        # Kick off the first album after the window is shown.
        self.root.after(100, self.show_current)

    # -- widget construction ------------------------------------------------ #
    def _build_widgets(self) -> None:
        self.title_label = tk.Label(
            self.root, text="", font=("Helvetica", 16, "bold"),
            fg="#ffffff", bg="#1e1e1e", wraplength=MAX_PREVIEW_SIZE, justify="center",
        )
        self.title_label.pack(pady=(14, 2))

        self.subtitle_label = tk.Label(
            self.root, text="", font=("Helvetica", 12),
            fg="#bbbbbb", bg="#1e1e1e", wraplength=MAX_PREVIEW_SIZE, justify="center",
        )
        self.subtitle_label.pack(pady=(0, 4))

        self.progress_label = tk.Label(
            self.root, text="", font=("Helvetica", 10),
            fg="#888888", bg="#1e1e1e",
        )
        self.progress_label.pack(pady=(0, 8))

        self.image_label = tk.Label(self.root, bg="#1e1e1e")
        self.image_label.pack(padx=20, pady=10)

        self.resolution_label = tk.Label(
            self.root, text="", font=("Helvetica", 10, "italic"),
            fg="#66cc99", bg="#1e1e1e",
        )
        self.resolution_label.pack(pady=(0, 8))

        button_frame = tk.Frame(self.root, bg="#1e1e1e")
        button_frame.pack(pady=(4, 16))

        self.approve_btn = tk.Button(
            button_frame, text="✔ Approve (Y / Enter)",
            command=self.approve, font=("Helvetica", 12, "bold"),
            bg="#2e7d32", fg="#ffffff", activebackground="#388e3c",
            width=22, height=2, bd=0,
        )
        self.approve_btn.grid(row=0, column=0, padx=10)

        self.deny_btn = tk.Button(
            button_frame, text="✘ Deny (N / Backspace)",
            command=self.deny, font=("Helvetica", 12, "bold"),
            bg="#c62828", fg="#ffffff", activebackground="#d32f2f",
            width=22, height=2, bd=0,
        )
        self.deny_btn.grid(row=0, column=1, padx=10)

        self.quit_btn = tk.Button(
            button_frame, text="Save & Quit (Esc)",
            command=self.quit, font=("Helvetica", 11),
            bg="#424242", fg="#ffffff", activebackground="#616161",
            width=18, height=2, bd=0,
        )
        self.quit_btn.grid(row=0, column=2, padx=10)

    def _bind_keys(self) -> None:
        self.root.bind("<y>", lambda e: self.approve())
        self.root.bind("<Y>", lambda e: self.approve())
        self.root.bind("<Return>", lambda e: self.approve())
        self.root.bind("<n>", lambda e: self.deny())
        self.root.bind("<N>", lambda e: self.deny())
        self.root.bind("<BackSpace>", lambda e: self.deny())
        self.root.bind("<Escape>", lambda e: self.quit())

    # -- control buttons toggle -------------------------------------------- #
    def _set_buttons_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.approve_btn.config(state=state)
        self.deny_btn.config(state=state)

    # -- main loop ---------------------------------------------------------- #
    def show_current(self) -> None:
        """Find the next pending album, fetch its artwork, and display it."""
        # Advance past anything already completed (e.g. resumed session).
        while self.index < len(self.albums):
            record = self.albums[self.index]
            if self.store.is_done(record.artist, record.album):
                self.index += 1
                continue
            break

        if self.index >= len(self.albums):
            self._finish()
            return

        record = self.albums[self.index]
        self.title_label.config(text=record.album)
        self.subtitle_label.config(text=f"by {record.artist}")
        self.progress_label.config(
            text=f"Album {self.index + 1} of {len(self.albums)}"
        )
        self.resolution_label.config(text="Searching Apple Music…")
        self.image_label.config(image="", text="Loading…",
                                fg="#888888", font=("Helvetica", 14))
        self._set_buttons_enabled(False)
        self._current_bytes = None
        self._current_url = None

        # Defer the (blocking) network work so the UI can paint "Loading…"
        # first. For ~250 sequential lookups this keeps things simple while
        # still feeling responsive.
        self.root.after(50, self._fetch_current)

    def _fetch_current(self) -> None:
        record = self.albums[self.index]
        try:
            url = lookup_album_artwork(record.artist, record.album)
        except requests.RequestException as exc:
            self._handle_error(record, f"Network error: {exc}")
            return

        if not url:
            # No artwork / album not found — log it and auto-skip.
            record.status = "not_found"
            record.note = "No matching album/artwork found on Apple Music."
            self.store.update(record)
            self.resolution_label.config(
                text="⚠ Not found on Apple Music — skipping…", fg="#e0a030")
            self.image_label.config(image="", text="🚫", font=("Helvetica", 48))
            self.root.after(900, self._advance)
            return

        try:
            data = download_image_bytes(url)
        except requests.RequestException as exc:
            self._handle_error(record, f"Image download failed: {exc}")
            return

        self._current_bytes = data
        self._current_url = url
        self._render_preview(data)

    def _render_preview(self, data: bytes) -> None:
        """Scale the full-res image down for on-screen viewing only."""
        try:
            image = Image.open(io.BytesIO(data))
            image.load()
        except Exception as exc:  # noqa: BLE001 - PIL raises many types
            self._handle_error(self.albums[self.index],
                               f"Could not decode image: {exc}")
            return

        full_w, full_h = image.size
        preview = image.copy()
        preview.thumbnail((MAX_PREVIEW_SIZE, MAX_PREVIEW_SIZE), Image.LANCZOS)
        self._tk_image = ImageTk.PhotoImage(preview)
        self.image_label.config(image=self._tk_image, text="")
        self.resolution_label.config(
            text=f"Full resolution: {full_w} × {full_h} px", fg="#66cc99")
        self._set_buttons_enabled(True)
        self.approve_btn.focus_set()

    def _handle_error(self, record: AlbumRecord, message: str) -> None:
        """Record a (retryable) error and let the user decide what to do."""
        record.status = "error"
        record.note = message
        self.store.update(record)
        self.resolution_label.config(text=f"⚠ {message}", fg="#e0a030")
        self.image_label.config(image="", text="⚠", font=("Helvetica", 48),
                                fg="#e0a030")
        # Offer retry vs skip rather than silently dropping the album.
        retry = messagebox.askretrycancel(
            "Error",
            f"{record.artist} — {record.album}\n\n{message}\n\n"
            "Retry this album? (Cancel skips to the next one.)",
        )
        if retry:
            self.show_current()
        else:
            self._advance()

    # -- user actions ------------------------------------------------------- #
    def approve(self) -> None:
        if self._current_bytes is None:
            return  # nothing loaded yet; ignore stray keypress
        record = self.albums[self.index]
        filename = safe_filename(record.artist, record.album)
        path = os.path.join(OUTPUT_FOLDER, filename)
        try:
            with open(path, "wb") as fh:
                fh.write(self._current_bytes)
        except OSError as exc:
            messagebox.showerror("Save failed", f"Could not save image:\n{exc}")
            return
        record.status = "approved"
        record.saved_path = path
        record.note = self._current_url
        self.store.update(record)
        self._advance()

    def deny(self) -> None:
        if self._current_bytes is None and \
                self.store.get_status(*self._current_key()) is None:
            # Allow denying even mid-load, but ignore truly empty states.
            pass
        record = self.albums[self.index]
        record.status = "denied"
        record.note = "Skipped by user."
        self.store.update(record)
        self._advance()

    def _current_key(self) -> tuple[str, str]:
        record = self.albums[self.index]
        return record.artist, record.album

    def _advance(self) -> None:
        self.index += 1
        self.show_current()

    # -- lifecycle ---------------------------------------------------------- #
    def _finish(self) -> None:
        approved = sum(1 for a in self.albums
                       if self.store.is_done(a.artist, a.album)
                       and (self.store.get_status(a.artist, a.album) or {}).get("status") == "approved")
        messagebox.showinfo(
            "All done!",
            f"You've reviewed every album.\n\n"
            f"Approved images are in: {os.path.abspath(OUTPUT_FOLDER)}\n"
            f"Approved this far: {approved}",
        )
        self.root.destroy()

    def quit(self) -> None:
        # Progress is already saved after each decision, so we can just close.
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


# --------------------------------------------------------------------------- #
# Entry point                                                                 #
# --------------------------------------------------------------------------- #

def load_albums(excel_path: str) -> list[AlbumRecord]:
    """Read the spreadsheet into a list of AlbumRecords."""
    if not os.path.exists(excel_path):
        sys.exit(f"Input file not found: {excel_path}\n"
                 f"Create an Excel file with '{ARTIST_COLUMN}' and "
                 f"'{ALBUM_COLUMN}' columns, or edit EXCEL_FILE in the script.")

    # read_only + data_only keeps memory low and gives us computed values, not
    # formulas. We only need the first worksheet.
    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    sheet = workbook.active

    rows = sheet.iter_rows(values_only=True)
    try:
        header = [("" if c is None else str(c).strip()) for c in next(rows)]
    except StopIteration:
        workbook.close()
        sys.exit("Spreadsheet appears to be empty.")

    missing = [c for c in (ARTIST_COLUMN, ALBUM_COLUMN) if c not in header]
    if missing:
        workbook.close()
        sys.exit(f"Spreadsheet is missing required column(s): {missing}\n"
                 f"Found columns: {header}")

    artist_idx = header.index(ARTIST_COLUMN)
    album_idx = header.index(ALBUM_COLUMN)

    albums: list[AlbumRecord] = []
    for row in rows:
        if row is None:
            continue
        artist = "" if artist_idx >= len(row) or row[artist_idx] is None \
            else str(row[artist_idx]).strip()
        album = "" if album_idx >= len(row) or row[album_idx] is None \
            else str(row[album_idx]).strip()
        if not artist or not album:
            continue  # skip blank rows
        albums.append(AlbumRecord(artist=artist, album=album))
    workbook.close()
    return albums


def main() -> None:
    albums = load_albums(EXCEL_FILE)
    if not albums:
        sys.exit("No valid album rows found in the spreadsheet.")

    store = ProgressStore(PROGRESS_FILE)

    remaining = [a for a in albums if not store.is_done(a.artist, a.album)]
    print(f"Loaded {len(albums)} albums "
          f"({len(albums) - len(remaining)} already reviewed, "
          f"{len(remaining)} to go).")

    if not remaining:
        print("Everything has already been reviewed. Delete "
              f"{PROGRESS_FILE} to start over.")
        return

    app = ReviewApp(albums, store)
    app.run()


if __name__ == "__main__":
    main()
