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

import argparse
import io
import json
import os
import re
import sys
import time
import unicodedata
from dataclasses import dataclass, asdict
from typing import Optional

import requests

try:
    import openpyxl
except ImportError:  # pragma: no cover - friendly guidance
    sys.exit("Missing dependency 'openpyxl'. Install with: pip install openpyxl")

try:
    import tkinter as tk
    from tkinter import messagebox, simpledialog
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

EXCEL_FILE = "album.xlsx"          # input spreadsheet (columns: Artist, Album)
ARTIST_COLUMN = "Artist"            # column header for the artist name
ALBUM_COLUMN = "Album"              # column header for the album title
OUTPUT_FOLDER = "HighRes_Covers"    # where approved covers are saved
PROGRESS_FILE = "progress.json"     # tracks per-row status so we can resume

ITUNES_ENDPOINT = "https://itunes.apple.com/search"
REQUEST_TIMEOUT = 15                # seconds for any single network request
MAX_PREVIEW_SIZE = 700             # max width/height (px) for the on-screen preview
TARGET_RESOLUTION = "10000x10000bb"  # request size; CDN downscales to the real max

# Some networks/CDNs reject the default "python-requests/x" User-Agent. A shared
# Session with a browser-like UA improves reliability and reuses the connection.
SESSION = requests.Session()
SESSION.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
})

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

    def reset_statuses(self, statuses: set[str]) -> int:
        """Forget entries whose status is in ``statuses`` so they re-run.

        Returns the number of cleared entries. Used by the --retry-* flags to
        re-attempt only the previously missing/denied albums with the improved
        matching logic, without touching anything already approved.
        """
        to_clear = [k for k, v in self._data.items()
                    if v.get("status") in statuses]
        for k in to_clear:
            del self._data[k]
        if to_clear:
            self.save()
        return len(to_clear)

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


# Light normalization used only to compare a candidate title against what we
# asked for. We deliberately keep matching STRICT (exact normalized title, or
# clean containment) so we never accept an unrelated album.
_BRACKET_RE = re.compile(r"[\(\[\{].*?[\)\]\}]")
_ARTIST_SPLIT_RE = re.compile(
    r"\s*(?:,|&|\bfeat\.?\b|\bfeaturing\b|\bwith\b|\bx\b|/)\s*", re.IGNORECASE
)


def _normalize(text: str) -> str:
    """Lowercase + strip accents/punctuation, for conservative comparison."""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def _primary_artist(artist: str) -> str:
    """Return the lead artist, dropping collaborators/features.

    "boygenius, Julien Baker, Lucy Dacus & Phoebe Bridgers" -> "boygenius"
    """
    parts = _ARTIST_SPLIT_RE.split(artist, maxsplit=1)
    return parts[0].strip() if parts and parts[0].strip() else artist.strip()


def _itunes_get_url(url: str, params: dict) -> requests.Response:
    """GET an iTunes endpoint with retry/backoff on rate-limit (403/429)."""
    backoff = 2
    for attempt in range(4):
        resp = SESSION.get(url, params=params, timeout=REQUEST_TIMEOUT)
        if resp.status_code in (403, 429) and attempt < 3:
            time.sleep(backoff)
            backoff *= 2
            continue
        return resp
    return resp


def _itunes_get(params: dict) -> requests.Response:
    """GET the iTunes Search endpoint (convenience wrapper)."""
    return _itunes_get_url(ITUNES_ENDPOINT, params)


def lookup_album_artwork(artist: str, album: str) -> Optional[tuple[str, str]]:
    """Query the iTunes Search API for an album cover (STRICT matching).

    Uses a single "<artist> <album>" query and accepts the first result only
    when its title matches what we asked for, so we don't return unrelated art.
    Returns ``(high_res_url, matched_label)`` or ``None``. Raises
    requests.RequestException on network failure.
    """
    query = f"{artist} {album}".strip()
    params = {"term": query, "entity": "album", "limit": 5, "media": "music"}
    resp = _itunes_get(params)
    resp.raise_for_status()
    results = resp.json().get("results", [])
    if not results:
        return None

    norm_album = _normalize(album)
    best = None
    # Prefer an exact (normalized) title match.
    for item in results:
        if _normalize(item.get("collectionName", "")) == norm_album:
            best = item
            break
    # Otherwise accept the top hit only if its title clearly corresponds to
    # the album we searched for (guards against unrelated first results).
    if best is None:
        first = results[0]
        cand = _normalize(first.get("collectionName", ""))
        if cand and (cand in norm_album or norm_album in cand):
            best = first
    if best is None:
        return None

    artwork = best.get("artworkUrl100") or best.get("artworkUrl60")
    if not artwork:
        return None
    label = (f"{best.get('artistName', '?')} — "
             f"{best.get('collectionName', '?')}")
    return build_highres_url(artwork), label


# --------------------------------------------------------------------------- #
# YouTube Music lookup (optional source, requires `ytmusicapi`)               #
# --------------------------------------------------------------------------- #

# YouTube/Google art is served from googleusercontent with a sizing token like
# "=w544-h544-l90-rj". Swapping that for a large size makes the CDN return the
# largest master it has (it will not upscale beyond the original).
_YT_SIZE_RE = re.compile(r"=w\d+-h\d+(?:-[a-z0-9]+)*$")
YT_TARGET_SIZE = 3000

_YT_CLIENT = None


def _youtube_highres_url(url: str, size: int = YT_TARGET_SIZE) -> str:
    """Rewrite a googleusercontent thumbnail URL to request a large size."""
    if _YT_SIZE_RE.search(url):
        return _YT_SIZE_RE.sub(f"=w{size}-h{size}", url)
    sep = "" if url.endswith("=") else "="
    return f"{url}{sep}w{size}-h{size}"


def _get_ytmusic():
    """Lazily create an (unauthenticated) YTMusic client; search needs no auth."""
    global _YT_CLIENT
    if _YT_CLIENT is None:
        try:
            from ytmusicapi import YTMusic
        except ImportError:  # pragma: no cover - friendly guidance
            raise RuntimeError(
                "The --source youtube option needs the 'ytmusicapi' package.\n"
                "Install it with:  pip install ytmusicapi")
        _YT_CLIENT = YTMusic()
    return _YT_CLIENT


def lookup_album_artwork_youtube(artist: str, album: str) -> Optional[tuple[str, str]]:
    """Find album artwork on YouTube Music (STRICT matching).

    Returns ``(high_res_url, matched_label)`` or ``None``. Raises
    requests.RequestException on lookup failure so the caller treats it as a
    retryable error rather than a genuine "not found".
    """
    yt = _get_ytmusic()
    try:
        results = yt.search(f"{artist} {album}", filter="albums", limit=10)
    except Exception as exc:  # ytmusicapi wraps network errors in various types
        raise requests.RequestException(f"YouTube Music search failed: {exc}")
    if not results:
        return None

    norm_album = _normalize(album)
    norm_artist = _normalize(artist)
    norm_primary = _normalize(_primary_artist(artist))

    def artist_ok(item) -> bool:
        cand = _normalize(" ".join(
            a.get("name", "") for a in item.get("artists", [])
            if isinstance(a, dict)))
        if not cand:
            return True  # some album entries omit artist; title match must carry
        return (cand in norm_artist or norm_artist in cand
                or norm_primary in cand or cand in norm_primary)

    best = None
    for item in results:
        cand_title = _normalize(item.get("title", ""))
        if not cand_title:
            continue
        title_match = (cand_title == norm_album
                       or cand_title in norm_album or norm_album in cand_title)
        if title_match and artist_ok(item):
            best = item
            break
    if best is None:
        return None

    thumbs = best.get("thumbnails", [])
    if not thumbs:
        return None
    url = thumbs[-1].get("url", "")  # last thumbnail is the largest
    if not url:
        return None
    artists = ", ".join(a.get("name", "") for a in best.get("artists", [])
                        if isinstance(a, dict))
    label = f"{artists or '?'} — {best.get('title', '?')}  [YT Music]"
    return _youtube_highres_url(url), label


def lookup_artwork(artist: str, album: str, source: str) -> Optional[tuple[str, str]]:
    """Dispatch to the selected artwork source ('itunes' or 'youtube')."""
    if source == "youtube":
        return lookup_album_artwork_youtube(artist, album)
    return lookup_album_artwork(artist, album)


# --------------------------------------------------------------------------- #
# Manual lookup by Apple Music URL / ID                                       #
# --------------------------------------------------------------------------- #
#
# The legacy iTunes *Search* API has coverage gaps: some albums that exist on
# Apple Music (often newer releases) simply aren't returned by a text search.
# The iTunes *Lookup* API resolves an album by its numeric ID reliably, so this
# lets you paste an Apple Music URL to rescue any album the search can't find.

ITUNES_LOOKUP = "https://itunes.apple.com/lookup"
# An Apple Music album URL ends in the collection id, e.g.
# .../album/thats-my-cue-a-solo-experience/1756160509   (an optional ?i=<trackid>
# is a *track* id, which we ignore — we want the album/collection id).
_APPLE_ID_RE = re.compile(r"/(?:album|id)/[^/]*?/?(\d{6,})", re.IGNORECASE)
_BARE_ID_RE = re.compile(r"(?<!\d)(\d{6,})(?!\d)")


def extract_apple_album_id(text: str) -> Optional[str]:
    """Pull the album/collection id out of an Apple Music URL or raw id string."""
    text = text.strip()
    # Prefer the path id; fall back to any long run of digits (raw id paste).
    m = _APPLE_ID_RE.search(text)
    if m:
        return m.group(1)
    # Avoid grabbing a ?i= track id if a clean numeric id wasn't in the path.
    no_track = re.sub(r"[?&]i=\d+", "", text)
    m = _BARE_ID_RE.search(no_track)
    return m.group(1) if m else None


def lookup_album_artwork_by_id(album_id: str) -> Optional[tuple[str, str]]:
    """Resolve artwork for an album by its iTunes/Apple Music collection id."""
    params = {"id": album_id, "entity": "album"}
    resp = _itunes_get_url(ITUNES_LOOKUP, params)
    resp.raise_for_status()
    for item in resp.json().get("results", []):
        # The collection entry (not the individual tracks) carries the artwork.
        if item.get("wrapperType") == "collection" or item.get("collectionName"):
            artwork = item.get("artworkUrl100") or item.get("artworkUrl60")
            if artwork:
                label = (f"{item.get('artistName', '?')} — "
                         f"{item.get('collectionName', '?')}")
                return build_highres_url(artwork), label
    return None


def download_image_bytes(url: str) -> bytes:
    """Download raw image bytes (used both for preview and for saving)."""
    resp = SESSION.get(url, timeout=REQUEST_TIMEOUT)
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

    def __init__(self, albums: list[AlbumRecord], store: ProgressStore,
                 source: str = "itunes"):
        self.albums = albums
        self.store = store
        self.source = source
        self.index = 0
        # Full-resolution bytes of the currently displayed image (for saving).
        self._current_bytes: Optional[bytes] = None
        self._current_url: Optional[str] = None
        self._current_label: Optional[str] = None
        self._awaiting_manual = False  # True while paused on a not-found album
        self._tk_image: Optional[ImageTk.PhotoImage] = None  # keep a reference!

        os.makedirs(OUTPUT_FOLDER, exist_ok=True)

        self.source_name = "YouTube Music" if source == "youtube" else "Apple Music"
        self.root = tk.Tk()
        self.root.title(f"{self.source_name} High-Res Artwork Reviewer")
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

        # Manual rescue: paste an Apple Music URL to fetch art by album id when
        # the automatic search can't find it.
        self.url_btn = tk.Button(
            button_frame, text="🔗 Paste Apple URL (U)",
            command=self.paste_url, font=("Helvetica", 11),
            bg="#1565c0", fg="#ffffff", activebackground="#1976d2",
            width=20, height=2, bd=0,
        )
        self.url_btn.grid(row=0, column=2, padx=10)

        self.quit_btn = tk.Button(
            button_frame, text="Save & Quit (Esc)",
            command=self.quit, font=("Helvetica", 11),
            bg="#424242", fg="#ffffff", activebackground="#616161",
            width=18, height=2, bd=0,
        )
        self.quit_btn.grid(row=0, column=3, padx=10)

    def _bind_keys(self) -> None:
        self.root.bind("<y>", lambda e: self.approve())
        self.root.bind("<Y>", lambda e: self.approve())
        self.root.bind("<Return>", lambda e: self.approve())
        self.root.bind("<n>", lambda e: self.deny())
        self.root.bind("<N>", lambda e: self.deny())
        self.root.bind("<BackSpace>", lambda e: self.deny())
        self.root.bind("<u>", lambda e: self.paste_url())
        self.root.bind("<U>", lambda e: self.paste_url())
        self.root.bind("<Escape>", lambda e: self.quit())

    # -- control buttons toggle -------------------------------------------- #
    def _set_buttons_enabled(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        self.approve_btn.config(state=state)
        self.deny_btn.config(state=state)
        # The manual-URL rescue is available whenever we're not actively loading.
        self.url_btn.config(state=state)

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
        self.resolution_label.config(text=f"Searching {self.source_name}…")
        self.image_label.config(image="", text="Loading…",
                                fg="#888888", font=("Helvetica", 14))
        self._set_buttons_enabled(False)
        self._current_bytes = None
        self._current_url = None
        self._current_label = None
        self._awaiting_manual = False

        # Defer the (blocking) network work so the UI can paint "Loading…"
        # first. For ~250 sequential lookups this keeps things simple while
        # still feeling responsive.
        self.root.after(50, self._fetch_current)

    def _fetch_current(self) -> None:
        record = self.albums[self.index]
        try:
            match = lookup_artwork(record.artist, record.album, self.source)
        except requests.RequestException as exc:
            self._handle_error(record, f"Lookup error: {exc}")
            return
        except RuntimeError as exc:
            # e.g. ytmusicapi not installed — fatal, tell the user and stop.
            messagebox.showerror("YouTube Music unavailable", str(exc))
            self.root.destroy()
            return

        if not match:
            # Not auto-found. Pause so the user can paste an Apple Music URL to
            # fetch it by id (the Search API misses some albums), or skip it.
            self._awaiting_manual = True
            self.resolution_label.config(
                text=(f"⚠ Not found on {self.source_name}.  Press U to paste an "
                      f"Apple Music URL, or N to skip."), fg="#e0a030")
            self.image_label.config(image="", text="🔗", font=("Helvetica", 48),
                                    fg="#e0a030")
            self.deny_btn.config(state=tk.NORMAL)
            self.url_btn.config(state=tk.NORMAL)
            self.url_btn.focus_set()
            return

        url, matched_label = match
        # Show what was actually matched so a wrong hit is easy to spot/deny.
        self.subtitle_label.config(text=f"by {record.artist}   ·   matched: {matched_label}")

        try:
            data = download_image_bytes(url)
        except requests.RequestException as exc:
            self._handle_error(record, f"Image download failed: {exc}")
            return

        self._current_bytes = data
        self._current_url = url
        self._current_label = matched_label
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
        record = self.albums[self.index]
        if self._awaiting_manual:
            # Skipping an album we couldn't auto-find: keep it as 'not_found'
            # so a later --retry-missing pass can pick it up again.
            record.status = "not_found"
            record.note = f"Not found on {self.source_name}; skipped by user."
        else:
            record.status = "denied"
            record.note = "Skipped by user."
        self.store.update(record)
        self._advance()

    def paste_url(self) -> None:
        """Prompt for an Apple Music URL/ID and fetch artwork by album id.

        Rescues albums the automatic search can't find (the iTunes Search API
        has coverage gaps; the Lookup-by-id API resolves them reliably).
        """
        raw = simpledialog.askstring(
            "Paste Apple Music URL",
            "Paste the Apple Music album URL (or numeric album id):",
            parent=self.root)
        if not raw:
            return
        album_id = extract_apple_album_id(raw)
        if not album_id:
            messagebox.showwarning(
                "Couldn't read id",
                "I couldn't find an album id in that text.\n\n"
                "Use the album page URL, e.g.\n"
                "https://music.apple.com/us/album/<name>/1756160509")
            return

        self.resolution_label.config(
            text=f"Looking up album id {album_id}…", fg="#888888")
        self.root.update_idletasks()
        try:
            match = lookup_album_artwork_by_id(album_id)
        except requests.RequestException as exc:
            messagebox.showerror("Lookup failed", f"Could not look up id:\n{exc}")
            return
        if not match:
            messagebox.showwarning(
                "No artwork",
                f"No album artwork found for id {album_id}.")
            return

        url, matched_label = match
        try:
            data = download_image_bytes(url)
        except requests.RequestException as exc:
            messagebox.showerror("Download failed", f"Could not download:\n{exc}")
            return

        record = self.albums[self.index]
        self.subtitle_label.config(
            text=f"by {record.artist}   ·   matched: {matched_label}")
        self._awaiting_manual = False
        self._current_bytes = data
        self._current_url = url
        self._current_label = matched_label
        self._render_preview(data)

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


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Review & download high-res Apple Music album artwork.")
    parser.add_argument(
        "--excel", default=EXCEL_FILE,
        help=f"Path to the input spreadsheet (default: {EXCEL_FILE}).")
    parser.add_argument(
        "--source", choices=("itunes", "youtube"), default="itunes",
        help="Artwork source: 'itunes' (Apple Music, default) or 'youtube' "
             "(YouTube Music; requires 'pip install ytmusicapi'). Combine with "
             "--retry-missing to re-check not-found albums on YouTube Music.")
    parser.add_argument(
        "--retry-missing", action="store_true",
        help="Re-attempt albums previously marked 'not_found'. "
             "Already-approved albums are left untouched.")
    parser.add_argument(
        "--retry-denied", action="store_true",
        help="Also re-review albums you previously skipped ('denied').")
    parser.add_argument(
        "--retry-all", action="store_true",
        help="Re-review everything except already-approved albums.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    albums = load_albums(args.excel)
    if not albums:
        sys.exit("No valid album rows found in the spreadsheet.")

    store = ProgressStore(PROGRESS_FILE)

    # Build the set of statuses to forget so those albums re-run this session.
    retry: set[str] = set()
    if args.retry_missing or args.retry_all:
        retry.add("not_found")
    if args.retry_denied or args.retry_all:
        retry.add("denied")
    if retry:
        cleared = store.reset_statuses(retry)
        print(f"Re-queued {cleared} album(s) with status {sorted(retry)} "
              f"for another pass.")

    remaining = [a for a in albums if not store.is_done(a.artist, a.album)]
    print(f"Loaded {len(albums)} albums "
          f"({len(albums) - len(remaining)} already reviewed, "
          f"{len(remaining)} to go).")

    if not remaining:
        print("Everything has already been reviewed. Re-run with "
              "--retry-missing to retry not-found albums, or delete "
              f"{PROGRESS_FILE} to start over.")
        return

    app = ReviewApp(albums, store, source=args.source)
    app.run()


if __name__ == "__main__":
    main()
