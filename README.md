# Apple Music High-Res Artwork Fetcher

A small desktop tool that pulls **high-resolution album artwork** from Apple
Music (via the public iTunes Search API) for a list of albums in an Excel
sheet, lets you **Approve / Deny** each cover in a simple `tkinter` window, and
**resumes** where you left off if you stop partway through ~250 albums.

## How the high-res trick works

The iTunes Search API returns an artwork URL like:

```
https://is1-ssl.mzstatic.com/.../source/100x100bb.jpg
```

The `100x100bb` (or `600x600bb`) segment is just a *resize instruction* for
Apple's CDN. The script rewrites the final `<n>x<n>bb` token to
`10000x10000bb`, and the CDN clamps that request down to the **largest original
master it actually has** (often 1400×1400, 3000×3000, or larger) — perfect for
printing. This is the same approach as Ben Dodson's *Apple Music Artwork
Finder*. The preview shown on screen is downscaled for viewing only; the file
saved to disk is always the full-resolution download.

## 1. Install dependencies

```bash
pip install -r requirements.txt
# or individually:
pip install openpyxl requests Pillow
```

`tkinter` ships with most Python installs. If you get a "tkinter not available"
error on Linux:

```bash
sudo apt-get install python3-tk        # Debian / Ubuntu
```

## 2. Prepare your spreadsheet

Create `albums.xlsx` next to the script with two columns: **`Artist`** and
**`Album`** (header names matter; ~250 rows is fine). To generate a small
sample to test with:

```bash
python make_sample_albums.py
```

You can change the filename / column names / output folder near the top of
`album_artwork_fetcher.py` (the `Configuration` section).

## 3. Run the reviewer

```bash
python album_artwork_fetcher.py
```

For each album a window shows the cover, the artist/album, and the detected
full resolution. Then:

| Action            | Buttons / Keys           |
|-------------------|--------------------------|
| Approve & download | `Y` or `Enter`          |
| Deny / skip        | `N` or `Backspace`      |
| Save & quit        | `Esc`                   |

Approved images are written to **`HighRes_Covers/`** as `Artist - Album.jpg`.

## 4. Resuming & state

Every decision is saved immediately to **`progress.json`**. Close the window
any time (`Esc` or the window's X) and re-run — it skips albums already marked
`approved`, `denied`, or `not_found` and continues from the next pending one.
Transient `error` rows are retried on the next run.

To start completely over, delete `progress.json` (and optionally
`HighRes_Covers/`).

## Error handling

- **Album not found / no artwork:** logged as `not_found` and auto-skipped
  after a brief on-screen notice.
- **Network timeout / download failure:** shows a Retry / Skip dialog; the row
  is marked `error` so it's retried next run if you skip.
- **Undecodable image:** treated like a download error (retry / skip).

## Files

| File                        | Purpose                                   |
|-----------------------------|-------------------------------------------|
| `album_artwork_fetcher.py`  | Main script (lookup + UI + state).        |
| `make_sample_albums.py`     | Generates a sample `albums.xlsx`.         |
| `requirements.txt`          | Python dependencies.                      |
| `progress.json`             | Auto-created review progress (resume).    |
| `HighRes_Covers/`           | Auto-created output folder.               |
