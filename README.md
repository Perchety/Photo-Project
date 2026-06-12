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

| Action               | Buttons / Keys      |
|----------------------|---------------------|
| Approve & download   | `Y` or `Enter`      |
| Deny / skip          | `N` or `Backspace`  |
| Paste Apple URL      | `U`                 |
| Save & quit          | `Esc`               |

Approved images are written to **`HighRes_Covers/`** as `Artist - Album.jpg`.

### When an album isn't found automatically

The iTunes **Search** API has coverage gaps — some albums that are clearly on
Apple Music (often newer releases) just aren't returned by a text search. When
that happens the reviewer **pauses** on the album instead of skipping, and you
can press **`U`** to paste that album's Apple Music URL (e.g.
`https://music.apple.com/us/album/<name>/1756160509`). It extracts the album id,
fetches the cover via the iTunes **Lookup** API (which resolves any id
reliably), and shows it for approval at full resolution. Press `N` to skip
instead — skipped albums stay `not_found` so a later `--retry-missing` can
revisit them.

## 4. Resuming & state

Every decision is saved immediately to **`progress.json`**. Close the window
any time (`Esc` or the window's X) and re-run — it skips albums already marked
`approved`, `denied`, or `not_found` and continues from the next pending one.
Transient `error` rows are retried on the next run.

To start completely over, delete `progress.json` (and optionally
`HighRes_Covers/`).

### Retrying albums that weren't found

Matching is intentionally **strict** — a cover is only accepted when the result's
title actually matches the album you asked for, so you never get an unrelated
image. Some albums therefore come back `not_found` on Apple Music. To re-attempt
those (the status is otherwise terminal):

```bash
python album_artwork_fetcher.py --retry-missing      # retry not_found only
python album_artwork_fetcher.py --retry-denied       # also re-review skipped
python album_artwork_fetcher.py --retry-all          # re-review all but approved
```

Already-**approved** albums are never touched. You can also point at a different
spreadsheet with `--excel path/to/file.xlsx`.

### Trying YouTube Music for the leftovers

If an album isn't on Apple Music, you can re-check the missing ones against
**YouTube Music** instead. This needs the optional `ytmusicapi` package:

```bash
pip install ytmusicapi
python album_artwork_fetcher.py --source youtube --retry-missing
```

This re-queues only the `not_found` albums and searches YouTube Music for each
(no login required). The reviewer header and the "matched:" line show
`[YT Music]` so you know which source you're approving from. YouTube Music art
is fetched from Google's CDN and upscaled-on-request to the largest master
available — usually high-res, though it can be smaller than Apple's for some
titles, so eyeball the resolution shown before approving.

## Error handling

- **Album not found / no artwork:** the reviewer pauses so you can paste the
  album's Apple Music URL (`U`) or skip it (`N`, kept as `not_found`).
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
