#!/usr/bin/env python3
"""
make_sample_albums.py
=====================

Convenience helper: writes a small example ``albums.xlsx`` with the expected
``Artist`` / ``Album`` columns so you can try the reviewer immediately.
Replace it with your own ~250-row spreadsheet using the same headers.
"""

import pandas as pd

SAMPLE = [
    ("Fleetwood Mac", "Rumours"),
    ("Pink Floyd", "The Dark Side of the Moon"),
    ("Daft Punk", "Random Access Memories"),
    ("Kendrick Lamar", "To Pimp a Butterfly"),
    ("Radiohead", "OK Computer"),
    ("Miles Davis", "Kind of Blue"),
    ("Tame Impala", "Currents"),
    ("Adele", "21"),
]

if __name__ == "__main__":
    df = pd.DataFrame(SAMPLE, columns=["Artist", "Album"])
    df.to_excel("albums.xlsx", index=False, engine="openpyxl")
    print(f"Wrote albums.xlsx with {len(df)} sample rows.")
