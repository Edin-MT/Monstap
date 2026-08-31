"""
generate_sample_data.py
------------------------
Pomoćna skripta koja generiše primer ulaznog fajla `data/input.xlsx` sa
sheet-om "Transactions", korišćen u primeru konfiguracije
(config/export_config.yaml).

Pokretanje (iz korena projekta):
    python scripts/generate_sample_data.py
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

OUTPUT_PATH = Path(__file__).resolve().parent.parent / "data" / "input.xlsx"

CATEGORIES = ["Agriculture", "Industry", "Services"]
REGIONS = ["Vojvodina", "Šumadija", "Beograd"]

ROWS = [
    # date, category, region, quantity, price, value_rsd
    ("2025-01-15", "Agriculture", "Vojvodina", 120, 850.0, 102000.0),
    ("2025-02-10", "Agriculture", "Vojvodina", 80, 900.0, 72000.0),
    ("2025-03-05", "Agriculture", "Šumadija", 200, 780.0, 156000.0),
    ("2025-04-20", "Agriculture", "Beograd", 60, 950.0, 57000.0),
    ("2025-05-11", "Industry", "Beograd", 30, 5200.0, 156000.0),
    ("2025-06-01", "Services", "Beograd", 15, 3000.0, 45000.0),
    ("2025-07-18", "Agriculture", "Šumadija", 150, 800.0, 120000.0),
    ("2025-08-09", "Agriculture", "Vojvodina", 100, 870.0, 87000.0),
    ("2025-09-22", "Industry", "Vojvodina", 40, 4800.0, 192000.0),
    ("2025-10-03", "Agriculture", "Beograd", 90, 910.0, 81900.0),
    ("2025-11-14", "Services", "Šumadija", 25, 2600.0, 65000.0),
    ("2025-12-30", "Agriculture", "Vojvodina", 110, 860.0, 94600.0),
    # Zapisi van 2025. godine i van kategorije Agriculture - koriste se da
    # pokažu da filteri iz config/export_config.yaml rade ispravno.
    ("2024-06-10", "Agriculture", "Vojvodina", 70, 800.0, 56000.0),
    ("2026-01-05", "Agriculture", "Vojvodina", 70, 800.0, 56000.0),
    ("2025-03-12", "Industry", "Šumadija", 20, 6000.0, 120000.0),
]


def generate() -> Path:
    """Kreira DataFrame sa primer podacima i upisuje ga u data/input.xlsx."""
    df = pd.DataFrame(
        ROWS,
        columns=["date", "category", "region", "quantity", "price", "value_rsd"],
    )
    df["date"] = pd.to_datetime(df["date"])

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)

    return OUTPUT_PATH


if __name__ == "__main__":
    path = generate()
    print(f"Primer podataka je generisan: {path}")
