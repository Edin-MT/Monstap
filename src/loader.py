"""
loader.py
---------
Učitavanje podataka iz lokalne Excel "baze" - jednog Excel fajla ili svih
Excel fajlova u nekom folderu (npr. mesečni izveštaji koji se spajaju u
jednu tabelu).
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional

import pandas as pd

EXCEL_EXTENSIONS = (".xlsx", ".xlsm", ".xls")


class LoaderError(Exception):
    """Podignuto kada ulazni podaci ne mogu da se pronađu, učitaju ili su nevalidni."""


def load_data(input_cfg: dict) -> pd.DataFrame:
    """Učitava podatke prema 'input' sekciji konfiguracije.

    Podržano:
      - 'path' kao putanja do jednog Excel fajla ili do foldera
        (svi Excel fajlovi u folderu se učitavaju i spajaju - concat po redovima),
      - izbor sheet-a po imenu preko 'sheet' (podrazumevano: prvi sheet),
      - opciono ograničavanje na listu kolona preko 'columns' (validira se
        da sve tražene kolone zaista postoje),
      - opciono dodavanje kolone '_source_file' preko 'tag_source_file: true',
        korisno kada se spaja više fajlova pa treba znati odakle je koji red.

    Args:
        input_cfg: rečnik iz konfiguracije, npr.
            {"path": "data/input.xlsx", "sheet": "Transactions", "columns": [...]}

    Returns:
        DataFrame sa učitanim (i eventualno spojenim) podacima.

    Raises:
        LoaderError: ako putanja ne postoji, sheet/kolone ne postoje, itd.
    """
    path = input_cfg.get("path")
    if not path:
        raise LoaderError("'input.path' nije definisan u konfiguraciji.")

    sheet_name = input_cfg.get("sheet", 0)
    columns: Optional[List[str]] = input_cfg.get("columns") or None
    tag_source_file: bool = bool(input_cfg.get("tag_source_file", False))

    files = _resolve_input_files(path)

    frames = [
        _load_single_file(file_path, sheet_name, columns, tag_source_file)
        for file_path in files
    ]

    if len(frames) == 1:
        return frames[0]

    return pd.concat(frames, ignore_index=True)


def _resolve_input_files(input_path: str) -> List[Path]:
    """Vraća listu Excel fajlova za datu 'input.path' vrednost.

    'input.path' može biti:
      - putanja do jednog Excel fajla, ili
      - putanja do foldera (učitavaju se svi .xlsx/.xlsm/.xls fajlovi u njemu).
    """
    path = Path(input_path)

    if path.is_dir():
        files = sorted(
            p for p in path.iterdir() if p.is_file() and p.suffix.lower() in EXCEL_EXTENSIONS
        )
        if not files:
            raise LoaderError(f"Nijedan Excel fajl nije pronađen u folderu: {input_path}")
        return files

    if path.is_file():
        if path.suffix.lower() not in EXCEL_EXTENSIONS:
            raise LoaderError(f"Fajl nije podržanog Excel formata (.xlsx/.xlsm/.xls): {input_path}")
        return [path]

    raise LoaderError(f"Ulazna putanja ne postoji: {input_path}")


def _load_single_file(
    file_path: Path,
    sheet_name,
    columns: Optional[List[str]],
    tag_source_file: bool,
) -> pd.DataFrame:
    """Učitava jedan sheet iz jednog Excel fajla i validira tražene kolone."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    except ValueError as exc:
        raise LoaderError(f"Sheet '{sheet_name}' ne postoji u fajlu '{file_path}': {exc}") from exc
    except Exception as exc:  # pragma: no cover - genericna zastita za neocekivane greske
        raise LoaderError(f"Greška pri čitanju fajla '{file_path}': {exc}") from exc

    if isinstance(df, dict):
        # Ako je sheet_name bio lista ili None, pandas vraća {ime_sheeta: DataFrame}.
        # Uzimamo prvi sheet - konkretan slučaj treba rešiti eksplicitnim 'sheet' u configu.
        df = next(iter(df.values()))

    if columns:
        missing = [c for c in columns if c not in df.columns]
        if missing:
            raise LoaderError(
                f"Fajl '{file_path}' nema tražene kolone {missing}. "
                f"Dostupne kolone: {list(df.columns)}"
            )
        df = df[columns].copy()
    else:
        df = df.copy()

    if tag_source_file:
        df["_source_file"] = file_path.name

    return df
