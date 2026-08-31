"""
transformer.py
--------------
Filtriranje, transformacija (nove/izmenjene kolone) i agregacija
(grupisanje - "pivot-like" operacije) podataka učitanih iz Excel-a.
"""

from __future__ import annotations

from typing import Any, Dict, List

import pandas as pd


class TransformError(Exception):
    """Podignuto kada filter, transformacija ili agregacija ne mogu da se primene."""


# ---------------------------------------------------------------------------
# Filtriranje
# ---------------------------------------------------------------------------

def apply_filters(df: pd.DataFrame, filters_cfg: Dict[str, Any]) -> pd.DataFrame:
    """Primenjuje filtere definisane u konfiguraciji nad DataFrame-om.

    Podržani ključevi u sekciji 'filters':
      date_column: ime kolone sa datumom (koristi se uz date_from/date_to)
      date_from:   "YYYY-MM-DD" - donja granica, uključivo
      date_to:     "YYYY-MM-DD" - gornja granica, uključivo
      in:          {kolona: [dozvoljena_vrednost, ...]}  npr. filter po kategoriji/regionu
      equals:      {kolona: vrednost}                     tačno poklapanje
      min:         {kolona: broj}                         kolona >= broj
      max:         {kolona: broj}                         kolona <= broj

    Sve navedene provere se kombinuju sa logičkim I (AND).

    Args:
        df: ulazni DataFrame.
        filters_cfg: 'filters' sekcija konfiguracije (može biti prazna/None).

    Returns:
        Novi, filtrirani DataFrame (indeks je resetovan).
    """
    if not filters_cfg:
        return df

    result = df

    date_column = filters_cfg.get("date_column")
    date_from = filters_cfg.get("date_from")
    date_to = filters_cfg.get("date_to")
    if date_from or date_to:
        if not date_column:
            raise TransformError("Za 'date_from'/'date_to' mora biti podešeno 'filters.date_column'.")
        _require_column(result, date_column, "filters.date_column")
        result = result.copy()
        result[date_column] = pd.to_datetime(result[date_column], errors="coerce")
        if date_from:
            result = result[result[date_column] >= pd.to_datetime(date_from)]
        if date_to:
            result = result[result[date_column] <= pd.to_datetime(date_to)]

    for column, allowed_values in (filters_cfg.get("in") or {}).items():
        _require_column(result, column, "filters.in")
        result = result[result[column].isin(allowed_values)]

    for column, value in (filters_cfg.get("equals") or {}).items():
        _require_column(result, column, "filters.equals")
        result = result[result[column] == value]

    for column, min_value in (filters_cfg.get("min") or {}).items():
        _require_column(result, column, "filters.min")
        result = result[result[column] >= min_value]

    for column, max_value in (filters_cfg.get("max") or {}).items():
        _require_column(result, column, "filters.max")
        result = result[result[column] <= max_value]

    return result.reset_index(drop=True)


# ---------------------------------------------------------------------------
# Transformacije (nove kolone, preimenovanje, zaokruživanje, brisanje kolona)
# ---------------------------------------------------------------------------

def apply_transformations(df: pd.DataFrame, transformations_cfg: Dict[str, Any]) -> pd.DataFrame:
    """Primenjuje transformacije definisane u konfiguraciji.

    Podržani ključevi u sekciji 'transformations':
      compute_columns: [{name, expression}, ...]
          Nova (ili izmenjena) kolona = rezultat izraza nad postojećim kolonama,
          npr. {"name": "total", "expression": "quantity * price"}.
          Izraz se evaluira preko pandas 'DataFrame.eval', koji radi samo nad
          postojećim kolonama - nema pristupa proizvoljnom Python kodu/modulima.
      round:          {kolona: broj_decimala}
      rename_columns: {staro_ime: novo_ime}
      drop_columns:   [kolona, ...]

    Redosled primene je fiksan: compute_columns -> round -> rename_columns -> drop_columns.

    Args:
        df: ulazni (obično već filtrirani) DataFrame.
        transformations_cfg: 'transformations' sekcija konfiguracije.

    Returns:
        Novi DataFrame sa primenjenim transformacijama.
    """
    if not transformations_cfg:
        return df

    result = df.copy()

    for spec in transformations_cfg.get("compute_columns", []):
        name = spec.get("name")
        expression = spec.get("expression")
        if not name or not expression:
            raise TransformError(f"Nevalidna 'compute_columns' stavka (potrebno 'name' i 'expression'): {spec}")
        try:
            result[name] = result.eval(expression, engine="python")
        except Exception as exc:
            raise TransformError(
                f"Greška pri izračunavanju kolone '{name}' iz izraza '{expression}': {exc}"
            ) from exc

    for column, decimals in (transformations_cfg.get("round") or {}).items():
        _require_column(result, column, "transformations.round")
        result[column] = result[column].round(decimals)

    rename_map = transformations_cfg.get("rename_columns") or {}
    if rename_map:
        for column in rename_map:
            _require_column(result, column, "transformations.rename_columns")
        result = result.rename(columns=rename_map)

    for column in transformations_cfg.get("drop_columns", []):
        _require_column(result, column, "transformations.drop_columns")
        result = result.drop(columns=[column])

    return result


# ---------------------------------------------------------------------------
# Agregacija / "Summary" sheet (grupisanje - pivot-like operacija)
# ---------------------------------------------------------------------------

def build_summary(df: pd.DataFrame, summary_cfg: Dict[str, Any]) -> pd.DataFrame:
    """Pravi agregiranu (grupisanu) tabelu za rezimirajući "Summary" sheet.

    Konfiguracija ('summary'):
      group_by:     [kolona, ...] - kolone po kojima se grupiše
      aggregations: {kolona: "sum"|"mean"|"count"|"min"|"max"|...} - agregacione funkcije

    Args:
        df: DataFrame nad kojim se radi grupisanje (obično već transformisan).
        summary_cfg: 'summary' sekcija konfiguracije.

    Returns:
        Agregirani DataFrame sa kolonama za grupisanje + agregirane kolone.
    """
    group_by: List[str] = summary_cfg.get("group_by") or []
    aggregations: Dict[str, str] = summary_cfg.get("aggregations") or {}

    for column in group_by:
        _require_column(df, column, "summary.group_by")
    for column in aggregations:
        _require_column(df, column, "summary.aggregations")

    grouped = df.groupby(group_by, dropna=False).agg(aggregations)
    return grouped.reset_index()


def _require_column(df: pd.DataFrame, column: str, context: str) -> None:
    """Baca TransformError ako 'column' ne postoji u DataFrame-u, sa jasnom porukom."""
    if column not in df.columns:
        raise TransformError(
            f"Kolona '{column}' (iz '{context}') ne postoji. Dostupne kolone: {list(df.columns)}"
        )
