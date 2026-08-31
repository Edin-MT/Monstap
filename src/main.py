"""
main.py
-------
CLI ulazna tačka aplikacije za obradu lokalne Excel "baze" i eksport u nove
Excel fajlove.

Primer pokretanja (iz korena projekta):
    python src/main.py --config config/export_config.yaml
    python src/main.py --config config/export_config.yaml --verbose
    python src/main.py --config config/export_config.yaml --input data/other.xlsx --output output/other.xlsx

Tok obrade:
    1. učitavanje konfiguracije (config.py),
    2. učitavanje Excel podataka (loader.py),
    3. filtriranje i transformacija podataka (transformer.py),
    4. eksport rezultata u novi Excel fajl sa formatiranjem (exporter.py).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

# Omogućava da se skripta pokreće i direktno ('python src/main.py') i kao modul,
# bez potrebe za instalacijom paketa.
sys.path.insert(0, str(Path(__file__).resolve().parent))

from config import ConfigError, load_config  # noqa: E402
from exporter import export_workbook  # noqa: E402
from loader import LoaderError, load_data  # noqa: E402
from transformer import TransformError, apply_filters, apply_transformations, build_summary  # noqa: E402


def parse_args(argv=None) -> argparse.Namespace:
    """Definiše i parsira argumente komandne linije."""
    parser = argparse.ArgumentParser(
        description="Obrada Excel podataka po konfiguraciji i eksport u novi Excel fajl.",
    )
    parser.add_argument(
        "--config",
        "-c",
        default="config/export_config.yaml",
        help="Putanja do YAML/JSON konfiguracionog fajla (podrazumevano: config/export_config.yaml).",
    )
    parser.add_argument(
        "--input",
        help="Opciono: prepisuje 'input.path' iz konfiguracije (ulazni fajl ili folder).",
    )
    parser.add_argument(
        "--output",
        help="Opciono: prepisuje 'output.path' iz konfiguracije (putanja izlaznog .xlsx fajla).",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        help="Ispisuje dodatne informacije o toku obrade (broj redova pre/posle filtera itd.).",
    )
    return parser.parse_args(argv)


def run(args: argparse.Namespace) -> int:
    """Izvršava ceo tok obrade (load -> filter -> transform -> export).

    Returns:
        0 pri uspehu, 1 pri bilo kojoj kontrolisanoj grešci (ispisanoj na stderr).
    """
    try:
        cfg = load_config(args.config)
    except ConfigError as exc:
        print(f"[GREŠKA][config] {exc}", file=sys.stderr)
        return 1

    if args.input:
        cfg["input"]["path"] = args.input
    if args.output:
        cfg["output"]["path"] = args.output

    try:
        raw_df = load_data(cfg["input"])
        if args.verbose:
            print(f"Učitano {len(raw_df)} redova iz '{cfg['input']['path']}'.")

        filtered_df = apply_filters(raw_df, cfg.get("filters", {}))
        if args.verbose:
            print(f"Nakon filtera: {len(filtered_df)} redova.")

        processed_df = apply_transformations(filtered_df, cfg.get("transformations", {}))

        output_cfg = cfg["output"]
        sheet_names = output_cfg.get("sheet_names", {})
        sheets: dict = {}

        if output_cfg.get("include_raw", True):
            sheets[sheet_names.get("raw", "Raw")] = raw_df

        if output_cfg.get("include_processed", True):
            sheets[sheet_names.get("processed", "Processed")] = processed_df

        summary_cfg = cfg.get("summary", {})
        if summary_cfg.get("enabled"):
            summary_df = build_summary(processed_df, summary_cfg)
            sheets[sheet_names.get("summary", "Summary")] = summary_df
            if args.verbose:
                print(f"Summary: {len(summary_df)} grupisanih redova.")

        export_workbook(output_cfg["path"], sheets, output_cfg)
        print(f"Gotovo. Izlazni fajl: {output_cfg['path']} (sheet-ovi: {', '.join(sheets)})")
        return 0

    except (LoaderError, TransformError) as exc:
        print(f"[GREŠKA] {exc}", file=sys.stderr)
        return 1
    except Exception as exc:  # pragma: no cover - genericna zastita za neocekivane greske
        print(f"[NEOČEKIVANA GREŠKA] {exc}", file=sys.stderr)
        return 1


def main() -> None:
    args = parse_args()
    sys.exit(run(args))


if __name__ == "__main__":
    main()
