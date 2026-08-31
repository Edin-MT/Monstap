"""
config.py
---------
Učitavanje i validacija konfiguracije aplikacije (YAML ili JSON).

Konfiguracioni fajl opisuje čitav tok obrade podataka:
- odakle se učitavaju ulazni Excel podaci ("input"),
- kako se ti podaci filtriraju ("filters"),
- kako se transformišu, npr. nove kolone ("transformations"),
- kako se agregiraju/grupišu za rezimirajući sheet ("summary"),
- gde i kako se upisuje rezultat ("output").

Vidi config/export_config.yaml za potpuno komentarisan primer.
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict

import yaml


class ConfigError(Exception):
    """Podignuto kada konfiguracija ne postoji, ne može da se parsira ili je nevalidna."""


def load_config(config_path: str) -> Dict[str, Any]:
    """Učitava konfiguraciju iz YAML ili JSON fajla i vrši osnovnu validaciju.

    Args:
        config_path: putanja do konfiguracionog fajla (.yaml, .yml ili .json).

    Returns:
        Rečnik sa konfiguracijom (spreman za dalju upotrebu u loader/transformer/exporter).

    Raises:
        ConfigError: ako fajl ne postoji, nije validan YAML/JSON ili mu nedostaju
            obavezna polja.
    """
    path = Path(config_path)
    if not path.exists():
        raise ConfigError(f"Konfiguracioni fajl ne postoji: {config_path}")

    text = path.read_text(encoding="utf-8")

    try:
        if path.suffix.lower() == ".json":
            data = json.loads(text)
        else:
            # YAML parser bez problema učitava i JSON (JSON je podskup YAML-a),
            # pa ovo pokriva i .yaml/.yml i nepoznate ekstenzije.
            data = yaml.safe_load(text)
    except (yaml.YAMLError, json.JSONDecodeError) as exc:
        raise ConfigError(f"Neispravan format konfiguracije ({config_path}): {exc}") from exc

    if not isinstance(data, dict):
        raise ConfigError(f"Konfiguracija mora biti mapa/objekat na najvišem nivou: {config_path}")

    validate_config(data)
    return data


def validate_config(cfg: Dict[str, Any]) -> None:
    """Osnovna strukturna validacija konfiguracije.

    Napomena: ovde se proverava samo oblik konfiguracije (obavezna polja).
    Validacija da li navedene kolone/sheet-ovi zaista postoje u Excel fajlu
    radi se kasnije, u loader.py, jer to zahteva da fajl bude učitan.
    """
    if "input" not in cfg or not isinstance(cfg["input"], dict):
        raise ConfigError("Konfiguracija mora imati sekciju 'input' (rečnik).")
    if not cfg["input"].get("path"):
        raise ConfigError("Sekcija 'input' mora imati 'path' (putanja do fajla ili foldera).")

    if "output" not in cfg or not isinstance(cfg["output"], dict):
        raise ConfigError("Konfiguracija mora imati sekciju 'output' (rečnik).")
    if not cfg["output"].get("path"):
        raise ConfigError("Sekcija 'output' mora imati 'path' (putanja izlaznog .xlsx fajla).")

    summary_cfg = cfg.get("summary") or {}
    if summary_cfg.get("enabled"):
        if not summary_cfg.get("group_by"):
            raise ConfigError(
                "Kada je 'summary.enabled: true', mora se navesti neprazan 'summary.group_by'."
            )
        if not summary_cfg.get("aggregations"):
            raise ConfigError(
                "Kada je 'summary.enabled: true', mora se navesti neprazan 'summary.aggregations'."
            )
