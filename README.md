# Excel Data Processor

Aplikacija za obradu podataka iz lokalne Excel "baze" (jedan ili više Excel
fajlova) i eksport rezultata u nove, formatirane Excel fajlove sa više
sheet-ova (`Raw`, `Processed`, `Summary`).

Ceo tok obrade (učitavanje → filtriranje → transformacija → agregacija →
eksport) je definisan kroz jedan YAML/JSON konfiguracioni fajl, bez potrebe
za izmenom koda za svaki novi izveštaj.

## Struktura projekta

```
config/
  export_config.yaml     - primer konfiguracije (scenario "Agriculture 2025")
data/
  input.xlsx              - primer ulaznih podataka (sheet "Transactions")
output/
  *.xlsx                  - ovde se upisuju generisani Excel fajlovi
scripts/
  generate_sample_data.py - generiše data/input.xlsx (primer podataka)
src/
  config.py                - učitavanje i validacija konfiguracije
  loader.py                 - učitavanje Excel podataka (fajl ili folder)
  transformer.py             - filtriranje, transformacije, agregacija
  exporter.py                 - eksport u novi .xlsx sa formatiranjem
  main.py                      - CLI ulazna tačka
requirements.txt
```

## Instalacija

Potreban je Python 3.10+.

```bash
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Brzi start (primer scenario)

Repo već sadrži primer ulaznih podataka (`data/input.xlsx`) i primer
konfiguracije (`config/export_config.yaml`) za sledeći scenario:

- **ulaz:** `data/input.xlsx`, sheet `Transactions`
- **filter:** samo zapisi iz 2025. godine i kategorija `Agriculture`
- **transformacija:** nova kolona `value_eur = value_rsd / 117`
- **izlaz:** `output/agriculture_2025.xlsx` sa sheet-ovima `Raw`, `Processed`, `Summary`

Pokretanje:

```bash
python src/main.py --config config/export_config.yaml --verbose
```

Ako želite da ponovo generišete primer ulaznih podataka:

```bash
python scripts/generate_sample_data.py
```

## Korišćenje sa sopstvenim podacima

1. Kopirajte svoj Excel fajl (ili folder sa više Excel fajlova) negde u
   projekat, npr. u `data/`.
2. Napravite kopiju `config/export_config.yaml` (npr. `config/moj_izvestaj.yaml`)
   i podesite sekcije `input`, `filters`, `transformations`, `summary` i `output`
   (svaka sekcija je komentarisana u samom fajlu).
3. Pokrenite:

   ```bash
   python src/main.py --config config/moj_izvestaj.yaml
   ```

### CLI argumenti

| Argument | Opis |
|---|---|
| `--config`, `-c` | Putanja do YAML/JSON konfiguracije (podrazumevano `config/export_config.yaml`) |
| `--input` | Prepisuje `input.path` iz konfiguracije, bez izmene fajla |
| `--output` | Prepisuje `output.path` iz konfiguracije, bez izmene fajla |
| `--verbose`, `-v` | Ispisuje broj redova posle svakog koraka (učitavanje, filter, summary) |

## Format konfiguracije

Konfiguracija je YAML (ili JSON) fajl sa sledećim sekcijama:

### `input` - odakle se čita

```yaml
input:
  path: "data/input.xlsx"   # fajl ILI folder (svi Excel fajlovi u folderu se spajaju)
  sheet: "Transactions"     # ime sheet-a (ili broj, 0 = prvi)
  columns: []               # opciono: whitelist kolona; prazno = sve kolone
  tag_source_file: false    # ako je 'path' folder, dodaje kolonu '_source_file'
```

### `filters` - filtriranje (datum, kategorija, region, ...)

```yaml
filters:
  date_column: "date"
  date_from: "2025-01-01"
  date_to: "2025-12-31"
  in:
    category: ["Agriculture"]
  equals: {}       # {kolona: vrednost} - tačno poklapanje
  min: {}          # {kolona: broj} - kolona >= broj
  max: {}          # {kolona: broj} - kolona <= broj
```

### `transformations` - nove/izmenjene kolone

```yaml
transformations:
  compute_columns:
    - name: "value_eur"
      expression: "value_rsd / 117"   # evaluira se preko pandas DataFrame.eval
  round:
    value_eur: 2
  rename_columns: {}
  drop_columns: []
```

### `summary` - agregacija za rezimirajući sheet

```yaml
summary:
  enabled: true
  group_by: ["category", "region"]
  aggregations:
    quantity: "sum"
    value_eur: "sum"
```

### `output` - gde i kako se piše rezultat

```yaml
output:
  path: "output/agriculture_2025.xlsx"
  sheet_names:
    raw: "Raw"
    processed: "Processed"
    summary: "Summary"
  include_raw: true
  include_processed: true
  number_format: "#,##0.00"
  freeze_header: true
  autofilter: true
```

## Validacija i greške

- Nedostajući fajl/folder, nepostojeći sheet ili nepostojeće kolone se
  prijavljuju sa jasnom porukom (i listom dostupnih kolona) pre nego što se
  bilo šta upiše u `output/`.
- Konfiguracija se strukturno validira pri učitavanju (`src/config.py`) -
  npr. obavezna polja `input.path` i `output.path`.

## Napomena o `compute_columns` izrazima

Izrazi u `transformations.compute_columns` se evaluiraju preko
[`pandas.DataFrame.eval`](https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.eval.html),
što znači da mogu da referenciraju samo postojeće kolone i osnovne
aritmetičke/logičke operacije - nemaju pristup proizvoljnom Python kodu,
fajl sistemu ili modulima.
