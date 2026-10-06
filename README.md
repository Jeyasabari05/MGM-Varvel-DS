# MGM Varvel Motor Datasheet Selector

A local Python application that reads the supplied motor workbooks, narrows valid combinations by family / efficiency / power, and fills the reference datasheet PDF. The gearbox tiles are navigation placeholders.

## Run

1. Install Python 3.10+.
2. From this folder run `python -m pip install -r requirements.txt`.
3. Run `python app.py` and open `http://127.0.0.1:8765` in a browser.

The server binds to localhost only. It reads the original workbooks from `data/source_files/` on startup. Keep all ten workbooks and `templates/DS-025-2028.pdf` in place.

Open `/api/diagnostics` on the local server for a current machine-readable source error and duplicate-conflict report.

## Data rules

- Workbooks are read from their actual worksheets and motor columns; no power/frame catalogue is hard-coded.
- A record is keyed by motor family, efficiency, rated power, and source-listed brake configuration when present. The user chooses AC/DC before the app evaluates duplicates across brake records.
- Identical overlapping motor records are coalesced while retaining both source filenames. Conflicting displayed engineering values block generation and name the conflict in the UI.
- Text such as `Italy` is an unavailable source value and is left blank in the PDF.
- The selector only shows source-backed mounting, cooling, insulation, temperature rise, and protection values.
- Datasheets retain the supplied one-page PDF artwork and table; source values are placed into the matching value cells.

## Files

- `app.py` — workbook data layer, validation, local HTTP app, PDF rendering.
- `static/` — selector and landing pages.
- `data/source_files/` — source workbooks, copied unchanged.
- `templates/DS-025-2028.pdf` — reference PDF template.
- `assets/` — provided landing art and MGM Varvel logo from the sibling selector project.
- `DATA_DICTIONARY.md` — workbook schema and verified catalog coverage.
- `tests/` — regression checks for catalog dependencies, source conflicts, invalid combinations, and generated PDF values.

Run the regression checks with `python -m unittest discover -s tests`.

The sibling `mgm-varvel-selector` project was reviewed for its read-only dropdown pattern and exact-match behavior. This motor selector is kept as an independent Python application because its inputs and reference datasheet use a different source-data structure.
