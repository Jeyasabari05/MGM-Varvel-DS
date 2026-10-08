# Workbook inspection and data dictionary

## Workbook layout

All ten files use a single worksheet named `Sheet2`. A sheet is a cross-tab rather than a conventional row-per-record table:

- Columns A–C identify the row, parameter label, and unit / sub-label.
- Each motor record occupies one column beginning at D.
- Row 1 identifies the motor family/type, row 2 rated power in kW, row 4 frame, row 5 mounting, and row 12 efficiency class.
- Rows 6 onward contain electrical, performance, mechanical, enclosure, construction, cable, finish, and (where present) brake parameters.
- Rows 20–25 hold load-specific efficiency and power-factor points. Row 30–31 split DE/NDE bearing values. Row 34–35 split housing/flange material.
- Brake motor workbooks include rows 55–62. In the SMX reference workbooks those brake rows are absent.
- Literal `Italy` entries are source placeholders, not numeric ratings. They are not converted or estimated.

## Source mapping and coverage

| Source file | Family data | Efficiency | Brake configuration | Distinct power values |
| --- | --- | --- | --- | ---: |
| `1. BA - IE3 - AC.xlsx` | BAX (0.25–30 kW), BAHX (37–132 kW) | IE3 | AC | 15 BAX + 7 BAHX |
| `2. BA - IE3 - DC.xlsx` | BAX | IE3 | DC | 17 |
| `3. BM - IE3.xlsx` | BMX | IE3 | no brake rows | 19 |
| `4. SM - IE3.xlsx` | BMX and SMX | IE3 | no brake rows | 26 combined columns |
| `5. BA - IE2 - AC.xlsx` | BAX (0.25–30 kW), BAHX (37–132 kW) | IE2 | AC | 15 BAX + 7 BAHX |
| `6. BA - IE2 - DC.xlsx` | BAX | IE2 | DC | 17 |
| `7. BM - IE2.xlsx` | BMX | IE2 | DC | 19 |
| `8. SM - IE2.xlsx` | SMX | IE2 | no brake rows | 24 |
| `SM - IE2 (upto 55kW).xlsx` | SMX | IE2 | no brake rows | overlaps the 55 kW range |
| `SM - IE3 (upto 55kW).xlsx` | SMX | IE3 | no brake rows | overlaps the 55 kW range |

On inspection, the combined unique class/range coverage is:

- BAX IE2 and IE3 AC: 0.25–30 kW, 15 source power points each; DC: 0.25–45 kW, 17 points each.
- BAHX IE2 and IE3 AC: 37–132 kW, 7 source power points each, read from the BAHX columns in the corresponding BA AC workbooks.
- BMX IE2 and IE3: 0.12–45 kW, 19 points each.
- SMX IE2 and IE3: 0.12–132 kW, 24 points each. The extra reference files overlap the base SMX records through 55 kW.

The exact list is computed from the files at runtime; the ranges above summarize the observed source catalog only.

## Document field mapping

The PDF page is US Letter portrait. The generation layer keeps its logo, labels, grid, title, footer, and page geometry. Source rows are mapped to the corresponding reference cells, including rated power, poles, frame, mounting, voltage, frequency, connection, current, speed, torque, efficiency class, duty, insulation, temperature rise, IP, cooling, locked-rotor ratios, load efficiencies, power factors, bearing, materials, terminal/cable fields, finish, vibration, sound, and brake details.

The three locked-rotor / torque values (`Current - Locked rotor / Rated`, `Torque - Locked rotor / Rated`, and `Torque - Breakdown / Rated`) are copied directly from their matching Excel cells into the generated PDF without calculation or conversion.

Brake coil voltage retains the workbook's AC or DC designation. For DC brake records, the static `(AC)` unit suffix in the reference template is replaced with `(DC)` to match the source coil-voltage field. Input supply voltage remains labeled AC, as it is in each workbook.

The reference calls the field `Temperature rise class`; workbook row 14 contains `155(F) utilized to 130(B)`, so the source-backed insulation class is F and temperature-rise class is B. The source files observed here support mounting `IM B5` (printed as B5 to match the PDF), protection `IP 55`, and cooling `TEFC / IC411`. No B3/B14, H insulation, F temperature rise, other IP classes, or other cooling methods were present, so none are offered.

## Duplicate checks observed

The selector compares duplicate source records within motor family + efficiency + rated power + selected brake type. AC and DC brake columns are separate source configurations and become unambiguous after the user chooses one; their differing values are not treated as duplicates. It merges repeated records only when fields used as motor/datasheet specifications agree. A conflict in a mapped field blocks the affected configuration. During inspection, source conflicts were found at:

- BMX IE3, 0.12 kW and 0.18 kW: `Type of terminal box` differs between `3. BM - IE3.xlsx` and the BMX columns in `4. SM - IE3.xlsx`.

Those combinations intentionally cannot generate a datasheet until the source data is corrected or clarified. Brake specification differences are treated as separate AC/DC configurations and must be selected explicitly where both exist.
