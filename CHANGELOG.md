# Changelog

## 0.1.0 - 2026-11-06

First release.

- Read tools: `workbook_overview`, `read_range` (values / formulas / both; markdown, csv, json),
  `find_cells` (text, regex, number), `describe_columns`, `check_formulas`.
- Edit tools: `set_cells`, `append_rows`, `add_sheet`, `rename_sheet`. All dry-run by default, all
  write a new `<name>.edited.xlsx` unless told otherwise, atomic write, formula counts verified
  after the save.
- Up-front detection (and refusal unless `allow_loss=true`) of everything openpyxl cannot round-trip:
  charts, images, shapes, pivot tables, macros, slicers, timelines, external links, threaded
  comments, form controls, embedded objects, data connections, the data model, dynamic-array
  metadata, sparklines, x14 conditional formatting and data validation.
- `csv_to_xlsx` and `xlsx_to_csv` with explicit encoding and delimiter.
- Safety: path sandbox (`--dir`), zip-bomb limits (entry count, sizes, ratio), refusal of any XML
  part with a DOCTYPE or ENTITY, UTF-8 only XML.
