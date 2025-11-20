# TruTrade Report Automation – Agent Notes

## Goal
Act as my coding pair to complete and harden the TruTrade report generator. You should:
- Fully understand the current Python + PPT + Excel flow (read below), then modify the specific sections called out when we iterate.
- Keep the pipeline stable: uploaded `.xlsm/.xlsx` ? parse sheets ? fill `TT_report.pptx` placeholders, tables, and the mileage chart ? stream PPT download.
- When editing, focus on the precise code locations noted in the "Code map" so changes stay scoped.

## App overview
- Web app: Flask (`app.py`) serves a single-page form, handles file upload, builds a PPT from a master template, and streams the result.
- Inputs: user-provided `TT_worksheet.xlsm` (or `.xlsx`) whose VBA macros already ran and stored AI insights + metrics in specific cells.
- Outputs: a stamped PPT copy from `TT_report.pptx` with placeholders replaced, tables copied/ formatted, and the mileage bar chart rebuilt.

## User flow
1) Download `TT_worksheet.xlsm` from `/download-template`.
2) Paste TruTrade datasets into the template; let the built-in macros run and finish (they call ChatGPT and write insights into the sheet).
3) Upload the completed workbook via `/generate` (accepts `.xlsm`/`.xlsx`).
4) Server parses the target sheets, maps values into the PPT template, rebuilds the mileage chart, and returns `TT_report_<timestamp>.pptx`.

## Key files
- `app.py`: all Flask routes, inline HTML UI, PPT generation, placeholder mapping, table copy/formatting, and chart rebuild logic.
- `TT_report.pptx`: master template whose placeholders, tables, and chart are populated.
- `TT_worksheet.xlsm`: macro workbook that gathers data, calls ChatGPT, and writes insights that Python reads (Python does not execute macros).

## Code map (what to edit, where)
- **Routes (bottom of app.py)**
  - `/`: renders inline HTML via `render_template_string(HTML)`; flashes if PPT template missing. Edit here for UI changes.
  - `/download-template`: streams `TT_worksheet.xlsm`. Adjust if template path/name changes.
  - `/generate`: uploads file, type-checks (`.xlsm/.xlsx`), saves to tempdir, calls `build_presentation`, returns PPT. Add validation or error handling here.

- **HTML template (top of app.py, `HTML` string)**
  - Contains the landing page, download button, upload form, overlay spinner, and JS fetch submit handler that downloads the PPT blob. Modify here for UX copy, styling, or extra fields.

- **`build_presentation(xlsm_path, template_path)`**
  - Loads `TT_report.pptx` into `prs`.
  - Reads Excel sheets via `pd.read_excel(..., sheet_name=["LeasingInfographic", "CompetitiveMarketPosition", "ZipCodes", "DrawDemo", "DistanceTravelled", "Frequency", "Duration", "MileageDemo"], engine="openpyxl")`.
  - Pulls dataframes: `df_leasing`, `df_sheet2`, `df_zipcodes`, `df_drawdemos`, `df_distance`, `df_frequency`, `df_duration`, `df_mileage`.
  - Calls `update_mileage_chart(prs.slides[10], df_mileage)` to rebuild the slide 10 bar chart.
  - `variable_mapping`: replaces text placeholders across shapes and tables (keys like `MTZIP1`, `EXECANALYSIS5`, `VL10`, `CMPANALYSIS10`, many `MA12X`/`MB12X`… etc.) with formatted values (percents, currency, thousands). Edit/add mappings here when template placeholders change or new fields are needed.
  - Table copy/format:
    - `slides_to_update = {9: df_sheet2, 11: df_drawdemos, 14: df_zipcodes, 30: df_distance, 31: df_frequency, 32: df_duration}`. These slides must already contain a table shape; code fills them row/col-wise.
    - `formatting_rules` controls percent/currency/thousands/decimal formatting per sheet; `DrawDemo` also has `percent_rows`/`currency_rows` to decide formatting based on the row label.
    - Headers for some sheets are overwritten with DF column names (except first column in `DrawDemo`). Row heights are fixed for `DrawDemo`.
    - Values are written with Roboto size 9; some headers bold/white for Distance/Frequency/Duration.
  - Returns an in-memory PPT stream.

- **`update_mileage_chart(slide, df_mileage)`**
  - Reads `MileageDemo!Q4:Q26` (0-based rows 2:26, col 16) ? values list; categories are blank.
  - Builds a BAR_CLUSTERED chart at a fixed position; hides axes/labels; shows outside-end value labels, size 9, black.
  - Colors bars: red if <100, green if >100, grey if ==100; no legend.
  - Edit here if chart ranges, styling, or placement change.

## Frontend behavior
- Inline CSS/JS in `HTML`: dark theme, Inter font, sticky header, form with accept `.xlsm,.xlsx`. JS intercepts submit, shows overlay spinner, posts via fetch, and triggers file download with name `TT_report_<ISO timestamp>.pptx`.
- Flash messages surface from Flask when template missing; upload validation is minimal.

## Running
- Install deps: Flask, pandas, python-pptx, openpyxl, werkzeug (`pip install -r requirements.txt`).
- Keep `TT_report.pptx` and `TT_worksheet.xlsm` alongside `app.py`.
- Local dev: `python app.py` (port 5000, debug=True). Optional: set `SECRET_KEY` env for Flask session/flash.

## Constraints / caveats
- Python does not execute Excel macros; the workbook must be pre-populated by VBA before upload.
- Supported sheets/columns are fixed; renamed/missing tabs will break mapping.
- Maps/extra charts beyond mileage are not auto-generated.
- No persistent storage; generated PPT is streamed per request.

## Next edits to consider (instructions for agent)
1) If placeholders change in `TT_report.pptx`, update `variable_mapping` in `build_presentation` and adjust formatting rules if new fields need percent/currency/thousands.
2) If sheet schemas change, update the `pd.read_excel` `sheet_name` list and the `slides_to_update`/`formatting_rules` blocks so tables still align with template tables.
3) If mileage chart range/layout changes, edit `update_mileage_chart` (row/col slices, colors, position/size, label formatting).
4) For stricter uploads or UX tweaks, edit `/generate` (validation/error messages) and the inline `HTML` string (form, copy, loading states).
