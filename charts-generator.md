# Chart Generator Notes

## Existing charts
- Mileage chart (slide 11 in template index, `prs.slides[10]`): built by `update_mileage_chart`; reads `MileageDemo!Q4:Q26`; hidden axes/labels; outside-end values; color red/green/grey based on 100 index; position `left=11.3", top=2.392", width=1.7", height=5.735"`.
- Ethnic makeup chart (slide 9 in template index, `prs.slides[7]`): built by `update_ethnicities_chart`; reads `DrawDemo!W24:X28` (start row 22 zero-based per latest adjustment); dark-blue bars; percent labels outside end; Roboto labels; hidden axes/grid; tuned bar spacing.

## Placement & styling for ETHNIC MAKEUP (current)
- Position: `left=0.44"`, `top=6.19"`, `width=3.8"`, `height=2.1"`.
- Data categories from col W, values from col X (decimals rendered as `%` with number_format `0%`).
- Axis lines/grid removed: value/major grid/axis lines hidden; category axis line hidden.
- Bar spacing: `series.gap_width = 50` for tighter stacking.
- Colors: dark blue `RGB(21,57,96)` for all points.
- Fonts: labels Roboto size 11, black; category axis labels also Roboto size 11, black.

## Adding new charts (pattern)
1) Slice Excel data via pandas (`df.iloc[start:end, col]`), convert to float/None, build `CategoryChartData` or `XyChartData` as needed.
2) Create chart on target slide: `slide.shapes.add_chart(chart_type, left, top, width, height, chart_data)`. Use Inches for layout.
3) Hide unwanted axes/legend (`chart.has_legend=False`, clear axis lines, tick labels `";;;"` when hiding values).
4) Set data labels: `series.data_labels.show_value=True`, position `OUTSIDE_END` for bars, and format (e.g., `0%`). Assign font family/size/color explicitly (Roboto where desired).
5) Style bars/points: iterate `series.points`, set `fill.solid()` and `fore_color.rgb`.
6) Adjust spacing: `series.gap_width` for clustered bars; modify width/height/left/top to align with template sections.
7) Call the new updater inside `build_presentation` after loading DataFrames and before saving the presentation, selecting the correct slide index.

## Slide indices
- Templates are zero-based in python-pptx: slide 8 in PPT UI is `prs.slides[7]`; slide 11 is `prs.slides[10]`.

## Key references in code
- Functions: `update_mileage_chart`, `update_ethnicities_chart` in `app.py`.
- Invocation: inside `build_presentation` after DataFrame loads.

Use these as the canonical rules when adding or adjusting charts.
