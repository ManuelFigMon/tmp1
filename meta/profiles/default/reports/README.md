# DMQ Reports — Setup & Library Notes

Standalone, single-file HTML reports for the DMQ analytics platform. No frameworks,
no build step — each report is a self-contained `.html` file using vanilla JavaScript,
with all third-party libraries loaded from the **local** `scripts/` and `styles/`
folders. **No external CDN is referenced at runtime** (air-gapped safe).

## Reports

| File | Description |
|---|---|
| `distribution_analysis.html` | Column distribution explorer |
| `kmeans_cluster.html` | K-Means cluster scatter + observations grid |
| `correlation_matrix.html` | Pearson correlation heatmap (template for all reports) |
| `boxplot.html` | Grouped box & whisker plots |
| `histogram_overlay.html` | Overlapping histograms with KDE curves |
| `hexbin_plot.html` | Hexagonal density bins |
| `time_series_explorer.html` | Multi-series time-series line graph + linked AG Grid table |

## Directory layout

```
meta/profiles/default/
├── data_config.json        ← baseFileURL / baseAPIURL / dataset→file Mapping
├── datasets.json           ← dataset metadata: columns, types, dmsrc_type
├── sas_config.json         ← SAS session / logon configuration
├── reports.json            ← report registry
└── reports/
    ├── *.html              ← the reports (this folder)
    ├── scripts/
    │   ├── d3.min.js                  (D3.js v7.9.0)
    │   ├── html2canvas.min.js         (html2canvas 1.4.1)
    │   ├── lodash.min.js              (Lodash 4.17.21)
    │   └── ag-grid-community.min.js   (AG Grid Community 35.3.1, UMD bundle)
    └── styles/
        ├── ag-grid.css                (AG Grid structural CSS, 35.3.1)
        └── ag-theme-alpine.css        (AG Grid Alpine legacy theme, 35.3.1)
```

The reports resolve config files via `getProfileBasePath()` → `'../'`, i.e. one level
above the `reports/` folder. Keep that structure when copying to user profiles
(`meta/profiles/{userid}/reports/`).

## Obtaining the libraries (one-time, from a connected machine)

Download each file below on an internet-connected machine, then transfer them into
`reports/scripts/` and `reports/styles/`:

| Target file | Download URL |
|---|---|
| `scripts/d3.min.js` | https://unpkg.com/d3@7.9.0/dist/d3.min.js |
| `scripts/html2canvas.min.js` | https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js |
| `scripts/lodash.min.js` | https://unpkg.com/lodash@4.17.21/lodash.min.js |
| `scripts/ag-grid-community.min.js` | https://unpkg.com/ag-grid-community@35.3.1/dist/ag-grid-community.min.js |
| `styles/ag-grid.css` | https://unpkg.com/ag-grid-community@35.3.1/styles/ag-grid.css |
| `styles/ag-theme-alpine.css` | https://unpkg.com/ag-grid-community@35.3.1/styles/ag-theme-alpine.css |

Alternatively, AG Grid releases can be downloaded from
https://github.com/ag-grid/ag-grid/releases (use the `ag-grid-community` package's
`dist/` and `styles/` folders), or via `npm pack ag-grid-community@35.3.1`.

> **AG Grid version note:** the reports use the modern `agGrid.createGrid()` API and
> set `theme: 'legacy'` in the grid options, which is required for the CSS-file themes
> (`ag-theme-alpine`) on AG Grid v33+. If you swap in a different major version,
> verify both of those still apply.

Each report performs a startup check: if any local library failed to load
(`d3`, `html2canvas`, `_`, or `agGrid` undefined), a red banner names the missing
file instead of failing silently.

## Time Series Explorer — feature summary

- **Inputs:** dataset; X-axis (date/time preferred, any column allowed); one or more
  numeric Y columns (each an overlaid line); optional character Group column that
  splits lines per distinct value (Y × group combos capped at 20 with a warning);
  template-style WHERE filters.
- **Configuration:** include/exclude outliers (1.5×IQR fences, excluded count shown
  in KPI strip), Mean / Median / Mode / P25 / P75 reference lines, running-minimum
  "best so far" red markers + dashed step line, Linear/Step/Smooth interpolation,
  point-size slider, legend toggle.
- **Chart:** D3 v7 SVG (responsive viewBox), light horizontal gridlines, rotated
  date labels, hover tooltip (series, group, X, Y, Δ vs mean), vertical crosshair,
  clickable legend to hide/show series, click-to-select points.
- **Grid:** AG Grid Community — multi-column sorting, floating per-column filters
  (text/number/date by column type), pagination (10/25/50/100, default 25),
  movable/resizable columns, checkbox multi-row selection, quick-filter box,
  built-in CSV export.
- **Linked selection (bidirectional):** chart point clicks select grid rows and
  scroll the first into view; grid row selection enlarges/rings the matching chart
  points (orange) and dims the rest. A shared `Set` of stable row IDs keeps both
  views consistent; Clear Selection resets everything.
- **Template features:** SAS session/logon, report save (PUT to profile
  `reports.json`), About/Help modals, Export Image (html2canvas), chart-data CSV
  export, Reset Form.
