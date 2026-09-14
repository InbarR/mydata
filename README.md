# MyData

MyData is a local-first browser tool for exploring tabular data. Import Excel
or CSV files, then search, filter, group, format, reorder, and export the data
without uploading the source file.

## Open MyData

[Open the hosted app](https://inbarr.github.io/mydata/)

## Privacy

MyData processes imported files locally in your browser. The hosted app has no
backend, analytics, telemetry, or file-upload service. See
[PRIVACY.md](./PRIVACY.md).

Exports contain the rows visible in the exported view. Treat exported files
according to the sensitivity of their source data.

## Features

- Multiple workbooks and worksheets
- Combined workbook and all-source views
- Smart search, including AND, OR, NOT, `!`, quoted phrases, regex, and optional match highlighting
- Advanced filters, grouping, sorting, and pagination
- Column visibility, renaming, resizing, and fit-to-content
- Manual row ordering
- Cell and row clipboard actions from the right-click menu
- Value colors, per-value palettes, and conditional formatting
- Excel, CSV, and standalone HTML exports
- Self-contained offline HTML artifact

## Run locally

```powershell
npm install
npm run dev
```

## Build

```powershell
npm ci
npm run build
```

The Vite build produces a self-contained `dist/index.html`.

## Offline artifact

Download and open [`mydata.html`](./mydata.html) directly in a modern browser.

## License

[MIT](./LICENSE)
