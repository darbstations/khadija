# khadija — داش بورد محطات درب الوقود

Fuel-station sales & visits dashboard (Vite + React + Tailwind). Analyzes
monthly Excel exports per station and renders revenue, visits, peak hours,
product mix, payment methods, and daily-traffic estimates for March–May 2026.

## Getting started

```bash
npm install
npm run dev
```

Open the printed local URL. You'll first see an upload screen — select the
station `.xlsx` files, then press **بدء التحليل** to run the dashboard.

## How file loading works

The dashboard component (`src/App.jsx`) reads Excel files via
`window.fs.readFile(name)` — an API that exists only in the Claude artifact
sandbox. To run it as a real web app **without modifying `App.jsx`**, this
project provides a `window.fs` shim:

- `src/FileGate.jsx` — upload screen that collects the `.xlsx` files.
- `src/fileStore.js` — keeps the files in memory and exposes `window.fs.readFile`.
- `src/main.jsx` — shows the gate first, mounts `App` once files are loaded.

Files are processed entirely in the browser and are never uploaded to a server.
Filenames should match the codes defined in `FILE_GROUPS` in `src/App.jsx`
(e.g. `MK072 3.xlsx`); matching is whitespace- and case-insensitive.

## Build

```bash
npm run build
npm run preview
```

## Standalone HTML (no install needed)

Two single-file versions exist for opening the dashboard directly in a browser:

- **`dashboard.html`** — loads React/Recharts/SheetJS/Tailwind from CDNs, so it
  needs internet on first open. Tiny file, easy to share.
- **`dashboard-offline.html`** — everything inlined (no CDN, no network at all).
  Regenerate with `npm run build:single`, which writes `dist-single/index.html`;
  copy that over `dashboard-offline.html` to refresh it.

Both show an upload screen first; pick the station `.xlsx` files, then press
**بدء التحليل**. Files are processed entirely in the browser.

## Tech

- [Vite](https://vitejs.dev/) + [React 18](https://react.dev/)
- [Tailwind CSS](https://tailwindcss.com/) (v3)
- [Recharts](https://recharts.org/) for charts
- [SheetJS (xlsx)](https://sheetjs.com/) for parsing Excel files
