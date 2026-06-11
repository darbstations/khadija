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

## Tech

- [Vite](https://vitejs.dev/) + [React 18](https://react.dev/)
- [Tailwind CSS](https://tailwindcss.com/) (v3)
- [Recharts](https://recharts.org/) for charts
- [SheetJS (xlsx)](https://sheetjs.com/) for parsing Excel files
