# Tenancy Schedule Tools

Browser-only Standardiser and TS Compare tools. Property and Unit `.xlsx` files are parsed and generated **in the browser**. Workbooks are never uploaded.

This is a Vite + vanilla JavaScript migration of the original standalone `index.html` app. Mapping configuration was copied from that file without remapping.

## Local development

```bash
npm install
npm run dev
```

Then open the URL Vite prints (typically `http://localhost:5173`).

## Production build

```bash
npm run build
npm run preview
```

The production output is written to `dist/`.

## GitHub Pages

The workflow in `.github/workflows/deploy.yml` builds on `main` (and on `workflow_dispatch`) with:

```
VITE_BASE_PATH=/${{ github.event.repository.name }}/
```

That sets Vite `base` to `/<repo-name>/` so assets resolve under GitHub Pages project sites. Local builds use `./` so `dist/` also works from a relative path.

Enable GitHub Pages for the repository with **GitHub Actions** as the source.

## Configuration

- **Default mappings** ship in `src/config/defaultConfig.js` (the original embedded configuration).
- **Runtime edits** are saved in this browser only, under the localStorage key `tenancyScheduleTools.config.v1`.
- On load, a valid stored configuration is used; otherwise the shipped defaults are used.
- **Export Configuration** downloads `Tenancy_Schedule_Config.json`.
- **Import Configuration** validates the JSON file before applying it. Invalid files are rejected in full — nothing partial is applied.
- **Reset Mapping** restores the original shipped defaults, not empty blanks.
- Excel files stay in memory for the session only. They are not written to localStorage or IndexedDB.

## Privacy

Files are processed locally in the browser with `xlsx-js-style`. No workbook contents are sent to a server.

## Verify mappings

```bash
npm run verify:config
```

This checks that `src/config/defaultConfig.js` still matches the extracted standalone configuration value-for-value.
