# Tenancy Schedule Tools

Browser-only tools for preparing and comparing valuation schedules. Property.xlsx and TS.xlsx files are parsed and generated in the browser. Workbooks are never uploaded.

## What the app does

```
Property.xlsx + TS.xlsx
        ↓
fixed source schema validation
        ↓
canonical Property / Unit data
        ↓
Default schedule, Property Only schedule, or TS Compare
```

- Source fields and output columns are defined in code. A new user does not configure mappings.
- Property Name is the fixed join between a Property record and its Unit rows.
- Sector, from Property `General > Sector`, groups the Default schedule. Blank sectors are labelled Unclassified.
- Property.xlsx is the master property list. A valid property is still written when it has no matching units.
- Rows with a blank Property Name, including source total rows, are left out of Default, Property Only, and TS Compare.
- Excel processing stays in the browser. No source data is sent to a server.

Visible schedule names are Default and Property Only. Internal ids remain `tenancy_schedule` and `property_summary`.

## Local development

```bash
npm install
npm run dev
```

Open the URL Vite prints, typically `http://localhost:5173`.

## Production build

```bash
npm run build
npm run preview
```

Output is written to `dist/`.

## Checks

```bash
npm test
```

This runs the local regression script. It builds temporary workbooks in memory. If `samples/Property_sample.xlsx` and `samples/ts_sample.xlsx` are present, those fixtures are parsed as well. They are not imported by the application.

## Where to change the fixed schema

| Change | File |
| --- | --- |
| Property source field names | `src/schema/sourceFields.js` (`PROPERTY_SOURCE_FIELDS`) |
| Unit source field names | `src/schema/sourceFields.js` (`UNIT_SOURCE_FIELDS`) |
| Required Property fields | `src/schema/propertySourceSchema.js` |
| Required Unit fields | `src/schema/unitSourceSchema.js` |
| Default columns, labels, sources, formats, widths | `src/templates/defaultSchedule.js` |
| Default Excel layout | `src/excel/writer.js` |
| Property Only columns, labels, sources, formats, widths | `src/templates/propertyOnly.js` |
| TS Compare identity and compared Unit fields | `src/schema/sourceFields.js` (`COMPARE_IDENTITY`) and Unit sources on `src/templates/defaultSchedule.js` |
| TS Compare statuses and highlighting | `src/compare/engine.js` |

Header detection looks for Property `External ID` and Unit `Lease ID`. Those marker names live in `src/schema/sourceFields.js`.

## GitHub Pages

`.github/workflows/deploy.yml` builds on `main` with `VITE_BASE_PATH=/<repo-name>/`. Local builds use `./`.

## Privacy

Files stay in memory for the browser session. They are not written to localStorage or sent to an API.
