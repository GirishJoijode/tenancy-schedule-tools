import { readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, '..');

let body = readFileSync(resolve(here, 'extracted-body.html'), 'utf8');

body = body.split('? Header row detected').join('\u2713 Header row detected');

body = body.replace(
  `        <div class="config-actions">
          <button type="button" id="btn-apply-config" class="btn-secondary">
            Apply Configuration
          </button>
          <button type="button" id="btn-export-html" class="btn-primary">
            Export Configured HTML
          </button>
          <button type="button" id="btn-reset-mapping" class="btn-danger">
            Reset Mapping
          </button>
        </div>
        <p class="settings-note">
          <strong>Export Configured HTML</strong> downloads a new standalone copy of this application
          with your mappings embedded inside it. The browser cannot overwrite the file that is currently open.
          Use the downloaded file when distributing a pre-configured app to other users.
        </p>`,
  `        <div class="config-actions">
          <button type="button" id="btn-apply-config" class="btn-secondary">
            Apply Configuration
          </button>
          <button type="button" id="btn-export-config" class="btn-primary">
            Export Configuration
          </button>
          <button type="button" id="btn-import-config" class="btn-secondary">
            Import Configuration
          </button>
          <button type="button" id="btn-reset-mapping" class="btn-danger">
            Reset Mapping
          </button>
          <input
            type="file"
            id="config-import-input"
            accept="application/json,.json"
            class="file-input"
          />
        </div>
        <p class="settings-note">
          <strong>Export Configuration</strong> downloads a JSON file of the current mappings.
          <strong>Import Configuration</strong> loads a previously exported JSON file after validation.
          Edits are stored in this browser only (localStorage) and never uploaded.
          <strong>Reset Mapping</strong> restores the original default mappings that shipped with the application.
        </p>`
);

const html = `<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Tenancy Schedule Standardiser</title>
  </head>
  <body>
${body}
    <script type="module" src="/src/main.js"></script>
  </body>
</html>
`;

writeFileSync(resolve(root, 'index.html'), html, 'utf8');
console.log('Wrote Vite index.html');
