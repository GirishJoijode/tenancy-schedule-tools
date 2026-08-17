/**
 * Historical helper. Shipped defaults now live in src/config/defaultConfig.js
 * (promoted from the live browser export). Do not regenerate from the old
 * standalone extract — that would overwrite current mappings.
 */
console.error(
  'Refusing to overwrite src/config/defaultConfig.js from extracted-config.json.\n' +
    'The shipped default is the promoted live configuration. Edit defaultConfig.js only via a verified live export.'
);
process.exit(1);
