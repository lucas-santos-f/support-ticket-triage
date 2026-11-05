// Copies data/messages.* into site/sample-data.js so the page works when opened from disk.
const fs = require('node:fs');
const path = require('node:path');

const root = path.join(__dirname, '..');
const sample = {
  asOf: '2026-11-05',
  csv: fs.readFileSync(path.join(root, 'data', 'messages.csv'), 'utf8'),
  xlsx: fs.readFileSync(path.join(root, 'data', 'messages.xlsx')).toString('base64'),
};
fs.writeFileSync(path.join(root, 'site', 'sample-data.js'), 'window.SAMPLE = ' + JSON.stringify(sample, null, 1) + ';\n');
console.log('wrote site/sample-data.js');
