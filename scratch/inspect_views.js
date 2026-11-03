const fs = require('fs');
const html = fs.readFileSync('mobile-web/index.html', 'utf8');

const m1 = html.match(/id="modal-bottom-sheet-tx"[\s\S]*?<\/div>\s*<\/div>/i);
if (m1) console.log('=== modal-bottom-sheet-tx ===\n', m1[0]);

const m2 = html.match(/id="modal-budget-alloc"[\s\S]*?<\/div>\s*<\/div>/i);
if (m2) console.log('=== modal-budget-alloc ===\n', m2[0]);
