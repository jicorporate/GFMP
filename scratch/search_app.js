const fs = require('fs');
const appJs = fs.readFileSync('mobile-web/app.js', 'utf8');

const terms = ['tx-compte', 'tx-categorie', 'tx-tiers', 'exp-tx-compte', 'alloc-categorie', 'populate', 'select', 'addCompte', 'addCategorie', 'addTiers'];

terms.forEach(term => {
  const lines = appJs.split('\n');
  const matches = [];
  lines.forEach((l, idx) => {
    if (l.toLowerCase().includes(term.toLowerCase())) {
      matches.push(`${idx + 1}: ${l.trim()}`);
    }
  });
  console.log(`=== Matches for "${term}" (${matches.length}) ===`);
  console.log(matches.slice(0, 10).join('\n'));
});
