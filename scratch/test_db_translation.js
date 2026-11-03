const fs = require('fs');
const baseline = JSON.parse(fs.readFileSync('mobile-web/baseline_data.json', 'utf8'));

console.log('--- COMPTES ---');
baseline.comptes.forEach(c => console.log(c.id, c.nom, c.classe));

console.log('--- CATEGORIES ---');
baseline.categories.forEach(c => console.log(c.id, c.nom, c.type));

console.log('--- TIERS ---');
baseline.tiers.forEach(t => console.log(t.id, t.nom, t.type));

console.log('--- NOTES SAMPLE ---');
const notes = [...new Set(baseline.transactions.map(t => t.notes))];
console.log(notes.slice(0, 15));
