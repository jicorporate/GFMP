const fs = require('fs');
const engineSrc = fs.readFileSync('mobile-web/analytics-engine.js', 'utf8');

// Check that dictionary from build_dictionary.js is clean
const dictCode = fs.readFileSync('scratch/build_dictionary.js', 'utf8');
const dictMatch = dictCode.match(/const dbEntityDict = \{[\s\S]*?\n\};/);
if (!dictMatch) throw new Error('Dict not found');

console.log('Dict length:', dictMatch[0].length);
