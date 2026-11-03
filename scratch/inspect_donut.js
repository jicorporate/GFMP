const fs = require('fs');
global.window = global;
global.localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {} };
eval(fs.readFileSync('mobile-web/baseline_data.js', 'utf8'));
eval(fs.readFileSync('mobile-web/analytics-engine.js', 'utf8'));
const engine = window.gfmpEngine;
engine.init();
const cf = engine.computeCashflow();

const size = 130;
const strokeW = 16;
const radius = (size - strokeW) / 2;
const center = size / 2;
const circ = 2 * Math.PI * radius;

let accumPct = 0;
cf.breakdown.forEach(item => {
  const slicePct = item.pct / 100;
  const strokeDash = `${(slicePct * circ).toFixed(2)} ${circ.toFixed(2)}`;
  const strokeOffset = (-accumPct * circ).toFixed(2);
  accumPct += slicePct;
  console.log(item.nom, 'color:', item.color, 'pct:', item.pct.toFixed(1), 'dash:', strokeDash, 'offset:', strokeOffset);
});
