// verify_cohort_match.js -- cross-check that the pooled Training Cohort 1 figures in index.html
// (SEED_DAYS + TRAINING_COHORT) equal the tracker's BATCH1_RAW[].<period> totals (bills/units/AOV exact,
// revenue within per-stylist rounding).
// Usage: node scripts/verify_cohort_match.js <period-key, e.g. sep>
// Exit code 0 = match, 1 = mismatch or parse problem.
const fs = require('fs');
const period = process.argv[2] || 'sep';

function grab(t, name) {
  const m = t.match(new RegExp('const ' + name + '\\s*='));
  if (!m) throw new Error('constant not found: ' + name);
  let i = m.index + m[0].length;
  while (/\s/.test(t[i])) i++;
  const o = t[i], c = o === '[' ? ']' : '}';
  let d = 0, j = i, q = null;
  for (;; j++) {
    const ch = t[j];
    if (q) { if (ch === '\\') j++; else if (ch === q) q = null; }
    else if ('"\'`'.includes(ch)) q = ch;
    else if (ch === o) d++;
    else if (ch === c && --d === 0) break;
  }
  return eval('(' + t.slice(i, j + 1) + ')');
}

const idx = fs.readFileSync('index.html', 'utf8');
const tr = fs.readFileSync('stylist-weekly-tracker.html', 'utf8');
const TC = grab(idx, 'TRAINING_COHORT');
const SD = grab(idx, 'SEED_DAYS');
const B1 = grab(tr, 'BATCH1_RAW');

// Cohort 1's store-scoping overrides live in COHORT_CONFIGS.cohort1.storeOverrides
const ovMatch = idx.match(/storeOverrides:\s*(\{[^}]*\})/);
const overrides = ovMatch ? eval('(' + ovMatch[1] + ')') : {};

const set = new Set(TC.map(s => s.toLowerCase().trim()));
let r = 0, b = 0, u = 0;
for (const d of Object.values(SD)) {
  for (const n of Object.keys(d.stylistStoreRev || {})) {
    const ln = n.toLowerCase().trim();
    if (!set.has(ln)) continue;
    for (const st of Object.keys(d.stylistStoreRev[n])) {
      if (overrides[ln] && overrides[ln] !== st) continue;
      r += d.stylistStoreRev[n][st];
      b += (d.stylistStoreBills[n] || {})[st] || 0;
      u += (d.stylistStoreUnits[n] || {})[st] || 0;
    }
  }
}
let r2 = 0, b2 = 0, u2 = 0;
for (const s of B1) {
  if (!s[period]) throw new Error('BATCH1_RAW entry has no period ' + period + ': ' + s.name);
  r2 += s[period].revenue; b2 += s[period].bills; u2 += s[period].units;
}
const aov1 = (r / b).toFixed(2), aov2 = (r2 / b2).toFixed(2);
console.log('index.html   revenue=' + r.toFixed(0) + ' bills=' + b + ' units=' + u + ' AOV=' + aov1);
console.log('tracker      revenue=' + r2.toFixed(0) + ' bills=' + b2 + ' units=' + u2 + ' AOV=' + aov2);
// The tracker stores each stylist's revenue rounded to the rupee, so allow up to 0.5 per stylist of drift
const tol = 0.5 * B1.length;
if (aov1 === aov2 && b === b2 && u === u2 && Math.abs(r - r2) <= tol) {
  console.log('MATCH');
} else {
  console.log('MISMATCH');
  process.exit(1);
}
