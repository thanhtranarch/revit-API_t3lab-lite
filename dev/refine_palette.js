// Refine icon.svg files from "muted" palette → Direction B (Dark Premium)
// Darker, cleaner navy · amber-orange replacing muddy orange
// Skip CloudLinks.stack (url buttons — locked)
const fs = require('fs');
const path = require('path');

const tab = path.join(__dirname, '..', 'T3Lab.extension', 'T3Lab.tab');

const refineMap = {
  // Orange: muted amber → cleaner warmer amber
  '#C87A30': '#D07818',
  '#A86A22': '#B86010',
  '#D08A20': '#D07818',

  // Dark navy family: push significantly darker for more contrast/drama
  '#3A5268': '#182A3E',
  '#4A6278': '#28445E',
  '#566A7A': '#3A5C7A',
  '#2A3A4C': '#0D1C2E',
  '#445E70': '#3A5C7A',

  // Greens → slightly richer
  '#4AAE88': '#3A9860',
  '#3A9A74': '#2A8850',

  // Reds → slightly more vivid
  '#C86060': '#D85050',
  '#B85252': '#C04040',
  '#A04848': '#A83030',

  // Blue accents → slightly deeper
  '#6890C4': '#4A78B0',
  '#3A7AB8': '#2E6898',
};

function applyRefine(svgContent) {
  let result = svgContent;
  for (const [old, refined] of Object.entries(refineMap)) {
    result = result.replace(new RegExp(old, 'gi'), refined);
  }
  return result;
}

function walkDir(dir, callback) {
  if (!fs.existsSync(dir)) return;
  for (const entry of fs.readdirSync(dir)) {
    const full = path.join(dir, entry);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      if (!full.includes('CloudLinks.stack')) walkDir(full, callback);
    } else if (entry === 'icon.svg') {
      callback(full);
    }
  }
}

let count = 0;
walkDir(tab, (svgPath) => {
  const original = fs.readFileSync(svgPath, 'utf8');
  const refined = applyRefine(original);
  if (refined !== original) {
    fs.writeFileSync(svgPath, refined, 'utf8');
    console.log('REFINED  ' + path.relative(tab, svgPath));
    count++;
  } else {
    console.log('SKIP     ' + path.relative(tab, svgPath));
  }
});
console.log(`\nDone. ${count} files refined to Direction B palette.`);
