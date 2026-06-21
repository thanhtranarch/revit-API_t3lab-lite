// Remap all icon SVG colors to a muted Revit-style palette
const fs = require('fs');
const path = require('path');

const tab = path.join(__dirname, '..', 'T3Lab.extension', 'T3Lab.tab');

// Old vivid → new muted (case-insensitive hex swap)
const colorMap = {
  '#C2440C': '#C87A30',  // vivid orange → muted amber
  '#A83300': '#A86A22',  // dark orange shadow → muted
  '#1E2A3A': '#3A5268',  // very dark navy → mid navy
  '#253344': '#4A6278',  // dark navy → medium blue-gray
  '#334155': '#566A7A',  // slate → lighter blue-gray
  '#2D3F52': '#445E70',  // mid-dark → lighter
  '#0D1825': '#2A3A4C',  // near-black → dark navy
  '#10B981': '#4AAE88',  // vivid emerald → muted green
  '#059669': '#3A9A74',  // emerald hover → muted
  '#EF4444': '#C86060',  // vivid red → muted red
  '#DC2626': '#B85252',  // red hover → muted
  '#B91C1C': '#A04848',  // dark red → muted
  '#3B82F6': '#6890C4',  // vivid blue → muted blue
  '#0077C8': '#3A7AB8',  // bluebeam blue → muted
  '#F59E0B': '#D08A20',  // vivid amber/star → muted
};

function applyMutedColors(svgContent) {
  let result = svgContent;
  for (const [oldColor, newColor] of Object.entries(colorMap)) {
    // Replace both uppercase and lowercase hex
    const regex = new RegExp(oldColor.replace('#', '#?'), 'gi');
    result = result.replace(new RegExp(oldColor, 'gi'), newColor);
  }
  return result;
}

// Walk all icon.svg files under T3Lab.tab
function walkDir(dir, callback) {
  if (!fs.existsSync(dir)) return;
  for (const entry of fs.readdirSync(dir)) {
    const full = path.join(dir, entry);
    const stat = fs.statSync(full);
    if (stat.isDirectory()) {
      walkDir(full, callback);
    } else if (entry === 'icon.svg') {
      callback(full);
    }
  }
}

let count = 0;
walkDir(tab, (svgPath) => {
  const original = fs.readFileSync(svgPath, 'utf8');
  const muted = applyMutedColors(original);
  if (muted !== original) {
    fs.writeFileSync(svgPath, muted, 'utf8');
    console.log('MUTED ' + path.relative(tab, svgPath));
    count++;
  } else {
    console.log('SKIP  ' + path.relative(tab, svgPath));
  }
});
console.log(`\nDone. ${count} files updated.`);
