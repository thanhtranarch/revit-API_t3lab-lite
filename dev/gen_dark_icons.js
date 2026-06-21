// Generate icon.dark.svg from icon.svg for every pushbutton/pulldown
// Dark ribbon (~#2D3440): dark fills must become light, light surfaces become dark
// Orange stays orange (it reads on both backgrounds)
const fs = require('fs');
const path = require('path');

const tab = path.join(__dirname, '..', 'T3Lab.extension', 'T3Lab.tab');

// Direction B light palette → dark-ribbon equivalents
const lightToDark = {
  // ── Dark navy fills → light blue (visible on dark ribbon) ──
  '#0D1C2E': '#8AB8CC',
  '#182A3E': '#7AB0CC',
  '#28445E': '#5A90B8',
  '#3A5C7A': '#4A7898',
  // Legacy muted (in case refine_palette hasn't run yet)
  '#3A5268': '#7AB0CC',
  '#4A6278': '#5A90B8',
  '#566A7A': '#4A7898',
  '#2A3A4C': '#90B8CC',
  '#445E70': '#4A7898',

  // ── Orange accent → slightly brighter (both backgrounds work) ──
  '#D07818': '#EAA030',
  '#C87A30': '#EAA030',
  '#B86010': '#D08020',
  '#A86A22': '#D08020',
  '#D08A20': '#EAA030',

  // ── Greens → lighter (more visible on dark ribbon) ──
  '#3A9860': '#50C880',
  '#4AAE88': '#60D090',
  '#2A8850': '#40B070',
  '#3A9A74': '#50C080',

  // ── Reds → lighter ──
  '#D85050': '#F07070',
  '#C86060': '#F07070',
  '#B85252': '#E06060',
  '#A04848': '#D05050',
  '#A83030': '#D05050',
  '#C04040': '#E06060',

  // ── Blue accents → lighter ──
  '#4A78B0': '#7AB8D8',
  '#6890C4': '#7AB8D8',
  '#2E6898': '#60A8CC',
  '#3A7AB8': '#60A8CC',

  // ── Light HEX surface fills → dark surfaces ──
  '#F0F5FC': '#1A2838',
  '#EEF2F7': '#1A2838',
  '#F4F8FC': '#141E2A',
  '#F4F7FA': '#141E2A',
  '#F8FAFC': '#141E2A',
  '#F1F5F9': '#162030',
  '#E2E8F0': '#243848',
  '#CBD5E1': '#2E4258',
  '#D8E4F0': '#243848',
  '#D8E4EF': '#243848',
  '#C8D8E8': '#2E4258',
  '#C8D8EC': '#2E4258',
  '#CCDCEA': '#2E4258',
  '#BCD1EE': '#2E4258',

  // ── Warm surface tints (active row backgrounds) → dark warm ──
  '#FFF4E8': '#2A1C0C',
  '#FFF7F5': '#201510',
  '#FFF3E5': '#201C08',
  '#FFF8F0': '#201C08',
  '#FFFBEB': '#1C1810',
  '#FEF3C7': '#1C1810',

  // ── Passive text / mid-tone elements ──
  '#7A93A8': '#4A6880',
  '#B0BEC5': '#4A6880',
  '#A8B8C8': '#3A5068',
  '#A8C4D8': '#3A5068',
  '#B0C8D8': '#3A5068',
  '#507090': '#3A5878',
};

function transformToDark(svgContent) {
  let result = svgContent;
  for (const [light, dark] of Object.entries(lightToDark)) {
    result = result.replace(new RegExp(light, 'gi'), dark);
  }
  // fill="white" stays white — white elements on dark ribbon have great contrast
  // stroke="white" stays white — line art should remain visible
  return result;
}

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
walkDir(tab, (lightSvgPath) => {
  const content = fs.readFileSync(lightSvgPath, 'utf8');
  const darkContent = transformToDark(content);
  const darkSvgPath = path.join(path.dirname(lightSvgPath), 'icon.dark.svg');
  fs.writeFileSync(darkSvgPath, darkContent, 'utf8');
  console.log('DARK  ' + path.relative(tab, darkSvgPath));
  count++;
});
console.log(`\nDone. ${count} icon.dark.svg files created.`);
