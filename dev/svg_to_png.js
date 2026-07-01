const { Resvg } = require('@resvg/resvg-js');
const fs = require('fs');
const path = require('path');

const tab = path.join(__dirname, '..', 'T3Lab.extension', 'T3Lab.tab');

const targets = [
  { panel: 'Views & Sheets.panel', tools: ['ManaViews.pushbutton', 'ManaSheets.pushbutton', 'BatchOut.pushbutton', 'SheetGen.pushbutton'] },
  { panel: 'Standard.panel', tools: ['AutoWork.pushbutton', 'UIShowcase.pushbutton'] },
  { panel: 'Standards & Settings.panel', tools: ['ManaLoca.pushbutton', 'ModelAuditor.pushbutton', 'ManaStyles.pushbutton', 'ManaWorkset.pushbutton'] },
  { panel: 'Annotation & Select.panel', tools: ['ManaAnno.pushbutton'] },
  { panel: 'Annotation & Select.panel/Mana.stack', tools: ['ManaAlign.pushbutton', 'ManaDWG.pushbutton', 'ManaSelect.pushbutton'] },
  { panel: 'Modeling & Datum.panel', tools: ['FamiGen.pushbutton', 'ManaFami.pushbutton', 'ElementAdjust.pulldown'] },
  { panel: 'Modeling & Datum.panel/Create.stack', tools: ['PropertyLine.pushbutton', 'Tile Layout.pushbutton', 'Create Elements.pulldown'] },
  { panel: 'Modeling & Datum.panel/Create.stack/Create Elements.pulldown', tools: ['CADToElements.pushbutton', 'DoorThreshold.pushbutton', 'ImageToDrafting.pushbutton', 'ManaDatums.pushbutton', 'PointCloud.pushbutton', 'RoomToFloor.pushbutton', 'Text to Element.pushbutton'] },
  { panel: 'Modeling & Datum.panel/ElementAdjust.pulldown', tools: ['AutoJoin.pushbutton', 'SplitElements.pushbutton', 'Wall Cut Profile.pushbutton', 'Wall_Adjust Base.pushbutton', 'Split.pulldown'] },
  { panel: 'Modeling & Datum.panel/ElementAdjust.pulldown/Split.pulldown', tools: ['Column_Split.pushbutton', 'Floor_Split.pushbutton', 'Wall_Split.pushbutton'] },
  { panel: 'Support.panel', tools: ['T3LabAssistant.pushbutton', 'MCPControl.pushbutton', 'PDF import.pushbutton', 'Feedback.pushbutton'] },
  { panel: 'Support.panel/CloudLinks.stack', tools: ['Autodesk Forma.urlbutton', 'Autodesk Health.urlbutton', 'Bluebeam Status.urlbutton'] },
  { panel: 'Support.panel/UI.stack', tools: ['BG Theme.pushbutton', 'ManaTabs.pushbutton', 'Ribbon Names.pushbutton'] },
  { panel: 'Data & IFC-SG.panel', tools: ['BCF Reader.pushbutton', 'Foundation Volume.pushbutton', 'IFC-SG.pushbutton'] },
  { panel: 'Data & IFC-SG.panel/manaData.stack', tools: ['ManaContains.pushbutton', 'ManaPara.pushbutton', 'ManaSched.pushbutton'] },
];

function renderSvgToPng(svgPath, pngPath) {
  const svg = fs.readFileSync(svgPath);
  const resvg = new Resvg(svg, { fitTo: { mode: 'width', value: 64 } });
  const png = resvg.render().asPng();
  fs.writeFileSync(pngPath, png);
}

for (const { panel, tools } of targets) {
  for (const tool of tools) {
    const dir = path.join(tab, panel, tool);
    const svgPath = path.join(dir, 'icon.svg');
    const pngPath = path.join(dir, 'icon.png');
    const darkSvgPath = path.join(dir, 'icon.dark.svg');
    const darkPngPath = path.join(dir, 'icon.dark.png');

    if (!fs.existsSync(svgPath)) {
      console.log('SKIP (no svg)        ' + tool);
      continue;
    }

    renderSvgToPng(svgPath, pngPath);

    if (fs.existsSync(darkSvgPath)) {
      renderSvgToPng(darkSvgPath, darkPngPath);
      console.log('OK  [light + dark]   ' + tool);
    } else {
      console.log('OK  [light only]     ' + tool);
    }
  }
}
console.log('Done.');
