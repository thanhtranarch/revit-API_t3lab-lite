const { Resvg } = require('@resvg/resvg-js');
const fs = require('fs');
const path = require('path');

const tab = path.join(__dirname, '..', 'T3Lab.extension', 'T3Lab.tab');

const targets = [
  { panel: 'Views & Sheets.panel', tools: ['ManaViews.pushbutton', 'ManaSheets.pushbutton', 'BatchOut.pushbutton', 'SheetGen.pushbutton'] },
  { panel: 'Standards & Settings.panel', tools: ['ManaLocations.pushbutton', 'ModelAuditor.pushbutton', 'VisualSettings.pushbutton', 'WorksetManager.pushbutton'] },
  { panel: 'Annotation & Select.panel', tools: ['ManaAnno.pushbutton'] },
  { panel: 'Annotation & Select.panel/Mana.stack', tools: ['ManaAlign.pushbutton', 'ManaDWG.pushbutton', 'ManaSelect.pushbutton'] },
  { panel: 'Modeling & Datum.panel', tools: ['FamilyCreator.pushbutton', 'ManaFamilies.pushbutton', 'ElementAdjust.pulldown'] },
  { panel: 'Modeling & Datum.panel/Create.stack', tools: ['PropertyLine.pushbutton', 'Tile Layout.pushbutton', 'Create Elements.pulldown'] },
  { panel: 'Modeling & Datum.panel/Create.stack/Create Elements.pulldown', tools: ['CADToElements.pushbutton', 'DoorThreshold.pushbutton', 'ImageToDrafting.pushbutton', 'ManaDatums.pushbutton', 'PointCloud.pushbutton', 'RoomToFloor.pushbutton', 'Text to Element.pushbutton'] },
  { panel: 'Modeling & Datum.panel/ElementAdjust.pulldown', tools: ['AutoJoin.pushbutton', 'SplitElements.pushbutton', 'Wall Cut Profile.pushbutton', 'Wall_Adjust Base.pushbutton', 'Split.pulldown'] },
  { panel: 'Modeling & Datum.panel/ElementAdjust.pulldown/Split.pulldown', tools: ['Column_Split.pushbutton', 'Floor_Split.pushbutton', 'Wall_Split.pushbutton'] },
  { panel: 'Support.panel', tools: ['T3LabAssistant.pushbutton', 'MCPControl.pushbutton', 'PDF import.pushbutton', 'Feedback.pushbutton'] },
  { panel: 'Support.panel/CloudLinks.stack', tools: ['Autodesk Forma.urlbutton', 'Autodesk Health.urlbutton', 'Bluebeam Status.urlbutton'] },
  { panel: 'Support.panel/UI.stack', tools: ['BG Theme.pushbutton', 'ManaTabs.pushbutton', 'Ribbon Names.pushbutton'] },
  { panel: 'Data & IFC-SG.panel', tools: ['BCF Reader.pushbutton', 'Foundation Volume.pushbutton', 'IFC-SG.pushbutton'] },
  { panel: 'Data & IFC-SG.panel/manaData.stack', tools: ['ContainsManager.pushbutton', 'ManaParameters.pushbutton', 'ManaSchedules.pushbutton'] },
];

for (const { panel, tools } of targets) {
  for (const tool of tools) {
    const svgPath = path.join(tab, panel, tool, 'icon.svg');
    const pngPath = path.join(tab, panel, tool, 'icon.png');
    if (!fs.existsSync(svgPath)) { console.log('SKIP (no svg) ' + tool); continue; }
    const svg = fs.readFileSync(svgPath);
    const resvg = new Resvg(svg, { fitTo: { mode: 'width', value: 64 } });
    const png = resvg.render().asPng();
    fs.writeFileSync(pngPath, png);
    console.log('OK  ' + tool);
  }
}
console.log('Done.');
