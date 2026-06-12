# -*- coding: utf-8 -*-
"""
IFC-SG Manual Assignment Tool
Assigns IFC Export parameters according to Singapore BIM standards
"""

__title__ = "Manual Assign\nIFC Class"
__author__ = "Dang Quoc Truong - DQT"
__doc__ = """Manual assignment of IFC Export classes with advanced UI.
Allows filtering, searching, and bulk editing of IFC assignments."""

import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
import System.Windows.Forms
import System.Drawing

from System import Array
from System.Collections.Generic import List
from System.Windows.Forms import (
    Form, Button, Label, TextBox, ComboBox, DataGridView, 
    DataGridViewTextBoxColumn, DataGridViewComboBoxColumn, DataGridViewCheckBoxColumn,
    CheckBox, Panel, SplitContainer, GroupBox,
    DockStyle, AnchorStyles, FormBorderStyle, FormStartPosition,
    DialogResult, MessageBox, MessageBoxButtons, MessageBoxIcon,
    DataGridViewContentAlignment, DataGridViewAutoSizeColumnMode,
    ProgressBar, TabControl, TabPage, CheckedListBox
)
from System.Drawing import Size, Point, Color, Font, FontStyle

from pyrevit import revit, DB, forms, script
from pyrevit.revit import Transaction
import json

# ---------------------------------------------------------------------------
# Blue theme palette (WinForms). In WinForms, Color.FromArgb is the standard,
# safe way to build colors (the BrushConverter rule applies to WPF only).
# ---------------------------------------------------------------------------
DQT_HEADER_GOLD = Color.FromArgb(31, 78, 121)      # #1F4E79 header band (deep blue)
DQT_MAIN_BG     = Color.FromArgb(255, 255, 255)    # #FFFFFF body
DQT_PANEL_BG    = Color.FromArgb(234, 242, 250)    # #EAF2FA light blue panels
DQT_ACCENT      = Color.FromArgb(91, 155, 213)     # #5B9BD5 borders/buttons
DQT_DARK        = Color.FromArgb(31, 78, 121)       # #1F4E79 header/footer/strong text
DQT_TEXT        = Color.FromArgb(51, 51, 51)        # #0F172A body text
DQT_LIGHT_BORDER = Color.FromArgb(208, 224, 240)   # #D0E0F0 separators/gridlines
DQT_ROW_ALT     = Color.FromArgb(221, 235, 247)    # #DDEBF7 alt row / selected
DQT_BTN_HOVER   = Color.FromArgb(189, 215, 238)    # #BDD7EE hover
DQT_HEADER_TEXT = Color.FromArgb(255, 255, 255)    # white text on blue header

DQT_FOOTER_TEXT = "Dang Quoc Truong - DQT (c) 2026"


def _style_dqt_button(btn, primary=False):
    """Apply DQT styling to a WinForms button."""
    btn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
    btn.FlatAppearance.BorderColor = DQT_ACCENT
    btn.FlatAppearance.BorderSize = 1
    btn.FlatAppearance.MouseOverBackColor = DQT_BTN_HOVER
    btn.FlatAppearance.MouseDownBackColor = DQT_ACCENT
    if primary:
        btn.BackColor = DQT_HEADER_GOLD
        btn.ForeColor = DQT_HEADER_TEXT
        btn.Font = Font("Segoe UI", 10, FontStyle.Bold)
    else:
        btn.BackColor = DQT_MAIN_BG
        btn.ForeColor = DQT_DARK

# Get the current document
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


def _eid_int(eid):
    """Get integer value of ElementId. Revit 2024+: .Value, older: .IntegerValue"""
    try:
        return eid.Value
    except:
        try:
            return eid.IntegerValue
        except:
            return 0


class _SilentFailures(DB.IFailuresPreprocessor):
    """Swallow warnings so the IFC assignment transaction is not interrupted."""
    def PreprocessFailures(self, fa):
        for f in fa.GetFailureMessages():
            if f.GetSeverity() == DB.FailureSeverity.Warning:
                fa.DeleteWarning(f)
        return DB.FailureProcessingResult.Continue

# IFC-SG Mapping Dictionary (Comprehensive based on CX Pilot Mapping)
IFC_SG_MAPPING = {
    "Walls": [
        {"entity": "IfcWall", "subtype": "", "desc": "Standard Wall"},
        {"entity": "IfcWall", "subtype": "PARAPET", "desc": "Parapet Wall"},
        {"entity": "IfcWall", "subtype": "RETAININGWALL", "desc": "Retaining Wall"},
        {"entity": "IfcWall", "subtype": "*BOUNDARYWALL", "desc": "Boundary Wall"},
        {"entity": "IfcWall", "subtype": "*REFUSECHUTE", "desc": "Refuse Chute"},
    ],
    "Curtain Walls": [
        {"entity": "IfcCurtainWall", "subtype": "", "desc": "Curtain Wall System"},
    ],
    "Floors": [
        {"entity": "IfcSlab", "subtype": "", "desc": "Standard Slab/Floor"},
        {"entity": "IfcSlab", "subtype": "*ACCESSIBLEROUTE", "desc": "Accessible Route"},
        {"entity": "IfcCovering", "subtype": "FLOORING", "desc": "Floor Covering"},
        {"entity": "IfcCivilElement", "subtype": "*ACCESSIBLEROUTE", "desc": "Civil - Accessible Route"},
        {"entity": "IfcCivilElement", "subtype": "*FOOTPATH", "desc": "Footpath"},
        {"entity": "IfcCivilElement", "subtype": "*DRIVEWAY", "desc": "Driveway"},
        {"entity": "IfcCivilElement", "subtype": "*CARRIAGEWAY", "desc": "Carriageway"},
        {"entity": "IfcCivilElement", "subtype": "*ROADKERB", "desc": "Road Kerb"},
    ],
    "Roofs": [
        {"entity": "IfcRoof", "subtype": "", "desc": "Standard Roof"},
        {"entity": "IfcSlab", "subtype": "ROOF", "desc": "Roof Slab"},
        {"entity": "IfcCovering", "subtype": "ROOFING", "desc": "Roof Covering"},
        {"entity": "IfcCovering", "subtype": "*SOFFIT", "desc": "Soffit"},
    ],
    "Ceilings": [
        {"entity": "IfcCovering", "subtype": "CEILING", "desc": "Ceiling"},
    ],
    "Doors": [
        {"entity": "IfcDoor", "subtype": "", "desc": "Standard Door"},
        {"entity": "IfcDoor", "subtype": "*BLASTDOOR", "desc": "Blast Door"},
        {"entity": "IfcDoor", "subtype": "*ROLLERSHUTTER", "desc": "Roller Shutter"},
        {"entity": "IfcDoor", "subtype": "*OPENING", "desc": "Door Opening"},
        {"entity": "IfcDoor", "subtype": "*ACCESSHATCH", "desc": "Access Hatch"},
        {"entity": "IfcDoor", "subtype": "*RECYCLABLESCHUTEACCESSPANEL", "desc": "Recyclable Chute Access Panel"},
        {"entity": "IfcDoor", "subtype": "*RECYCLABLESCHUTEHOPPER", "desc": "Recyclable Chute Hopper"},
        {"entity": "IfcDoor", "subtype": "*REFUSECHUTEACCESSPANEL", "desc": "Refuse Chute Access Panel"},
        {"entity": "IfcDoor", "subtype": "*REFUSECHUTEHOPPER", "desc": "Refuse Chute Hopper"},
    ],
    "Windows": [
        {"entity": "IfcWindow", "subtype": "", "desc": "Standard Window"},
        {"entity": "IfcWindow", "subtype": "WINDOW", "desc": "Window"},
        {"entity": "IfcWindow", "subtype": "SKYLIGHT", "desc": "Skylight"},
        {"entity": "IfcWindow", "subtype": "LOUVRE", "desc": "Louvre Window"},
        {"entity": "IfcWindow", "subtype": "*OPENING", "desc": "Window Opening"},
        {"entity": "IfcWindow", "subtype": "*BAYWINDOW", "desc": "Bay Window"},
        {"entity": "IfcWindow", "subtype": "*VENTILATIONSLEEVE", "desc": "Ventilation Sleeve"},
    ],
    "Columns": [
        {"entity": "IfcColumn", "subtype": "", "desc": "Column"},
    ],
    "Structural Columns": [
        {"entity": "IfcColumn", "subtype": "", "desc": "Structural Column"},
    ],
    "Structural Framing": [
        {"entity": "IfcBeam", "subtype": "", "desc": "Beam"},
    ],
    "Structural Foundations": [
        {"entity": "IfcFooting", "subtype": "", "desc": "Footing"},
        {"entity": "IfcPile", "subtype": "", "desc": "Pile"},
    ],
    "Stairs": [
        {"entity": "IfcStair", "subtype": "", "desc": "Stair"},
        {"entity": "IfcStairFlight", "subtype": "", "desc": "Stair Flight"},
    ],
    "Ramps": [
        {"entity": "IfcRamp", "subtype": "*ACCESSIBLEROUTE", "desc": "Accessible Ramp"},
        {"entity": "IfcRamp", "subtype": "*CURVEDRAMP", "desc": "Curved Ramp"},
        {"entity": "IfcRamp", "subtype": "*FLAREDKERBRAMP", "desc": "Flared Kerb Ramp"},
        {"entity": "IfcRamp", "subtype": "STRAIGHT_RUN_RAMP", "desc": "Straight Run Ramp"},
    ],
    "Railings": [
        {"entity": "IfcRailing", "subtype": "GUARDRAIL", "desc": "Guard Rail"},
        {"entity": "IfcRailing", "subtype": "*BOLLARD", "desc": "Bollard"},
    ],
    "Rooms": [
        {"entity": "IfcSpace", "subtype": "", "desc": "Space/Room"},
        {"entity": "IfcSpace", "subtype": "SPACE", "desc": "Standard Space"},
    ],
    "Areas": [
        {"entity": "IfcSpace", "subtype": "SPACE", "desc": "Area Space"},
        {"entity": "IfcSpace", "subtype": "*ACCESSIBLEROUTE", "desc": "Accessible Route Space"},
        {"entity": "IfcSpace", "subtype": "*ACCESSWAY", "desc": "Access Way"},
        {"entity": "IfcSpace", "subtype": "*PARKINGACCESSWAY", "desc": "Parking Access Way"},
        {"entity": "IfcSpace", "subtype": "*FIREENGINEACCESSROAD", "desc": "Fire Engine Access Road"},
        {"entity": "IfcSpace", "subtype": "*FIREENGINEACCESSWAY", "desc": "Fire Engine Access Way"},
        {"entity": "IfcSpace", "subtype": "*VEHICULARSERVICEROAD", "desc": "Vehicular Service Road"},
        {"entity": "IfcSpace", "subtype": "*AREA_CONNECTIVITY", "desc": "Area Connectivity"},
        {"entity": "IfcSpace", "subtype": "*AREA_GFA", "desc": "GFA Area"},
        {"entity": "IfcSpace", "subtype": "*AREA_LANDSCAPE", "desc": "Landscape Area"},
        {"entity": "IfcSpace", "subtype": "*AREA_STRATA", "desc": "Strata Area"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*DRIVEWAY", "desc": "Driveway Proxy"},
    ],
    "Furniture": [
        {"entity": "IfcFurniture", "subtype": "", "desc": "Standard Furniture"},
        {"entity": "IfcFurniture", "subtype": "CHAIR", "desc": "Chair"},
        {"entity": "IfcFurniture", "subtype": "*BENCH", "desc": "Bench"},
        {"entity": "IfcFurniture", "subtype": "*CHANGINGBED", "desc": "Changing Bed"},
        {"entity": "IfcFurniture", "subtype": "*CHILDPROTECTIONSEAT", "desc": "Child Protection Seat"},
        {"entity": "IfcFurniture", "subtype": "*DIAPERCHANGINGTABLE", "desc": "Diaper Changing Table"},
        {"entity": "IfcFurniture", "subtype": "*PLANTERBOX", "desc": "Planter Box"},
        {"entity": "IfcFurniture", "subtype": "*RACK", "desc": "Rack"},
    ],
    "Generic Models": [
        {"entity": "IfcBuildingElementProxy", "subtype": "", "desc": "Generic Proxy"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*ACCESSIBLEROUTE", "desc": "Accessible Route Proxy"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*BOREHOLE", "desc": "Borehole"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*TACTILETILE", "desc": "Tactile Tile"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*PORTABLEFIREEXTINGUISHER", "desc": "Portable Fire Extinguisher"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*CARLOT", "desc": "Car Lot"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*MOTOR-CYCLELOT", "desc": "Motorcycle Lot"},
        {"entity": "IfcBuildingElementProxy", "subtype": "*SIGNAGE_EXIT", "desc": "Exit Signage"},
        {"entity": "IfcCivilElement", "subtype": "*CULVERT", "desc": "Culvert"},
        {"entity": "IfcCivilElement", "subtype": "*ENTRANCECULVERT", "desc": "Entrance Culvert"},
        {"entity": "IfcCivilElement", "subtype": "*CROSSCULVERT", "desc": "Cross Culvert"},
        {"entity": "IfcCovering", "subtype": "CLADDING", "desc": "Cladding"},
        {"entity": "IfcCovering", "subtype": "*FIRECURTAIN", "desc": "Fire Curtain"},
    ],
    "Lighting Fixtures": [
        {"entity": "IfcLightFixture", "subtype": "", "desc": "Light Fixture"},
        {"entity": "IfcLightFixture", "subtype": "SECURITYLIGHTING", "desc": "Security Lighting"},
    ],
    "Plumbing Fixtures": [
        {"entity": "IfcSanitaryTerminal", "subtype": "BATH", "desc": "Bath"},
        {"entity": "IfcSanitaryTerminal", "subtype": "BIDET", "desc": "Bidet"},
        {"entity": "IfcSanitaryTerminal", "subtype": "SHOWER", "desc": "Shower"},
        {"entity": "IfcSanitaryTerminal", "subtype": "URINAL", "desc": "Urinal"},
        {"entity": "IfcSanitaryTerminal", "subtype": "WASHHANDBASIN", "desc": "Wash Hand Basin"},
        {"entity": "IfcSanitaryTerminal", "subtype": "*WATERCLOSET", "desc": "Water Closet"},
        {"entity": "IfcDistributionChamberElement", "subtype": "INSPECTIONCHAMBER", "desc": "Inspection Chamber"},
        {"entity": "IfcDistributionChamberElement", "subtype": "MANHOLE", "desc": "Manhole"},
        {"entity": "IfcDistributionChamberElement", "subtype": "SUMP", "desc": "Sump"},
        {"entity": "IfcFireSuppressionTerminal", "subtype": "BREECHINGINLET", "desc": "Breeching Inlet"},
        {"entity": "IfcFireSuppressionTerminal", "subtype": "FIREHYDRANT", "desc": "Fire Hydrant"},
    ],
    "Mechanical Equipment": [
        {"entity": "IfcPump", "subtype": "", "desc": "Pump"},
        {"entity": "IfcPump", "subtype": "SUMPPUMP", "desc": "Sump Pump"},
        {"entity": "IfcTank", "subtype": "STORAGE", "desc": "Storage Tank"},
        {"entity": "IfcTank", "subtype": "VESSEL", "desc": "Vessel"},
        {"entity": "IfcFireSuppressionTerminal", "subtype": "HOSEREEL", "desc": "Hose Reel"},
        {"entity": "IfcTransportElement", "subtype": "ESCALATOR", "desc": "Escalator"},
    ],
    "Specialty Equipment": [
        {"entity": "IfcTransportElement", "subtype": "*LIFT", "desc": "Lift/Elevator"},
        {"entity": "IfcTransportElement", "subtype": "*CARLIFT", "desc": "Car Lift"},
    ],
    "Pipes": [
        {"entity": "IfcPipeSegment", "subtype": "RIGIDSEGMENT", "desc": "Rigid Pipe Segment"},
        {"entity": "IfcPipeSegment", "subtype": "GUTTER", "desc": "Gutter"},
        {"entity": "IfcPipeSegment", "subtype": "SPOOL", "desc": "Pipe Spool"},
    ],
    "Pipe Fittings": [
        {"entity": "IfcPipeFitting", "subtype": "", "desc": "Pipe Fitting"},
        {"entity": "IfcPipeFitting", "subtype": "BEND", "desc": "Bend"},
        {"entity": "IfcPipeFitting", "subtype": "JUNCTION", "desc": "Junction"},
    ],
    "Pipe Accessories": [
        {"entity": "IfcValve", "subtype": "ISOLATING", "desc": "Isolating Valve"},
        {"entity": "IfcValve", "subtype": "CHECK", "desc": "Check Valve"},
        {"entity": "IfcFlowMeter", "subtype": "WATERMETER", "desc": "Water Meter"},
    ],
    "Ducts": [
        {"entity": "IfcDuctSegment", "subtype": "", "desc": "Duct Segment"},
    ],
    "Duct Fittings": [
        {"entity": "IfcDuctFitting", "subtype": "", "desc": "Duct Fitting"},
    ],
    "Duct Accessories": [
        {"entity": "IfcDamper", "subtype": "FIREDAMPER", "desc": "Fire Damper"},
        {"entity": "IfcDamper", "subtype": "FIRESMOKEDAMPER", "desc": "Fire Smoke Damper"},
        {"entity": "IfcDamper", "subtype": "SMOKEDAMPER", "desc": "Smoke Damper"},
    ],
    "Topography": [
        {"entity": "IfcGeographicElement", "subtype": "TERRAIN", "desc": "Terrain"},
        {"entity": "IfcGeographicElement", "subtype": "*EXISTINGEARTHWORKS", "desc": "Existing Earthworks"},
        {"entity": "IfcGeographicElement", "subtype": "*PROPOSEDEARTHWORKS", "desc": "Proposed Earthworks"},
    ],
    "Planting": [
        {"entity": "IfcGeographicElement", "subtype": "*LANDSCAPE_TREE", "desc": "Tree"},
        {"entity": "IfcGeographicElement", "subtype": "*LANDSCAPE_PALM", "desc": "Palm"},
        {"entity": "IfcGeographicElement", "subtype": "*LANDSCAPE_SHRUBS", "desc": "Shrubs"},
    ],
}


class ElementData:
    """Class to hold element data for the grid"""
    def __init__(self, element, category, name, current_ifc):
        self.element = element
        self.element_id = element.Id
        self.category = category
        self.name = name
        self.current_ifc = current_ifc
        self.new_ifc = current_ifc  # Will be modified by user
        self.is_selected = True


class IFCAssignmentForm(Form):
    """Advanced UI for manual IFC assignment"""
    
    def __init__(self, elements_data):
        self.elements_data = elements_data
        self.filtered_data = list(elements_data)
        self._loading = False
        self.InitializeComponent()
        self.LoadData()
        
    def InitializeComponent(self):
        """Initialize all UI components"""
        self.Text = "IFC-SG Manual Assignment Tool - DQT"
        self.Size = Size(1400, 820)
        self.StartPosition = FormStartPosition.CenterScreen
        self.MinimumSize = Size(1200, 620)
        self.BackColor = DQT_MAIN_BG

        # Create main container
        main_panel = Panel()
        main_panel.Dock = DockStyle.Fill
        main_panel.Padding = System.Windows.Forms.Padding(10)
        main_panel.BackColor = DQT_MAIN_BG

        # Bottom panel for statistics and actions
        self.CreateBottomPanel(main_panel)

        # Middle panel for data grid
        self.CreateDataGrid(main_panel)

        # Top panel for filters and search
        self.CreateFilterPanel(main_panel)

        # Gold header band (added last with Dock=Top so it sits on top)
        self.CreateHeaderBand(main_panel)

        self.Controls.Add(main_panel)

    def CreateHeaderBand(self, parent):
        """Create the gold DQT header band with accent underline."""
        header = Panel()
        header.Height = 52
        header.Dock = DockStyle.Top
        header.BackColor = DQT_HEADER_GOLD

        title = Label()
        title.Text = "IFC-SG Manual Assignment"
        title.Location = Point(12, 8)
        title.Size = Size(600, 30)
        title.ForeColor = DQT_HEADER_TEXT
        title.Font = Font("Segoe UI", 14, FontStyle.Bold)
        title.BackColor = Color.Transparent
        header.Controls.Add(title)

        subtitle = Label()
        subtitle.Text = "Singapore BIM IFC Export class assignment"
        subtitle.Location = Point(15, 33)
        subtitle.Size = Size(600, 16)
        subtitle.ForeColor = DQT_HEADER_TEXT
        subtitle.Font = Font("Segoe UI", 8, FontStyle.Italic)
        subtitle.BackColor = Color.Transparent
        header.Controls.Add(subtitle)

        # Accent underline strip
        underline = Panel()
        underline.Height = 3
        underline.Dock = DockStyle.Bottom
        underline.BackColor = DQT_ACCENT
        header.Controls.Add(underline)

        parent.Controls.Add(header)
        
    def CreateFilterPanel(self, parent):
        """Create filter and search panel"""
        filter_panel = Panel()
        filter_panel.Height = 120
        filter_panel.Dock = DockStyle.Top
        filter_panel.BorderStyle = getattr(System.Windows.Forms.BorderStyle, "None")
        filter_panel.BackColor = DQT_PANEL_BG
        filter_panel.Padding = System.Windows.Forms.Padding(8)
        
        # Search box
        search_label = Label()
        search_label.Text = "Search (Name/ID):"
        search_label.Location = Point(10, 15)
        search_label.Size = Size(120, 20)
        search_label.ForeColor = DQT_DARK
        filter_panel.Controls.Add(search_label)
        
        self.search_box = TextBox()
        self.search_box.Location = Point(140, 12)
        self.search_box.Size = Size(200, 20)
        self.search_box.TextChanged += self.OnSearchChanged
        filter_panel.Controls.Add(self.search_box)
        
        # Category filter
        category_label = Label()
        category_label.Text = "Category Filter:"
        category_label.Location = Point(360, 15)
        category_label.Size = Size(100, 20)
        category_label.ForeColor = DQT_DARK
        filter_panel.Controls.Add(category_label)
        
        self.category_combo = ComboBox()
        self.category_combo.Location = Point(470, 12)
        self.category_combo.Size = Size(200, 20)
        self.category_combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        
        # Populate categories
        categories = sorted(set([elem.category for elem in self.elements_data]))
        self.category_combo.Items.Add("All Categories")
        for cat in categories:
            self.category_combo.Items.Add(cat)
        self.category_combo.SelectedIndex = 0
        self.category_combo.SelectedIndexChanged += self.OnFilterChanged
        filter_panel.Controls.Add(self.category_combo)
        
        # IFC Entity filter
        ifc_label = Label()
        ifc_label.Text = "Current IFC Filter:"
        ifc_label.Location = Point(690, 15)
        ifc_label.Size = Size(110, 20)
        ifc_label.ForeColor = DQT_DARK
        filter_panel.Controls.Add(ifc_label)
        
        self.ifc_combo = ComboBox()
        self.ifc_combo.Location = Point(810, 12)
        self.ifc_combo.Size = Size(200, 20)
        self.ifc_combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        
        # Populate IFC entities
        ifc_entities = sorted(set([elem.current_ifc for elem in self.elements_data if elem.current_ifc]))
        self.ifc_combo.Items.Add("All IFC Types")
        self.ifc_combo.Items.Add("(Not Assigned)")
        for ifc in ifc_entities:
            if ifc:
                self.ifc_combo.Items.Add(ifc)
        self.ifc_combo.SelectedIndex = 0
        self.ifc_combo.SelectedIndexChanged += self.OnFilterChanged
        filter_panel.Controls.Add(self.ifc_combo)
        
        # Bulk assignment section
        bulk_label = Label()
        bulk_label.Text = "Bulk Assign to Selected:"
        bulk_label.Location = Point(10, 50)
        bulk_label.Size = Size(150, 20)
        bulk_label.Font = Font(bulk_label.Font, FontStyle.Bold)
        bulk_label.ForeColor = DQT_DARK
        filter_panel.Controls.Add(bulk_label)
        
        self.bulk_ifc_combo = ComboBox()
        self.bulk_ifc_combo.Location = Point(170, 47)
        self.bulk_ifc_combo.Size = Size(300, 20)
        self.bulk_ifc_combo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        filter_panel.Controls.Add(self.bulk_ifc_combo)
        
        bulk_apply_btn = Button()
        bulk_apply_btn.Text = "Apply to Selected"
        bulk_apply_btn.Location = Point(480, 45)
        bulk_apply_btn.Size = Size(120, 25)
        bulk_apply_btn.Click += self.OnBulkApply
        _style_dqt_button(bulk_apply_btn, primary=True)
        filter_panel.Controls.Add(bulk_apply_btn)
        
        # Action buttons
        select_all_btn = Button()
        select_all_btn.Text = "Select All"
        select_all_btn.Location = Point(620, 45)
        select_all_btn.Size = Size(90, 25)
        select_all_btn.Click += self.OnSelectAll
        _style_dqt_button(select_all_btn)
        filter_panel.Controls.Add(select_all_btn)
        
        deselect_all_btn = Button()
        deselect_all_btn.Text = "Deselect All"
        deselect_all_btn.Location = Point(720, 45)
        deselect_all_btn.Size = Size(90, 25)
        deselect_all_btn.Click += self.OnDeselectAll
        _style_dqt_button(deselect_all_btn)
        filter_panel.Controls.Add(deselect_all_btn)
        
        clear_filter_btn = Button()
        clear_filter_btn.Text = "Clear Filters"
        clear_filter_btn.Location = Point(1030, 12)
        clear_filter_btn.Size = Size(100, 25)
        clear_filter_btn.Click += self.OnClearFilters
        _style_dqt_button(clear_filter_btn)
        filter_panel.Controls.Add(clear_filter_btn)
        
        # Instructions
        instructions = Label()
        instructions.Text = "Instructions: Search and filter elements, select rows, choose IFC class from dropdown, or use bulk assignment."
        instructions.Location = Point(10, 85)
        instructions.Size = Size(1000, 20)
        instructions.ForeColor = DQT_DARK
        filter_panel.Controls.Add(instructions)
        
        parent.Controls.Add(filter_panel)
        
    def CreateDataGrid(self, parent):
        """Create the data grid view"""
        grid_panel = Panel()
        grid_panel.Dock = DockStyle.Fill
        grid_panel.Padding = System.Windows.Forms.Padding(0, 5, 0, 5)
        
        self.data_grid = DataGridView()
        self.data_grid.Dock = DockStyle.Fill
        self.data_grid.AllowUserToAddRows = False
        self.data_grid.AllowUserToDeleteRows = False
        self.data_grid.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        self.data_grid.MultiSelect = True
        self.data_grid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        self.data_grid.RowHeadersVisible = True

        # DQT grid theming
        self.data_grid.BackgroundColor = DQT_MAIN_BG
        self.data_grid.BorderStyle = getattr(System.Windows.Forms.BorderStyle, "None")
        self.data_grid.GridColor = DQT_LIGHT_BORDER
        self.data_grid.EnableHeadersVisualStyles = False
        self.data_grid.ColumnHeadersDefaultCellStyle.BackColor = DQT_HEADER_GOLD
        self.data_grid.ColumnHeadersDefaultCellStyle.ForeColor = DQT_HEADER_TEXT
        self.data_grid.ColumnHeadersDefaultCellStyle.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.data_grid.ColumnHeadersHeight = 30
        self.data_grid.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single
        self.data_grid.DefaultCellStyle.BackColor = DQT_MAIN_BG
        self.data_grid.DefaultCellStyle.ForeColor = DQT_TEXT
        self.data_grid.DefaultCellStyle.SelectionBackColor = DQT_ROW_ALT
        self.data_grid.DefaultCellStyle.SelectionForeColor = DQT_DARK
        self.data_grid.AlternatingRowsDefaultCellStyle.BackColor = DQT_PANEL_BG
        self.data_grid.RowHeadersDefaultCellStyle.BackColor = DQT_PANEL_BG
        
        # Add columns
        # Select column (checkbox)
        select_col = DataGridViewCheckBoxColumn()
        select_col.Name = "Select"
        select_col.HeaderText = "Select"
        select_col.Width = 60
        select_col.FillWeight = 5
        self.data_grid.Columns.Add(select_col)
        
        # Element ID
        id_col = DataGridViewTextBoxColumn()
        id_col.Name = "ElementID"
        id_col.HeaderText = "Element ID"
        id_col.ReadOnly = True
        id_col.Width = 100
        id_col.FillWeight = 10
        self.data_grid.Columns.Add(id_col)
        
        # Category
        cat_col = DataGridViewTextBoxColumn()
        cat_col.Name = "Category"
        cat_col.HeaderText = "Category"
        cat_col.ReadOnly = True
        cat_col.Width = 150
        cat_col.FillWeight = 15
        self.data_grid.Columns.Add(cat_col)
        
        # Element Name
        name_col = DataGridViewTextBoxColumn()
        name_col.Name = "ElementName"
        name_col.HeaderText = "Element Name / Type"
        name_col.ReadOnly = True
        name_col.Width = 250
        name_col.FillWeight = 25
        self.data_grid.Columns.Add(name_col)
        
        # Current IFC
        current_ifc_col = DataGridViewTextBoxColumn()
        current_ifc_col.Name = "CurrentIFC"
        current_ifc_col.HeaderText = "Current IFC Export As"
        current_ifc_col.ReadOnly = True
        current_ifc_col.Width = 200
        current_ifc_col.FillWeight = 20
        self.data_grid.Columns.Add(current_ifc_col)
        
        # New IFC (ComboBox - editable)
        new_ifc_col = DataGridViewComboBoxColumn()
        new_ifc_col.Name = "NewIFC"
        new_ifc_col.HeaderText = "New IFC Export As"
        new_ifc_col.Width = 250
        new_ifc_col.FillWeight = 25
        new_ifc_col.DisplayStyle = System.Windows.Forms.DataGridViewComboBoxDisplayStyle.DropDownButton
        self.data_grid.Columns.Add(new_ifc_col)
        
        # Event handler for cell value change
        self.data_grid.CellValueChanged += self.OnCellValueChanged
        self.data_grid.CurrentCellDirtyStateChanged += self.OnCurrentCellDirtyStateChanged
        self.data_grid.SelectionChanged += self.OnSelectionChanged
        # Suppress the default DataGridView error dialog (e.g. combo value not valid)
        self.data_grid.DataError += self.OnGridDataError
        
        grid_panel.Controls.Add(self.data_grid)
        parent.Controls.Add(grid_panel)
        
    def CreateBottomPanel(self, parent):
        """Create bottom panel with statistics and action buttons"""
        bottom_panel = Panel()
        bottom_panel.Height = 110
        bottom_panel.Dock = DockStyle.Bottom
        bottom_panel.BorderStyle = getattr(System.Windows.Forms.BorderStyle, "None")
        bottom_panel.BackColor = DQT_PANEL_BG

        # Accent top border strip
        top_strip = Panel()
        top_strip.Height = 3
        top_strip.Dock = DockStyle.Top
        top_strip.BackColor = DQT_ACCENT
        bottom_panel.Controls.Add(top_strip)

        # Statistics
        self.stats_label = Label()
        self.stats_label.Location = Point(10, 12)
        self.stats_label.Size = Size(800, 20)
        self.stats_label.Font = Font("Segoe UI", 9, FontStyle.Bold)
        self.stats_label.ForeColor = DQT_DARK
        bottom_panel.Controls.Add(self.stats_label)

        self.selection_label = Label()
        self.selection_label.Location = Point(10, 35)
        self.selection_label.Size = Size(800, 20)
        self.selection_label.ForeColor = DQT_TEXT
        bottom_panel.Controls.Add(self.selection_label)

        # Action buttons
        apply_btn = Button()
        apply_btn.Text = "Apply Changes to Revit"
        apply_btn.Location = Point(900, 15)
        apply_btn.Size = Size(180, 35)
        apply_btn.Click += self.OnApplyChanges
        _style_dqt_button(apply_btn, primary=True)
        bottom_panel.Controls.Add(apply_btn)

        export_btn = Button()
        export_btn.Text = "Export Report"
        export_btn.Location = Point(1090, 15)
        export_btn.Size = Size(120, 35)
        export_btn.Click += self.OnExportReport
        _style_dqt_button(export_btn)
        bottom_panel.Controls.Add(export_btn)

        cancel_btn = Button()
        cancel_btn.Text = "Close"
        cancel_btn.Location = Point(1220, 15)
        cancel_btn.Size = Size(120, 35)
        cancel_btn.Click += self.OnClose
        _style_dqt_button(cancel_btn)
        bottom_panel.Controls.Add(cancel_btn)

        # Instructions
        help_label = Label()
        help_label.Text = "Tip: Click on 'New IFC Export As' dropdown to change IFC class for individual elements"
        help_label.Location = Point(10, 62)
        help_label.Size = Size(800, 20)
        help_label.ForeColor = DQT_TEXT
        help_label.Font = Font("Segoe UI", 8, FontStyle.Italic)
        bottom_panel.Controls.Add(help_label)

        # DQT copyright footer
        footer_label = Label()
        footer_label.Text = DQT_FOOTER_TEXT
        footer_label.Location = Point(10, 86)
        footer_label.Size = Size(1330, 18)
        footer_label.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        footer_label.ForeColor = DQT_DARK
        footer_label.Font = Font("Segoe UI", 8, FontStyle.Regular)
        footer_label.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
        bottom_panel.Controls.Add(footer_label)

        parent.Controls.Add(bottom_panel)
        
    def LoadData(self):
        """Load element data into the grid"""
        self._loading = True
        try:
            self.data_grid.Rows.Clear()

            for elem_data in self.filtered_data:
                row_idx = self.data_grid.Rows.Add()
                row = self.data_grid.Rows[row_idx]

                # Store reference FIRST so any event handler has a valid Tag
                row.Tag = elem_data

                # Populate New IFC dropdown based on category BEFORE setting value
                new_ifc_cell = row.Cells["NewIFC"]
                ifc_options = self.GetIFCOptionsForCategory(elem_data.category)
                for option in ifc_options:
                    new_ifc_cell.Items.Add(option)

                # Set cell values
                row.Cells["Select"].Value = elem_data.is_selected
                row.Cells["ElementID"].Value = str(_eid_int(elem_data.element_id))
                row.Cells["Category"].Value = elem_data.category
                row.Cells["ElementName"].Value = elem_data.name
                row.Cells["CurrentIFC"].Value = elem_data.current_ifc if elem_data.current_ifc else "(Not Assigned)"

                # Set default value only if it is a valid item in the combo
                desired = elem_data.new_ifc
                if desired and new_ifc_cell.Items.Contains(desired):
                    new_ifc_cell.Value = desired
                elif len(ifc_options) > 0:
                    new_ifc_cell.Value = ifc_options[0]

            # Populate bulk assignment dropdown
            self.PopulateBulkIFCCombo()
        finally:
            self._loading = False

        self.UpdateStatistics()
        
    def GetIFCOptionsForCategory(self, category):
        """Get IFC options for a specific category"""
        options = ["(Keep Current)", "(Clear)"]
        
        # Check if category exists in mapping
        if category in IFC_SG_MAPPING:
            for mapping in IFC_SG_MAPPING[category]:
                entity = mapping["entity"]
                subtype = mapping["subtype"]
                desc = mapping["desc"]
                
                if subtype:
                    option = "{}.{}".format(entity, subtype)
                else:
                    option = entity
                
                # Add description in tooltip format
                full_option = "{} [{}]".format(option, desc) if desc else option
                options.append(full_option)
        
        # Also add some common generic options
        options.extend([
            "IfcBuildingElementProxy",
            "IfcElement",
        ])
        
        return options
        
    def PopulateBulkIFCCombo(self):
        """Populate bulk IFC assignment combo"""
        self.bulk_ifc_combo.Items.Clear()
        self.bulk_ifc_combo.Items.Add("-- Select IFC Class --")
        
        # Get unique IFC options from all categories
        all_options = set()
        for category in IFC_SG_MAPPING.keys():
            for mapping in IFC_SG_MAPPING[category]:
                entity = mapping["entity"]
                subtype = mapping["subtype"]
                
                if subtype:
                    option = "{}.{}".format(entity, subtype)
                else:
                    option = entity
                
                all_options.add(option)
        
        for option in sorted(all_options):
            self.bulk_ifc_combo.Items.Add(option)
        
        self.bulk_ifc_combo.SelectedIndex = 0
        
    def OnSearchChanged(self, sender, args):
        """Handle search text change"""
        self.ApplyFilters()
        
    def OnFilterChanged(self, sender, args):
        """Handle filter change"""
        self.ApplyFilters()
        
    def ApplyFilters(self):
        """Apply search and filters to data"""
        search_text = self.search_box.Text.lower()
        selected_category = self.category_combo.SelectedItem
        selected_ifc = self.ifc_combo.SelectedItem
        
        self.filtered_data = []
        
        for elem_data in self.elements_data:
            # Apply search filter
            if search_text:
                if (search_text not in elem_data.name.lower() and 
                    search_text not in str(_eid_int(elem_data.element_id))):
                    continue
            
            # Apply category filter
            if selected_category != "All Categories":
                if elem_data.category != selected_category:
                    continue
            
            # Apply IFC filter
            if selected_ifc == "(Not Assigned)":
                if elem_data.current_ifc:
                    continue
            elif selected_ifc != "All IFC Types":
                if elem_data.current_ifc != selected_ifc:
                    continue
            
            self.filtered_data.append(elem_data)
        
        self.LoadData()
        
    def OnClearFilters(self, sender, args):
        """Clear all filters"""
        self.search_box.Text = ""
        self.category_combo.SelectedIndex = 0
        self.ifc_combo.SelectedIndex = 0
        self.ApplyFilters()
        
    def _set_combo_cell_value(self, cell, ifc_value):
        """Safely set a DataGridViewComboBoxCell value.

        The cell stores items formatted as 'IfcEntity.SUBTYPE [Description]'.
        We must match against an existing item; if none matches we add the
        plain value as a new item first, otherwise DataGridView raises
        'value is not valid'.
        """
        # Find an existing item whose IFC part matches ifc_value
        for item in cell.Items:
            item_value = str(item).split(" [")[0]
            if item_value == ifc_value:
                cell.Value = item
                return
        # No match: add the plain value as a valid item, then select it
        cell.Items.Add(ifc_value)
        cell.Value = ifc_value

    def OnBulkApply(self, sender, args):
        """Apply bulk IFC assignment to selected rows"""
        selected_ifc = self.bulk_ifc_combo.SelectedItem
        
        if not selected_ifc or selected_ifc == "-- Select IFC Class --":
            MessageBox.Show("Please select an IFC class to apply.", 
                          "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            return
        
        # Extract just the IFC class without description
        ifc_value = str(selected_ifc).split(" [")[0]

        self._loading = True
        try:
            selected_count = 0
            for row in self.data_grid.Rows:
                if row.Cells["Select"].Value == True:
                    self._set_combo_cell_value(row.Cells["NewIFC"], ifc_value)
                    elem_data = row.Tag
                    if elem_data is not None:
                        elem_data.new_ifc = ifc_value
                    selected_count += 1
        finally:
            self._loading = False
        
        if selected_count > 0:
            MessageBox.Show("Applied '{}' to {} selected elements.".format(ifc_value, selected_count),
                          "Bulk Assignment", MessageBoxButtons.OK, MessageBoxIcon.Information)
            self.UpdateStatistics()
        else:
            MessageBox.Show("No elements selected. Please check the 'Select' column for elements you want to modify.",
                          "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        
    def OnSelectAll(self, sender, args):
        """Select all visible rows"""
        for row in self.data_grid.Rows:
            row.Cells["Select"].Value = True
            elem_data = row.Tag
            elem_data.is_selected = True
        self.UpdateStatistics()
        
    def OnDeselectAll(self, sender, args):
        """Deselect all rows"""
        for row in self.data_grid.Rows:
            row.Cells["Select"].Value = False
            elem_data = row.Tag
            elem_data.is_selected = False
        self.UpdateStatistics()
        
    def OnCellValueChanged(self, sender, args):
        """Handle cell value change"""
        if self._loading or args.RowIndex < 0:
            return

        row = self.data_grid.Rows[args.RowIndex]
        elem_data = row.Tag
        if elem_data is None:
            return

        if args.ColumnIndex == self.data_grid.Columns["Select"].Index:
            elem_data.is_selected = row.Cells["Select"].Value
            self.UpdateStatistics()
        elif args.ColumnIndex == self.data_grid.Columns["NewIFC"].Index:
            new_value = row.Cells["NewIFC"].Value
            # Extract IFC class without description
            if new_value:
                ifc_value = str(new_value).split(" [")[0]
                elem_data.new_ifc = ifc_value
        
    def OnCurrentCellDirtyStateChanged(self, sender, args):
        """Commit checkbox changes immediately"""
        if self._loading:
            return
        if self.data_grid.IsCurrentCellDirty:
            self.data_grid.CommitEdit(System.Windows.Forms.DataGridViewDataErrorContexts.Commit)
            
    def OnSelectionChanged(self, sender, args):
        """Handle selection change"""
        if self._loading:
            return
        self.UpdateStatistics()

    def OnGridDataError(self, sender, args):
        """Swallow DataGridView data errors (e.g. combo value not in items)
        so the default error dialog never interrupts the user."""
        args.ThrowException = False
        
    def UpdateStatistics(self):
        """Update statistics labels"""
        total = len(self.filtered_data)
        selected = sum(1 for row in self.data_grid.Rows if row.Cells["Select"].Value == True)
        assigned = sum(1 for elem in self.filtered_data if elem.current_ifc)
        not_assigned = total - assigned
        
        self.stats_label.Text = "Total Elements: {}  |  Assigned: {}  |  Not Assigned: {}".format(
            total, assigned, not_assigned)
        
        self.selection_label.Text = "Selected for Modification: {} elements".format(selected)
        
    def OnApplyChanges(self, sender, args):
        """Apply changes to Revit"""
        # Count changes
        changes = []
        for row in self.data_grid.Rows:
            if row.Cells["Select"].Value == True:
                elem_data = row.Tag
                new_ifc = str(row.Cells["NewIFC"].Value)
                
                # Skip if keeping current or if no change
                if new_ifc == "(Keep Current)":
                    continue
                
                # Extract IFC value
                ifc_value = new_ifc.split(" [")[0] if " [" in new_ifc else new_ifc
                
                if new_ifc == "(Clear)":
                    ifc_value = ""
                
                # Only include if changed
                if ifc_value != elem_data.current_ifc:
                    changes.append((elem_data.element, ifc_value))
        
        if len(changes) == 0:
            MessageBox.Show("No changes to apply.", "Information", 
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
            return
        
        # Confirm
        result = MessageBox.Show(
            "Apply IFC assignments to {} elements?\n\nThis will modify the 'IfcExportAs' parameter.".format(len(changes)),
            "Confirm Changes",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)
        
        if result != DialogResult.Yes:
            return
        
        # Apply changes in transaction
        success_count = 0
        fail_count = 0
        error_messages = []
        
        with Transaction(doc, "DQT - Manual IFC Assignment") as t:
            t.Start()

            # Suppress warning dialogs during bulk parameter writes
            try:
                opts = t.GetFailureHandlingOptions()
                opts.SetFailuresPreprocessor(_SilentFailures())
                opts.SetClearAfterRollback(True)
                t.SetFailureHandlingOptions(opts)
            except:
                pass

            for element, ifc_value in changes:
                try:
                    # Try to get the parameter - Revit 2024+ first
                    param = element.LookupParameter("Export to IFC As")
                    
                    if param is None:
                        # Try Revit 2019-2023 name
                        param = element.LookupParameter("IfcExportAs")
                    
                    if param is None:
                        # Fallback to built-in parameter
                        param = element.get_Parameter(DB.BuiltInParameter.IFC_EXPORT_ELEMENT_AS)
                    
                    if param and not param.IsReadOnly:
                        if ifc_value:
                            param.Set(ifc_value)
                        else:
                            param.Set("")
                        success_count += 1
                    else:
                        fail_count += 1
                        error_messages.append("Element {} - Parameter not found or read-only".format(
                            _eid_int(element.Id)))
                        
                except Exception as e:
                    fail_count += 1
                    error_messages.append("Element {} - {}".format(
                        _eid_int(element.Id), str(e)))
            
            t.Commit()
        
        # Show results
        result_msg = "Assignment Complete!\n\n"
        result_msg += "Successfully assigned: {}\n".format(success_count)
        result_msg += "Failed: {}\n".format(fail_count)
        
        if fail_count > 0 and fail_count <= 10:
            result_msg += "\nErrors:\n" + "\n".join(error_messages[:10])
        elif fail_count > 10:
            result_msg += "\nShowing first 10 errors:\n" + "\n".join(error_messages[:10])
            result_msg += "\n... and {} more errors".format(fail_count - 10)
        
        MessageBox.Show(result_msg, "Results", MessageBoxButtons.OK, 
                       MessageBoxIcon.Information if fail_count == 0 else MessageBoxIcon.Warning)
        
        # Refresh current IFC values
        self.RefreshCurrentIFC()
        
    def RefreshCurrentIFC(self):
        """Refresh current IFC values from Revit"""
        for elem_data in self.elements_data:
            elem_data.current_ifc = GetCurrentIFCExport(elem_data.element)
            elem_data.new_ifc = elem_data.current_ifc
        
        self.LoadData()
        
    def OnExportReport(self, sender, args):
        """Export report to file"""
        try:
            from System.Windows.Forms import SaveFileDialog, DialogResult as DR
            
            save_dialog = SaveFileDialog()
            save_dialog.Filter = "CSV Files (*.csv)|*.csv|Text Files (*.txt)|*.txt"
            save_dialog.FileName = "IFC_Assignment_Report.csv"
            
            if save_dialog.ShowDialog() == DR.OK:
                filepath = save_dialog.FileName
                
                with open(filepath, 'w') as f:
                    # Write header
                    f.write("Element ID,Category,Element Name,Current IFC,New IFC\n")
                    
                    # Write data
                    for row in self.data_grid.Rows:
                        elem_id = row.Cells["ElementID"].Value
                        category = row.Cells["Category"].Value
                        name = row.Cells["ElementName"].Value
                        current_ifc = row.Cells["CurrentIFC"].Value
                        new_ifc = row.Cells["NewIFC"].Value
                        
                        # Clean values
                        name = str(name).replace(",", ";")
                        current_ifc = str(current_ifc).replace(",", ";")
                        new_ifc = str(new_ifc).replace(",", ";")
                        
                        f.write("{},{},{},{},{}\n".format(
                            elem_id, category, name, current_ifc, new_ifc))
                
                MessageBox.Show("Report exported successfully to:\n{}".format(filepath),
                              "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information)
        except Exception as e:
            MessageBox.Show("Error exporting report:\n{}".format(str(e)),
                          "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        
    def OnClose(self, sender, args):
        """Close the form"""
        self.Close()


def GetCurrentIFCExport(element):
    """Get current IFC Export As parameter value"""
    try:
        # Try Revit 2024+ parameter name first
        param = element.LookupParameter("Export to IFC As")
        
        if param is None:
            # Try Revit 2019-2023 parameter name
            param = element.LookupParameter("IfcExportAs")
        
        if param is None:
            # Fallback to built-in parameter (always works)
            param = element.get_Parameter(DB.BuiltInParameter.IFC_EXPORT_ELEMENT_AS)
        
        if param and param.HasValue:
            return param.AsString()
    except:
        pass
    
    return None


def GetElementName(element):
    """Get element name or type name"""
    try:
        # Try to get element name
        name_param = element.get_Parameter(DB.BuiltInParameter.ELEM_NAME_PARAM)
        if name_param and name_param.HasValue:
            name = name_param.AsString()
            if name:
                return name
        
        # Try to get type name
        elem_type_id = element.GetTypeId()
        if elem_type_id != DB.ElementId.InvalidElementId:
            elem_type = doc.GetElement(elem_type_id)
            if elem_type:
                type_name = elem_type.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
                if type_name and type_name.HasValue:
                    return type_name.AsString()
        
        # Fallback to family and type
        if hasattr(element, 'Name'):
            return element.Name
            
    except:
        pass
    
    return "Unnamed Element"


def CollectElements():
    """Collect all relevant elements from the document"""
    output = script.get_output()
    output.print_md("## Collecting Elements from Active View...")
    
    # Categories to include
    categories_to_include = [
        DB.BuiltInCategory.OST_Walls,
        DB.BuiltInCategory.OST_Floors,
        DB.BuiltInCategory.OST_Roofs,
        DB.BuiltInCategory.OST_Ceilings,
        DB.BuiltInCategory.OST_Doors,
        DB.BuiltInCategory.OST_Windows,
        DB.BuiltInCategory.OST_Columns,
        DB.BuiltInCategory.OST_StructuralColumns,
        DB.BuiltInCategory.OST_StructuralFraming,
        DB.BuiltInCategory.OST_StructuralFoundation,
        DB.BuiltInCategory.OST_Stairs,
        DB.BuiltInCategory.OST_Ramps,
        DB.BuiltInCategory.OST_Railings,
        DB.BuiltInCategory.OST_Rooms,
        DB.BuiltInCategory.OST_Areas,
        DB.BuiltInCategory.OST_GenericModel,
        DB.BuiltInCategory.OST_Furniture,
        DB.BuiltInCategory.OST_LightingFixtures,
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_MechanicalEquipment,
        DB.BuiltInCategory.OST_SpecialityEquipment,
        DB.BuiltInCategory.OST_PipeCurves,
        DB.BuiltInCategory.OST_PipeFitting,
        DB.BuiltInCategory.OST_PipeAccessory,
        DB.BuiltInCategory.OST_DuctCurves,
        DB.BuiltInCategory.OST_DuctFitting,
        DB.BuiltInCategory.OST_DuctAccessory,
        DB.BuiltInCategory.OST_Topography,
        DB.BuiltInCategory.OST_Planting,
        DB.BuiltInCategory.OST_CurtainWallPanels,
    ]
    
    # Create multi-category filter
    cat_filters = [DB.ElementCategoryFilter(cat) for cat in categories_to_include]
    multi_filter = DB.LogicalOrFilter(List[DB.ElementFilter](cat_filters))
    
    # Collect elements
    collector = DB.FilteredElementCollector(doc, doc.ActiveView.Id)
    elements = collector.WherePasses(multi_filter).WhereElementIsNotElementType().ToElements()
    
    output.print_md("Found **{}** elements in active view".format(len(elements)))
    
    # Create ElementData objects
    elements_data = []
    
    for elem in elements:
        try:
            category = elem.Category.Name if elem.Category else "Unknown"
            name = GetElementName(elem)
            current_ifc = GetCurrentIFCExport(elem)
            
            elem_data = ElementData(elem, category, name, current_ifc)
            elements_data.append(elem_data)
        except:
            continue
    
    output.print_md("Prepared **{}** elements for assignment".format(len(elements_data)))
    
    return elements_data


# Main execution
if __name__ == '__main__':
    output = script.get_output()
    
    output.print_md("# IFC-SG Manual Assignment Tool")
    output.print_md("---")
    
    # Collect elements
    elements_data = CollectElements()
    
    if len(elements_data) == 0:
        forms.alert("No elements found in active view. Please open a view with elements and try again.",
                   exitscript=True)
    
    # Show statistics
    output.print_md("## Statistics")
    categories = {}
    assigned = 0
    not_assigned = 0
    
    for elem_data in elements_data:
        # Count by category
        if elem_data.category not in categories:
            categories[elem_data.category] = 0
        categories[elem_data.category] += 1
        
        # Count assigned/not assigned
        if elem_data.current_ifc:
            assigned += 1
        else:
            not_assigned += 1
    
    output.print_md("- **Total Elements:** {}".format(len(elements_data)))
    output.print_md("- **Currently Assigned:** {}".format(assigned))
    output.print_md("- **Not Assigned:** {}".format(not_assigned))
    output.print_md("")
    output.print_md("### By Category:")
    for cat in sorted(categories.keys()):
        output.print_md("- {}: **{}**".format(cat, categories[cat]))
    
    output.print_md("---")
    output.print_md("Opening assignment interface...")
    
    # Show form
    form = IFCAssignmentForm(elements_data)
    form.ShowDialog()
    
    output.print_md("## Complete!")
    output.print_md("Tool execution finished.")