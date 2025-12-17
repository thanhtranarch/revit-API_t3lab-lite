# -*- coding: utf-8 -*-
"""
Location Points Exporter with Base Points and List Selection

Exports Revit element location points and base points to CSV format, 
including Project Base Point and Survey Point coordinates.
Features list-based selection with Category and Family Type display.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
LinkedIn: linkedin.com/in/sunarch7899/
"""

# IMPORT LIBRARIES
# ==================================================
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
import csv
import os
import sys
from pyrevit import forms


# DEFINE VARIABLES
# ==================================================
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# CLASS/FUNCTIONS
# ==================================================
def get_save_file_path():
    try:
        import clr
        clr.AddReference('System.Windows.Forms')
        clr.AddReference('System')
        
        from System.Windows.Forms import SaveFileDialog, DialogResult
        from datetime import datetime
        
        date_string = datetime.now().strftime("%y%m%d")
        doc_title = doc.Title
        if doc_title.endswith('.rvt'):
            doc_name = doc_title[:-4]
        else:
            doc_name = doc_title
        parts = doc_name.split("_")  # Tách bằng dấu gạch dưới
        prefix = parts[0]  # FIX
        level = "L" + parts[-2]
        level = parts[-2] 
        default_filename = "{}_{}_Drilling point_{}".format(prefix, level, date_string)

        save_dialog = SaveFileDialog()
        save_dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        save_dialog.DefaultExt = "csv"
        save_dialog.AddExtension = True
        save_dialog.FileName = default_filename
        save_dialog.InitialDirectory = os.path.expanduser("~\\Documents")
        save_dialog.Title = "Choose location to save CSV file"
        
        result = save_dialog.ShowDialog()
        
        if result.ToString() == "OK":
            return save_dialog.FileName
        else:
            return None
            
    except Exception as e:
        print("\n[WARNING] File dialog failed: {}".format(str(e)))
        print("\nPlease enter the full path where you want to save the CSV file:")
        print("Example: C:\\Users\\YourName\\Documents\\MyModel_HangerBracket_Location.csv")
        
        file_path = raw_input("\nEnter file path: ").strip()
        
        if file_path:
            if not file_path.lower().endswith('.csv'):
                file_path = file_path + '.csv'
            return file_path
        else:
            return None

class BasePointsRetriever:
    
    def __init__(self, doc):
        self.doc = doc
    
    def get_base_point_element_info(self, base_point_element):
        try:
            if base_point_element is None:
                return None
            
            element_id = base_point_element.Id.IntegerValue
            element_name = base_point_element.Name if hasattr(base_point_element, 'Name') else ""
            
            location = base_point_element.Position
            if location is None:
                return None
            
            point = location
            
            x_feet = point.X
            y_feet = point.Y
            z_feet = point.Z
            
            # Convert coordinates from internal units (feet) to millimeters using Revit API
            x_mm = UnitUtils.ConvertFromInternalUnits(x_feet, UnitTypeId.Millimeters)
            y_mm = UnitUtils.ConvertFromInternalUnits(y_feet, UnitTypeId.Millimeters)
            z_mm = UnitUtils.ConvertFromInternalUnits(z_feet, UnitTypeId.Millimeters)
            
            category_name = "Base Point"
            if base_point_element.Category is not None:
                category_name = base_point_element.Category.Name
            
            element_type_name = ""
            try:
                element_type = self.doc.GetElement(base_point_element.GetTypeId())
                if element_type is not None:
                    element_type_name = element_type.Name
            except:
                pass
            
            is_shared = base_point_element.IsShared
            
            info = {
                'id': element_id,
                'name': element_name,
                'x_mm': x_mm,
                'y_mm': y_mm,
                'z_mm': z_mm,
                'category': category_name,
                'type': element_type_name if element_type_name else ("Survey Point" if is_shared else "Project Base Point"),
                'location_point': point,
                'is_shared': is_shared
            }
            
            return info
            
        except Exception as e:
            print("Error extracting base point info: {}".format(str(e)))
            return None
    
    def get_base_points_info(self):
        result = {}
        
        try:
            project_bp = BasePoint.GetProjectBasePoint(self.doc)
            if project_bp is not None:
                project_info = self.get_base_point_element_info(project_bp)
                if project_info is not None:
                    result['project_base_point'] = project_info
        except Exception as e:
            print("Error retrieving Project Base Point: {}".format(str(e)))
        
        try:
            survey_bp = BasePoint.GetSurveyPoint(self.doc)
            if survey_bp is not None:
                survey_info = self.get_base_point_element_info(survey_bp)
                if survey_info is not None:
                    result['survey_point'] = survey_info
        except Exception as e:
            print("Error retrieving Survey Point: {}".format(str(e)))
        
        return result


class LocationPointsExporter:
    
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.elements_data = []
    
    def get_element_location(self, element):
        try:
            location = element.Location
            
            if location is None:
                return None
            
            if isinstance(location, LocationPoint):
                point = location.Point
                return point
            
            return None
            
        except:
            return None
    
    def get_element_info(self, element):
        try:
            element_id = element.Id.IntegerValue
            
            location = self.get_element_location(element)
            if location is None:
                return None
            
            x_feet = location.X
            y_feet = location.Y
            z_feet = location.Z
            
            # Convert coordinates from internal units (feet) to millimeters using Revit API
            x_mm = UnitUtils.ConvertFromInternalUnits(x_feet, UnitTypeId.Millimeters)
            y_mm = UnitUtils.ConvertFromInternalUnits(y_feet, UnitTypeId.Millimeters)
            z_mm = UnitUtils.ConvertFromInternalUnits(z_feet, UnitTypeId.Millimeters)
            
            category_name = ""
            if element.Category is not None:
                category_name = element.Category.Name
            
            element_type_name = ""
            try:
                element_type = self.doc.GetElement(element.GetTypeId())
                if element_type is not None:
                    element_type_name = element_type.Name
            except:
                pass
            
            element_name = element.Name if hasattr(element, 'Name') else ""
            
            # Get Family name from the element type
            family_name = ""
            try:
                element_type = self.doc.GetElement(element.GetTypeId())
                if element_type is not None:
                    # Try to get Family parameter from the type
                    family_param = element_type.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
                    if family_param is not None and family_param.AsString():
                        family_name = family_param.AsString()
                    # Alternative: get from type's symbol family name
                    if not family_name and hasattr(element_type, 'FamilyName'):
                        family_name = element_type.FamilyName
            except:
                pass
            
            info = {
                'id': element_id,
                'x_mm': x_mm,
                'y_mm': y_mm,
                'z_mm': z_mm,
                'category': category_name,
                'type': element_type_name,
                'family': family_name,
                'name': element_name
            }
            
            return info
            
        except Exception as e:
            print("Error processing element: {}".format(str(e)))
            return None
    
    def get_all_elements_in_view(self):
        """Get all elements in active view that have location points"""
        try:
            active_view = self.uidoc.ActiveView
            if active_view is None:
                print("No active view found")
                return []
            
            # Get all elements in the view using FilteredElementCollector
            collector = FilteredElementCollector(self.doc, active_view.Id)
            all_elements = list(collector.WhereElementIsNotElementType())
            
            # Filter elements that have location point
            valid_elements = []
            for element in all_elements:
                try:
                    # Check if element is in Specialty Equipment category
                    if element.Category is not None:
                        category_name = element.Category.Name
                        if category_name != "Specialty Equipment":
                            continue  # Skip elements that are not Specialty Equipment
                    
                    location = self.get_element_location(element)
                    if location is not None:
                        valid_elements.append(element)
                except:
                    pass
            
            return valid_elements
        except Exception as e:
            print("Error getting elements from view: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []
    
    def display_and_select_elements(self, elements):
        """Display elements using pyRevit SelectFromList UI and get user selection"""
        if not elements:
            print("No elements with location points found in current view")
            return []
        
        try:
            print("Showing selection dialog...")
            
            # Create custom wrapper objects with Category and Name display
            class ElementDisplay(object):
                def __init__(self, element):
                    self.element = element
                    
                def __str__(self):
                    try:
                        category_name = ""
                        if self.element.Category is not None:
                            category_name = self.element.Category.Name
                        element_name = self.element.Name if hasattr(self.element, 'Name') else ""
                        
                        # Create display format: "Category - Name"
                        return "{} - {}".format(category_name, element_name) if category_name else element_name
                    except:
                        return str(self.element)
            
            # Wrap elements with display class
            display_items = [ElementDisplay(element) for element in elements]
            
            # Show SelectFromList with custom display
            selected_items = forms.SelectFromList.show(
                display_items,
                multiselect=True,
                title="Select Elements to Export",
                button_name="Export"
            )
            
            if selected_items:
                # Extract actual elements from the wrapper objects
                selected_elements = [item.element for item in selected_items]
                print("Selected {} element(s)".format(len(selected_elements)))
                return selected_elements
            else:
                print("No elements selected")
                return []
                
        except Exception as e:
            print("Selection error: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []
    
    def export_to_csv(self, file_path, base_points_info):
        try:
            # Define CSV headers (columns) exactly as required
            fieldnames = ['ID', 'X (mm)', 'Y (mm)', 'Z (mm)', 'Category', 'Family and Type']
            
            with open(file_path, 'wb') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                # # Export Project Base Point if available
                # if base_points_info:
                #     if 'project_base_point' in base_points_info:
                #         pbp = base_points_info['project_base_point']
                #         row = {
                #             'ID': pbp['id'],
                #             'X (mm)': "{:.2f}".format(pbp['x_mm']),
                #             'Y (mm)': "{:.2f}".format(pbp['y_mm']),
                #             'Z (mm)': "{:.2f}".format(pbp['z_mm']),
                #             'Category': 'Base Point',
                #             'Family and Type': 'Project Base Point'
                #         }
                #         writer.writerow(row)
                    
                #     # Export Survey Point if available
                #     if 'survey_point' in base_points_info:
                #         sp = base_points_info['survey_point']
                #         row = {
                #             'ID': sp['id'],
                #             'X (mm)': "{:.2f}".format(sp['x_mm']),
                #             'Y (mm)': "{:.2f}".format(sp['y_mm']),
                #             'Z (mm)': "{:.2f}".format(sp['z_mm']),
                #             'Category': 'Base Point',
                #             'Family and Type': 'Survey Point'
                #         }
                #         writer.writerow(row)
                
                # Export selected elements location points
                for data in self.elements_data:
                    # Combine Family and Type for display
                    family_and_type = data['family']
                    if data['type']:
                        family_and_type = "{} - {}".format(data['family'], data['type']) if data['family'] else data['type']
                    
                    row = {
                        'ID': data['id'],
                        'X (mm)': "{:.2f}".format(data['x_mm']),
                        'Y (mm)': "{:.2f}".format(data['y_mm']),
                        'Z (mm)': "{:.2f}".format(data['z_mm']),
                        'Category': data['category'],
                        'Family and Type': family_and_type
                    }
                    writer.writerow(row)
            
            print("Export successful!")
            print("File saved to: {}".format(file_path))
            # print("Total elements exported: {}".format(len(self.elements_data) + len([x for x in base_points_info if base_points_info.get(x)])))
            
            return True
            
        except Exception as e:
            print("Error exporting to CSV: {}".format(str(e)))
            return False
    
    def print_summary(self):
        print("\n" + "=" * 140)
        print("EXPORT SUMMARY")
        print("=" * 140)
        
        if not self.elements_data:
            print("No data to display")
            return
        
        print("\n{:<10} {:<18} {:<18} {:<18} {:<25} {:<45}".format(
            "ID", "X (mm)", "Y (mm)", "Z (mm)", "Category", "Family and Type"))
        print("-" * 140)
        
        for data in self.elements_data:
            # Combine Family and Type for display
            family_and_type = data['family']
            if data['type']:
                family_and_type = "{} - {}".format(data['family'], data['type']) if data['family'] else data['type']
            
            print("{:<10} {:<18.2f} {:<18.2f} {:<18.2f} {:<25} {:<45}".format(
                data['id'],
                data['x_mm'],
                data['y_mm'],
                data['z_mm'],
                data['category'][:24],
                family_and_type[:44]
            ))
        
        print("\n" + "=" * 140)

# MAIN SCRIPT
# ==================================================
try:
    print("\n" + "#" * 80)
    print("# LOCATION POINTS EXPORTER WITH BASE POINTS")
    print("#" * 80)
    print("\nDocument: {}".format(doc.Title))
    
    print("\n[STEP 1] Getting elements from current view...")
    exporter = LocationPointsExporter(doc, uidoc)
    all_elements = exporter.get_all_elements_in_view()
    
    if not all_elements:
        print("No elements with location points found in current view")
        sys.exit()
    
    # Show pyRevit selection dialog
    selected_elements = exporter.display_and_select_elements(all_elements)
    
    if not selected_elements:
        print("No elements selected")
        sys.exit()
    
    # Process selected elements
    exporter.elements_data = []
    
    for element in selected_elements:
        info = exporter.get_element_info(element)
        if info is not None:
            exporter.elements_data.append(info)
    
    if not exporter.elements_data:
        print("No valid location points found in selected elements")
        sys.exit()
    
    print("\n[STEP 2] Retrieving Base Points...")
    base_points_retriever = BasePointsRetriever(doc)
    base_points_info = base_points_retriever.get_base_points_info()
    
    if base_points_info:
        print("[SUCCESS] Base Points retrieved")
        if 'project_base_point' in base_points_info:
            pbp = base_points_info['project_base_point']
            print("  Project Base Point: X={:.2f}, Y={:.2f}, Z={:.2f}".format(pbp['x_mm'], pbp['y_mm'], pbp['z_mm']))
        if 'survey_point' in base_points_info:
            sp = base_points_info['survey_point']
            print("  Survey Point: X={:.2f}, Y={:.2f}, Z={:.2f}".format(sp['x_mm'], sp['y_mm'], sp['z_mm']))
    else:
        print("[WARNING] No base points found")
    
    print("\n[STEP 3] Opening file save dialog...")
    output_file = get_save_file_path()
    
    if output_file is None:
        print("[CANCELLED] Export cancelled by user")
        sys.exit()
    
    print("[SUCCESS] Save location selected: {}".format(output_file))
    
    exporter.print_summary()
    
    print("\n[STEP 4] Exporting to CSV...")
    
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    exporter.export_to_csv(output_file, base_points_info)
    
except Exception as e:
    print("Error: {}".format(str(e)))
    import traceback
    traceback.print_exc()