# -*- coding: utf-8 -*-
"""
Drilling Point Coordination - Parameter Selection Exporter

Exports Revit element drilling point coordinates with selected parameters to CSV format,
including Project Base Point and Survey Point coordinates.
Features parameter selection UI and handles missing parameters with N/A values.

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
        parts = doc_name.split("_")
        prefix = parts[0]
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


class ParameterExtractor:
    """Extract and manage parameters from Revit elements"""
    
    def __init__(self, doc):
        self.doc = doc
        self.parameters_map = {}  # {parameter_name: parameter_definition}
    
    def get_element_parameters(self, element):
        """Get all parameters from a single element with their values"""
        try:
            params_dict = {}
            
            for param in element.Parameters:
                try:
                    param_name = param.Definition.Name
                    param_value = self._get_parameter_value(param)
                    
                    if param_value is not None:
                        params_dict[param_name] = param_value
                except:
                    pass
            
            return params_dict
        except Exception as e:
            print("Error extracting parameters from element: {}".format(str(e)))
            return {}
    
    def _get_parameter_value(self, param):
        """Extract parameter value with proper formatting"""
        try:
            if param.StorageType == StorageType.String:
                return param.AsString()
            elif param.StorageType == StorageType.Integer:
                return str(param.AsInteger())
            elif param.StorageType == StorageType.Double:
                # Convert from internal units to mm
                value_internal = param.AsDouble()
                value_mm = UnitUtils.ConvertFromInternalUnits(value_internal, UnitTypeId.Millimeters)
                return "{:.2f}".format(value_mm)
            elif param.StorageType == StorageType.ElementId:
                elem_id = param.AsElementId()
                if elem_id.IntegerValue == -1:
                    return None
                try:
                    ref_element = self.doc.GetElement(elem_id)
                    return ref_element.Name if ref_element else str(elem_id.IntegerValue)
                except:
                    return str(elem_id.IntegerValue)
            else:
                return param.AsValueString()
        except:
            return None
    
    def extract_unique_parameters(self, elements):
        """Extract unique parameters from all elements (deduplicated)"""
        unique_params = {}
        
        for element in elements:
            params = self.get_element_parameters(element)
            for param_name, param_value in params.items():
                if param_name not in unique_params:
                    unique_params[param_name] = param_name
        
        return sorted(unique_params.keys())
    
    def get_parameter_value_from_element(self, element, parameter_name):
        """Get specific parameter value from element, return N/A if not found"""
        try:
            for param in element.Parameters:
                if param.Definition.Name == parameter_name:
                    value = self._get_parameter_value(param)
                    return value if value is not None else "N/A"
            return "N/A"
        except:
            return "N/A"


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


class DrillingPointCoordination:
    
    def __init__(self, doc, uidoc):
        self.doc = doc
        self.uidoc = uidoc
        self.elements_data = []
        self.parameter_extractor = ParameterExtractor(doc)
    
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
    
    def get_element_info(self, element, selected_parameters=None):
        """Get element info with selected parameters"""
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
                    family_param = element_type.get_Parameter(BuiltInParameter.ELEM_FAMILY_PARAM)
                    if family_param is not None and family_param.AsString():
                        family_name = family_param.AsString()
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
            
            # Add selected parameters (nếu có)
            if selected_parameters:
                for param_name in selected_parameters:
                    param_value = self.parameter_extractor.get_parameter_value_from_element(element, param_name)
                    info[param_name] = param_value
            
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
            
            collector = FilteredElementCollector(self.doc, active_view.Id)
            all_elements = list(collector.WhereElementIsNotElementType())
            
            valid_elements = []
            for element in all_elements:
                try:
                    if element.Category is not None:
                        category_name = element.Category.Name
                        if category_name != "Specialty Equipment":
                            continue
                    
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
    
    def display_parameter_selection(self, parameters):
        """Show UI to select parameters for export - OPTIONAL (không bắt buộc)"""
        if not parameters:
            print("No parameters found")
            return []
        
        try:
            print("Showing parameter selection dialog...")
            print("(You can skip this step by clicking Cancel)")
            
            selected_params = forms.SelectFromList.show(
                parameters,
                multiselect=True,
                title="Add-on Parameters to Export (Optional)",
                button_name="Next"
            )
            
            if selected_params:
                print("Selected {} parameter(s): {}".format(len(selected_params), ", ".join(selected_params)))
                return selected_params
            else:
                print("No parameters selected - will export with default columns only")
                return []
                
        except Exception as e:
            print("Parameter selection error: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []
    
    def display_and_select_elements(self, elements):
        """Display elements using pyRevit SelectFromList UI and get user selection"""
        if not elements:
            print("No elements with location points found in current view")
            return []
        
        try:
            print("Showing element selection dialog...")
            
            class ElementDisplay(object):
                def __init__(self, element):
                    self.element = element
                    
                def __str__(self):
                    try:
                        category_name = ""
                        if self.element.Category is not None:
                            category_name = self.element.Category.Name
                        element_name = self.element.Name if hasattr(self.element, 'Name') else ""
                        
                        return "{} - {}".format(category_name, element_name) if category_name else element_name
                    except:
                        return str(self.element)
            
            display_items = [ElementDisplay(element) for element in elements]
            
            selected_items = forms.SelectFromList.show(
                display_items,
                multiselect=True,
                title="Select Elements to Export",
                button_name="Export"
            )
            
            if selected_items:
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
    
    def export_to_csv(self, file_path, selected_parameters, base_points_info):
        """Export to CSV with selected parameters (hoặc không có parameters)"""
        try:
            # Define base CSV headers
            fieldnames = ['ID', 'X (mm)', 'Y (mm)', 'Z (mm)', 'Category', 'Family and Type']
            
            # Add selected parameters to fieldnames (nếu có)
            if selected_parameters:
                fieldnames.extend(selected_parameters)
            
            with open(file_path, 'wb') as csvfile:
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                # Export selected elements location points
                for data in self.elements_data:
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
                    
                    # Add selected parameter values (nếu có)
                    if selected_parameters:
                        for param_name in selected_parameters:
                            row[param_name] = data.get(param_name, "N/A")
                    
                    writer.writerow(row)
            
            print("Export successful!")
            print("File saved to: {}".format(file_path))
            print("Total elements exported: {}".format(len(self.elements_data)))
            
            return True
            
        except Exception as e:
            print("Error exporting to CSV: {}".format(str(e)))
            return False
    
    def print_summary(self, selected_parameters):
        """Print export summary"""
        print("\n" + "=" * 160)
        print("EXPORT SUMMARY")
        print("=" * 160)
        
        if not self.elements_data:
            print("No data to display")
            return
        
        # Calculate column widths
        col_widths = {
            'ID': 10,
            'X (mm)': 15,
            'Y (mm)': 15,
            'Z (mm)': 15,
            'Category': 20,
            'Family and Type': 35
        }
        
        # Print header
        header = "{:<10} {:<15} {:<15} {:<15} {:<20} {:<35}".format(
            "ID", "X (mm)", "Y (mm)", "Z (mm)", "Category", "Family and Type")
        
        if selected_parameters:
            for param in selected_parameters[:3]:  # Show first 3 params in summary
                header += " {:<20}".format(param[:19])
        
        print(header)
        print("-" * 160)
        
        # Print data rows
        for data in self.elements_data:
            family_and_type = data['family']
            if data['type']:
                family_and_type = "{} - {}".format(data['family'], data['type']) if data['family'] else data['type']
            
            row = "{:<10} {:<15.2f} {:<15.2f} {:<15.2f} {:<20} {:<35}".format(
                data['id'],
                data['x_mm'],
                data['y_mm'],
                data['z_mm'],
                data['category'][:19],
                family_and_type[:34]
            )
            
            if selected_parameters:
                for param in selected_parameters[:3]:
                    param_value = data.get(param, "N/A")
                    row += " {:<20}".format(str(param_value)[:19])
            
            print(row)
        
        print("\n" + "=" * 160)


# MAIN SCRIPT
# ==================================================
try:
    print("\n" + "#" * 80)
    print("# DRILLING POINT COORDINATION - PARAMETER SELECTION EXPORTER")
    print("#" * 80)
    print("\nDocument: {}".format(doc.Title))
    
    print("\n[STEP 1] Getting drilling points from current view...")
    exporter = DrillingPointCoordination(doc, uidoc)
    all_elements = exporter.get_all_elements_in_view()
    
    if not all_elements:
        print("No drilling points with location data found in current view")
        sys.exit()
    
    print("Found {} drilling point(s)".format(len(all_elements)))
    
    print("\n[STEP 2] Extracting unique parameters from drilling points...")
    unique_parameters = exporter.parameter_extractor.extract_unique_parameters(all_elements)
    
    selected_parameters = []  # Initialize as empty list
    
    if unique_parameters:
        print("Found {} unique parameter(s)".format(len(unique_parameters)))
        print("Available parameters:")
        for i, param in enumerate(unique_parameters[:10], 1):
            print("  {}. {}".format(i, param))
        if len(unique_parameters) > 10:
            print("  ... and {} more".format(len(unique_parameters) - 10))
        
        # Show parameter selection dialog (OPTIONAL - không bắt buộc)
        selected_parameters = exporter.display_parameter_selection(unique_parameters)
    else:
        print("No parameters found in drilling points - will export default columns only")
    
    print("\n[STEP 3] Selecting drilling points to export...")
    # Show element selection dialog
    selected_elements = exporter.display_and_select_elements(all_elements)
    
    if not selected_elements:
        print("No drilling points selected")
        sys.exit()
    
    # Process selected elements with selected parameters
    exporter.elements_data = []
    
    for element in selected_elements:
        info = exporter.get_element_info(element, selected_parameters if selected_parameters else None)
        if info is not None:
            exporter.elements_data.append(info)
    
    if not exporter.elements_data:
        print("No valid drilling point coordinates found in selected elements")
        sys.exit()
    
    print("\n[STEP 4] Opening file save dialog...")
    output_file = get_save_file_path()
    
    if output_file is None:
        print("[CANCELLED] Export cancelled by user")
        sys.exit()
    
    print("[SUCCESS] Save location selected: {}".format(output_file))
    
    exporter.print_summary(selected_parameters)
    
    print("\n[STEP 5] Exporting to CSV...")
    
    output_dir = os.path.dirname(output_file)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Retrieve base points (optional - currently not exported)
    base_points_retriever = BasePointsRetriever(doc)
    base_points_info = base_points_retriever.get_base_points_info()
    
    exporter.export_to_csv(output_file, selected_parameters, base_points_info)
    
except Exception as e:
    print("Error: {}".format(str(e)))
    import traceback
    traceback.print_exc()