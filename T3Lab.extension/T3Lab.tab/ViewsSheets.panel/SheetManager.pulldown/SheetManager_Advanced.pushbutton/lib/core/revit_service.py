# -*- coding: utf-8 -*-
"""
Sheet Manager - Revit Service
CLEANED - Sheet Methods Only

Copyright © Dang Quoc Truong (DQT)
"""

from Autodesk.Revit.DB import FilteredElementCollector, ViewSheet, BuiltInParameter


class RevitService(object):
    """Revit service for sheet operations"""
    
    def __init__(self, doc):
        self.doc = doc
    
    def get_all_sheets(self):
        """Get all sheets in the document"""
        from core.data_models import SheetModel
        
        collector = FilteredElementCollector(self.doc).OfClass(ViewSheet)
        sheets = []
        
        for sheet in collector:
            if not sheet.IsPlaceholder:
                sheets.append(SheetModel(sheet))
        
        return sheets
    
    def update_sheet(self, sheet_model):
        """Update sheet parameters"""
        try:
            sheet = sheet_model.element
            
            # Update sheet number - use direct property
            if sheet_model.sheet_number != sheet_model._original_sheet_number:
                sheet.SheetNumber = sheet_model.sheet_number
            
            # Update sheet name - use direct property
            if sheet_model.sheet_name != sheet_model._original_sheet_name:
                sheet.Name = sheet_model.sheet_name
            
            return True
        except Exception as e:
            print("Error updating sheet: {}".format(str(e)))
            return False
    
    def create_sheet(self, sheet_number, sheet_name, titleblock_id):
        """Create a new sheet"""
        try:
            new_sheet = ViewSheet.Create(self.doc, titleblock_id)
            
            # Set number - use direct property
            new_sheet.SheetNumber = sheet_number
            
            # Set name - use direct property
            new_sheet.Name = sheet_name
            
            return new_sheet
        except Exception as e:
            print("Error creating sheet: {}".format(str(e)))
            return None
    
    def delete_sheet(self, sheet_id):
        """Delete a sheet"""
        try:
            self.doc.Delete(sheet_id)
            return True
        except Exception as e:
            print("Error deleting sheet: {}".format(str(e)))
            return False
    
    def duplicate_sheet(self, source_sheet):
        """Duplicate a sheet with unique number"""
        try:
            from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ViewSheet
            import random
            
            # Get titleblock from source sheet
            titleblock_id = None
            collector = FilteredElementCollector(self.doc, source_sheet.Id).OfCategory(
                BuiltInCategory.OST_TitleBlocks)
            
            for elem in collector:
                titleblock_id = elem.GetTypeId()
                break
            
            if not titleblock_id:
                print("DEBUG: No titleblock found on source sheet")
                return None
            
            # Generate unique sheet number
            base_number = source_sheet.SheetNumber
            new_name = source_sheet.Name
            
            # Try to find unique number
            attempt = 0
            while attempt < 100:
                if attempt == 0:
                    new_number = base_number + " - Copy"
                else:
                    # Add random suffix to avoid conflicts
                    suffix = random.randint(100, 999)
                    new_number = base_number + " - Copy{}".format(suffix)
                
                # Check if number exists
                existing = FilteredElementCollector(self.doc).OfClass(ViewSheet).WhereElementIsNotElementType()
                exists = False
                for sheet in existing:
                    if sheet.SheetNumber == new_number:
                        exists = True
                        break
                
                if not exists:
                    break
                
                attempt += 1
            
            if attempt >= 100:
                print("ERROR: Could not generate unique sheet number after 100 attempts")
                return None
            
            new_sheet = self.create_sheet(new_number, new_name, titleblock_id)
            return new_sheet
            
        except Exception as e:
            print("Error duplicating sheet: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return None
            return None
    
    def get_all_titleblocks(self):
        """Get ALL titleblock types in project - SIMPLE"""
        from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory
        
        try:
            print("DEBUG: Searching for titleblocks in project...")
            
            # Method 1: Get ALL titleblock family symbols (types)
            collector = FilteredElementCollector(self.doc).OfCategory(
                BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType()
            
            titleblocks = []
            for tb_type in collector:
                try:
                    # Get name
                    name = tb_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                    if not name:
                        name = "Titleblock"
                    
                    titleblocks.append((tb_type.Id, name))
                    print("DEBUG: Found titleblock type: {}".format(name))
                except Exception as e:
                    print("DEBUG: Error reading titleblock: {}".format(str(e)))
                    continue
            
            print("DEBUG: Total titleblocks found: {}".format(len(titleblocks)))
            
            # If still nothing, try getting instances and their types
            if not titleblocks:
                print("DEBUG: No types found, trying instances...")
                instance_collector = FilteredElementCollector(self.doc).OfCategory(
                    BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType()
                
                type_ids = set()
                for instance in instance_collector:
                    type_id = instance.GetTypeId()
                    if type_id not in type_ids:
                        type_ids.add(type_id)
                        tb_type = self.doc.GetElement(type_id)
                        if tb_type:
                            try:
                                name = tb_type.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM).AsString()
                                if not name:
                                    name = "Titleblock"
                                titleblocks.append((type_id, name))
                                print("DEBUG: Found from instance: {}".format(name))
                            except:
                                titleblocks.append((type_id, "Titleblock"))
            
            return titleblocks
            
        except Exception as e:
            print("ERROR getting titleblocks: {}".format(str(e)))
            import traceback
            traceback.print_exc()
            return []