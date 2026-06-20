# -*- coding: utf-8 -*-
"""
Families Scanners (Phase 5)
Scanners for families: detail components, unused families, unused family types, annotation families, profile families

Copyright (c) 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Family,
    FamilySymbol,
    FamilyInstance,
    BuiltInCategory,
    ElementId
)
from Autodesk.Revit import DB

try:
    from purge_scanner import BasePurgeScanner
except:
    # Fallback for testing
    class BasePurgeScanner:
        def __init__(self, doc):
            self.doc = doc


class DetailComponentsScanner(BasePurgeScanner):
    """Scanner for unused detail component families"""
    
    def scan(self):
        """Scan for detail component families with no instances"""
        unused_items = []
        
        try:
            print("DEBUG: Starting DetailComponentsScanner...")
            
            # Get all detail component families
            collector = FilteredElementCollector(self.doc)
            families = collector.OfClass(Family).ToElements()
            
            # Filter to detail components only
            detail_families = []
            for family in families:
                try:
                    if not family or not family.IsValidObject:
                        continue
                    
                    # Check if family category is detail items
                    if family.FamilyCategory:
                        cat_id = family.FamilyCategory.Id.IntegerValue
                        if cat_id == int(BuiltInCategory.OST_DetailComponents):
                            detail_families.append(family)
                except:
                    continue
            
            print("DEBUG: Found {} detail families".format(len(detail_families)))
            
            if len(detail_families) == 0:
                print("DEBUG: No detail component families in project")
                return unused_items
            
            # OPTIMIZATION: Build family usage dictionary once
            print("DEBUG: Building usage dictionary...")
            family_usage = {}
            instances = FilteredElementCollector(self.doc)\
                .OfClass(FamilyInstance)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            for inst in instances:
                try:
                    if inst and inst.IsValidObject and inst.Symbol:
                        symbol = inst.Symbol
                        if symbol and symbol.Family:
                            family_id = symbol.Family.Id.IntegerValue
                            family_usage[family_id] = True
                except:
                    continue
            
            print("DEBUG: {} families are used".format(len(family_usage)))
            
            # Check each detail family
            for family in detail_families:
                try:
                    # Get all types (symbols) in this family
                    symbol_ids = family.GetFamilySymbolIds()
                    
                    # Check if symbol_ids is valid
                    if not symbol_ids or len(list(symbol_ids)) == 0:
                        continue
                    
                    # Check if family is used (quick lookup)
                    family_id = family.Id.IntegerValue
                    if family_id not in family_usage:
                        # Family is unused!
                        item = self.create_item_dict(family, {
                            'type': 'Detail Component',
                            'type_count': len(list(symbol_ids)),
                            'instance_count': 0
                        })
                        unused_items.append(item)
                
                except Exception as e:
                    # Skip problematic families
                    continue
            
            print("DEBUG: Found {} unused detail components".format(len(unused_items)))
        
        except Exception as e:
            print("ERROR scanning detail components: {}".format(str(e)))
            import traceback
            traceback.print_exc()
        
        return unused_items


class UnusedFamiliesScanner(BasePurgeScanner):
    """Scanner for families with no instances placed"""
    
    def scan(self):
        """Scan for families with no instances in the project"""
        unused_items = []
        
        try:
            print("DEBUG: Starting UnusedFamiliesScanner...")
            
            # OPTIMIZATION: Build family usage dictionary once
            print("DEBUG: Building family usage dictionary...")
            family_usage = {}
            instances = FilteredElementCollector(self.doc)\
                .OfClass(FamilyInstance)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            instance_list = list(instances)
            print("DEBUG: Found {} instances total".format(len(instance_list)))
            
            for inst in instance_list:
                try:
                    if inst and inst.IsValidObject and inst.Symbol:
                        symbol = inst.Symbol
                        if symbol and symbol.Family:
                            family_id = symbol.Family.Id.IntegerValue
                            family_usage[family_id] = True
                except:
                    continue
            
            print("DEBUG: {} families are used".format(len(family_usage)))
            
            # Get all families
            collector = FilteredElementCollector(self.doc)
            families = collector.OfClass(Family).ToElements()
            
            family_list = list(families)
            print("DEBUG: Checking {} families...".format(len(family_list)))
            
            for i, family in enumerate(family_list):
                try:
                    if i % 50 == 0 and i > 0:
                        print("DEBUG: Progress: {}/{}".format(i, len(family_list)))
                    
                    # Skip if null or invalid
                    if not family or not family.IsValidObject:
                        continue
                    
                    # Skip system families (walls, floors, etc.)
                    if not family.FamilyCategory:
                        continue
                    
                    # Skip annotation families (handled by separate scanner)
                    try:
                        cat_id = family.FamilyCategory.Id.IntegerValue
                        # Skip common annotation categories
                        annotation_cats = [
                            int(BuiltInCategory.OST_TextNotes),
                            int(BuiltInCategory.OST_Dimensions),
                            int(BuiltInCategory.OST_Tags),
                            int(BuiltInCategory.OST_GenericAnnotation)
                        ]
                        if cat_id in annotation_cats:
                            continue
                    except:
                        pass
                    
                    # Get all types (symbols) in this family
                    symbol_ids = family.GetFamilySymbolIds()
                    if not symbol_ids or len(list(symbol_ids)) == 0:
                        continue
                    
                    # Check if family is used (quick lookup)
                    family_id = family.Id.IntegerValue
                    if family_id not in family_usage:
                        # Family is unused!
                        category_name = "Unknown"
                        try:
                            if family.FamilyCategory:
                                category_name = family.FamilyCategory.Name
                        except:
                            pass
                        
                        item = self.create_item_dict(family, {
                            'type': 'Family',
                            'category': category_name,
                            'type_count': len(list(symbol_ids)),
                            'instance_count': 0
                        })
                        unused_items.append(item)
                
                except Exception as e:
                    # Skip problematic families
                    continue
            
            print("DEBUG: Found {} unused families".format(len(unused_items)))
        
        except Exception as e:
            print("ERROR scanning unused families: {}".format(str(e)))
            import traceback
            traceback.print_exc()
        
        return unused_items


class UnusedFamilyTypesScanner(BasePurgeScanner):
    """Scanner for family types (symbols) with no instances"""
    
    def scan(self):
        """Scan for family types that have no instances placed"""
        unused_items = []
        
        try:
            print("DEBUG: Starting UnusedFamilyTypesScanner...")
            
            # OPTIMIZATION: Build usage dictionary once
            print("DEBUG: Building symbol usage dictionary...")
            symbol_usage = {}
            instances = FilteredElementCollector(self.doc)\
                .OfClass(FamilyInstance)\
                .WhereElementIsNotElementType()\
                .ToElements()
            
            instance_list = list(instances)
            print("DEBUG: Found {} instances total".format(len(instance_list)))
            
            for inst in instance_list:
                try:
                    if inst and inst.IsValidObject and inst.Symbol:
                        symbol_id = inst.Symbol.Id.IntegerValue
                        symbol_usage[symbol_id] = True
                except:
                    continue
            
            print("DEBUG: {} symbols are used".format(len(symbol_usage)))
            
            # Get all family symbols (types)
            collector = FilteredElementCollector(self.doc)
            symbols = collector.OfClass(FamilySymbol).ToElements()
            
            symbol_list = list(symbols)
            print("DEBUG: Checking {} symbols...".format(len(symbol_list)))
            
            for i, symbol in enumerate(symbol_list):
                try:
                    if i % 100 == 0 and i > 0:
                        print("DEBUG: Progress: {}/{}".format(i, len(symbol_list)))
                    
                    # Skip if null or invalid
                    if not symbol or not symbol.IsValidObject:
                        continue
                    
                    # Skip system families
                    try:
                        family = symbol.Family
                        if not family or not family.FamilyCategory:
                            continue
                    except:
                        continue
                    
                    # Check if symbol is used (quick lookup)
                    symbol_id = symbol.Id.IntegerValue
                    if symbol_id not in symbol_usage:
                        # Symbol is unused!
                        # Get family and category info
                        family_name = "Unknown"
                        category_name = "Unknown"
                        try:
                            family = symbol.Family
                            if family:
                                family_name = family.Name
                                if family.FamilyCategory:
                                    category_name = family.FamilyCategory.Name
                        except:
                            pass
                        
                        item = self.create_item_dict(symbol, {
                            'type': 'Family Type',
                            'family': family_name,
                            'category': category_name,
                            'instance_count': 0
                        })
                        unused_items.append(item)
                
                except Exception as e:
                    # Skip problematic types
                    continue
            
            print("DEBUG: Found {} unused family types".format(len(unused_items)))
        
        except Exception as e:
            print("ERROR scanning unused family types: {}".format(str(e)))
            import traceback
            traceback.print_exc()
        
        return unused_items


class AnnotationFamiliesScanner(BasePurgeScanner):
    """Scanner for unused annotation families (tags, symbols, etc.)"""
    
    def scan(self):
        """Scan for annotation families with no instances"""
        unused_items = []
        
        try:
            print("DEBUG: Starting AnnotationFamiliesScanner...")
            
            # Annotation categories to check
            annotation_categories = [
                BuiltInCategory.OST_GenericAnnotation,
                BuiltInCategory.OST_Tags,
                BuiltInCategory.OST_Callouts,
                BuiltInCategory.OST_DoorTags,
                BuiltInCategory.OST_WindowTags,
                BuiltInCategory.OST_RoomTags,
                BuiltInCategory.OST_WallTags,
                BuiltInCategory.OST_AreaTags,
                BuiltInCategory.OST_SpaceTags
            ]
            
            # OPTIMIZATION: Build family usage dictionary once
            print("DEBUG: Building annotation family usage dictionary...")
            family_usage = {}
            
            try:
                instances = FilteredElementCollector(self.doc)\
                    .OfClass(FamilyInstance)\
                    .WhereElementIsNotElementType()\
                    .ToElements()
                
                for inst in instances:
                    try:
                        if inst and inst.IsValidObject and inst.Symbol:
                            symbol = inst.Symbol
                            if symbol and symbol.Family:
                                family_id = symbol.Family.Id.IntegerValue
                                family_usage[family_id] = True
                    except:
                        continue
            except Exception as e:
                print("WARNING: Could not get instances: {}".format(str(e)))
            
            print("DEBUG: {} families are used".format(len(family_usage)))
            
            # Get all families
            collector = FilteredElementCollector(self.doc)
            families = collector.OfClass(Family).ToElements()
            
            family_list = list(families)
            print("DEBUG: Checking {} families for annotations...".format(len(family_list)))
            
            checked_count = 0
            for family in family_list:
                try:
                    # Skip if null or invalid
                    if not family or not family.IsValidObject:
                        continue
                    
                    # Check if family is annotation type
                    if not family.FamilyCategory:
                        continue
                    
                    cat_id = family.FamilyCategory.Id.IntegerValue
                    is_annotation = any(cat_id == int(cat) for cat in annotation_categories)
                    
                    if not is_annotation:
                        continue
                    
                    checked_count += 1
                    
                    # Get all types in this family
                    symbol_ids = family.GetFamilySymbolIds()
                    if not symbol_ids or len(list(symbol_ids)) == 0:
                        continue
                    
                    # Check if family is used (quick lookup)
                    family_id = family.Id.IntegerValue
                    if family_id not in family_usage:
                        # Family is unused!
                        category_name = "Annotation"
                        try:
                            category_name = family.FamilyCategory.Name
                        except:
                            pass
                        
                        item = self.create_item_dict(family, {
                            'type': 'Annotation',
                            'category': category_name,
                            'type_count': len(list(symbol_ids)),
                            'instance_count': 0
                        })
                        unused_items.append(item)
                
                except Exception as e:
                    # Skip problematic families
                    continue
            
            print("DEBUG: Checked {} annotation families, found {} unused".format(checked_count, len(unused_items)))
        
        except Exception as e:
            print("ERROR scanning annotation families: {}".format(str(e)))
            import traceback
            traceback.print_exc()
        
        return unused_items


class ProfileFamiliesScanner(BasePurgeScanner):
    """Scanner for unused profile families"""
    
    def scan(self):
        """Scan for profile families not used in any elements"""
        unused_items = []
        
        try:
            print("DEBUG: Starting ProfileFamiliesScanner...")
            
            # Get all families
            try:
                collector = FilteredElementCollector(self.doc)
                families = list(collector.OfClass(Family).ToElements())
            except Exception as e:
                print("ERROR: Cannot get families: {}".format(str(e)))
                return unused_items
            
            # Filter to profile families
            profile_families = []
            for family in families:
                try:
                    if not family or not family.IsValidObject:
                        continue
                    
                    # Check if family category is profiles
                    if family.FamilyCategory:
                        try:
                            cat_id = family.FamilyCategory.Id.IntegerValue
                            if cat_id == int(BuiltInCategory.OST_ProfileFamilies):
                                profile_families.append(family)
                        except:
                            continue
                except:
                    continue
            
            print("DEBUG: Found {} profile families".format(len(profile_families)))
            
            if len(profile_families) == 0:
                print("DEBUG: No profile families in project")
                return unused_items
            
            # Get all profile usage (sweeps, reveals, railings, etc.)
            # Profile usage is complex - check multiple element types
            used_profile_ids = set()
            
            # Check wall sweeps - VERY DEFENSIVE
            try:
                print("DEBUG: Checking wall sweeps...")
                try:
                    sweeps = FilteredElementCollector(self.doc)\
                        .OfCategory(BuiltInCategory.OST_Cornices)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
                    
                    sweep_list = list(sweeps)
                    print("DEBUG: Found {} sweeps".format(len(sweep_list)))
                    
                    for sweep in sweep_list:
                        try:
                            if not sweep or not sweep.IsValidObject:
                                continue
                            
                            # Get profile used by sweep
                            try:
                                profile_param = sweep.get_Parameter(DB.BuiltInParameter.WALL_SWEEP_PROFILE)
                                if profile_param:
                                    profile_id = profile_param.AsElementId()
                                    if profile_id and profile_id != ElementId.InvalidElementId:
                                        # Get family from symbol
                                        try:
                                            profile_symbol = self.doc.GetElement(profile_id)
                                            if profile_symbol and hasattr(profile_symbol, 'Family'):
                                                if profile_symbol.Family:
                                                    used_profile_ids.add(profile_symbol.Family.Id.IntegerValue)
                                        except:
                                            pass
                            except:
                                pass
                        except:
                            continue
                except Exception as e:
                    print("DEBUG: Cannot access wall sweeps (might not exist): {}".format(str(e)))
            except Exception as e:
                print("DEBUG: Wall sweep check completely failed: {}".format(str(e)))
            
            # Check reveals - VERY DEFENSIVE
            try:
                print("DEBUG: Checking reveals...")
                try:
                    reveals = FilteredElementCollector(self.doc)\
                        .OfCategory(BuiltInCategory.OST_Reveals)\
                        .WhereElementIsNotElementType()\
                        .ToElements()
                    
                    reveal_list = list(reveals)
                    print("DEBUG: Found {} reveals".format(len(reveal_list)))
                    
                    for reveal in reveal_list:
                        try:
                            if not reveal or not reveal.IsValidObject:
                                continue
                            
                            try:
                                profile_param = reveal.get_Parameter(DB.BuiltInParameter.REVEAL_PROFILE)
                                if profile_param:
                                    profile_id = profile_param.AsElementId()
                                    if profile_id and profile_id != ElementId.InvalidElementId:
                                        try:
                                            profile_symbol = self.doc.GetElement(profile_id)
                                            if profile_symbol and hasattr(profile_symbol, 'Family'):
                                                if profile_symbol.Family:
                                                    used_profile_ids.add(profile_symbol.Family.Id.IntegerValue)
                                        except:
                                            pass
                            except:
                                pass
                        except:
                            continue
                except Exception as e:
                    print("DEBUG: Cannot access reveals (might not exist): {}".format(str(e)))
            except Exception as e:
                print("DEBUG: Reveal check completely failed: {}".format(str(e)))
            
            print("DEBUG: Found {} profiles in use".format(len(used_profile_ids)))
            
            # Check each profile family
            for family in profile_families:
                try:
                    family_id = family.Id.IntegerValue
                    
                    # If not in used set, it's unused
                    if family_id not in used_profile_ids:
                        try:
                            symbol_ids = family.GetFamilySymbolIds()
                            type_count = len(list(symbol_ids)) if symbol_ids else 0
                        except:
                            type_count = 0
                        
                        try:
                            item = self.create_item_dict(family, {
                                'type': 'Profile',
                                'type_count': type_count,
                                'used_in': 'None'
                            })
                            unused_items.append(item)
                        except Exception as e:
                            print("DEBUG: Could not create item for profile: {}".format(str(e)))
                            continue
                
                except Exception as e:
                    # Skip problematic profiles
                    continue
            
            print("DEBUG: Found {} unused profiles".format(len(unused_items)))
        
        except Exception as e:
            print("ERROR scanning profile families: {}".format(str(e)))
            import traceback
            traceback.print_exc()
        
        return unused_items