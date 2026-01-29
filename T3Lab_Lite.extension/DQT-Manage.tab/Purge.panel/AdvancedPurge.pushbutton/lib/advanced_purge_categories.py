# -*- coding: utf-8 -*-
"""
Advanced Purge Categories
Defines all advanced purge categories and their properties

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

from config_advanced import Icons
import Autodesk.Revit.DB


class AdvancedPurgeCategory(object):
    """Represents an advanced purge category"""
    
    def __init__(self, id, name, description, is_dangerous=False, 
                 requires_worksets=False, scanner_class=None):
        """
        Initialize category
        
        Args:
            id: Unique identifier
            name: Display name
            description: Description for tooltip
            is_dangerous: Whether this is a dangerous operation
            requires_worksets: Whether this requires workshared model
            scanner_class: Scanner class name (string)
        """
        self.id = id
        self.name = name
        self.description = description
        self.is_dangerous = is_dangerous
        self.requires_worksets = requires_worksets
        self.scanner_class = scanner_class
        self.count = 0
        self.status = None  # None, 'safe', 'warning', 'error'


# ============================================================================
# GROUP 1: ADVANCED VIEWS (9 categories - Unreferenced only)
# ============================================================================

UNREFERENCED_VIEWS = [
    AdvancedPurgeCategory(
        id="unreferenced_3d",
        name="Unreferenced 3D Views",
        description="3D views not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_area",
        name="Unreferenced Area Plans",
        description="Area plans not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_detail",
        name="Unreferenced Detail Views",
        description="Detail views not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_drafting",
        name="Unreferenced Drafting Views",
        description="Drafting views not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_elevation",
        name="Unreferenced Elevations",
        description="Elevation views not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_engineering",
        name="Unreferenced Engineering Plans",
        description="Engineering plans not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_floor",
        name="Unreferenced Floor Plans",
        description="Floor plans not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_rcp",
        name="Unreferenced Ceiling Plans",
        description="Reflected ceiling plans not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
    AdvancedPurgeCategory(
        id="unreferenced_section",
        name="Unreferenced Sections",
        description="Section views not placed on any sheet",
        scanner_class="UnreferencedViewsScanner"
    ),
]


# ============================================================================
# GROUP 2: WORKSET CLEANUP (Dynamic categories based on model worksets)
# ============================================================================

def get_workset_cleanup_categories(doc):
    """
    Generate workset cleanup categories dynamically based on model worksets
    
    Args:
        doc: Revit document
        
    Returns:
        List of AdvancedPurgeCategory for each workset in the model
    """
    categories = []
    
    print("=" * 80)
    print("DEBUG: get_workset_cleanup_categories called")
    
    # Check if document is workshared
    if not doc.IsWorkshared:
        print("DEBUG: Document is NOT workshared - returning empty list")
        print("=" * 80)
        return categories
    
    print("DEBUG: Document IS workshared - scanning worksets...")
    
    try:
        # Use FilteredWorksetCollector to get all worksets
        from Autodesk.Revit.DB import FilteredWorksetCollector, WorksetKind
        
        workset_collector = FilteredWorksetCollector(doc)
        all_worksets = workset_collector.OfKind(WorksetKind.UserWorkset).ToWorksets()
        
        print("DEBUG: Found {} user worksets".format(len(list(all_worksets))))
        
        for workset in all_worksets:
            workset_name = workset.Name
            workset_id = workset.Id
            workset_kind = workset.Kind
            
            print("DEBUG: Workset '{}' - ID: {} - Kind: {}".format(
                workset_name, workset_id.IntegerValue, workset_kind
            ))
            
            # Create category for this workset
            category = AdvancedPurgeCategory(
                id="workset_{}".format(workset_id.IntegerValue),
                name='Elements on "{}"'.format(workset_name),
                description="Remove all elements on {} workset".format(workset_name),
                requires_worksets=True,
                is_dangerous=True,
                scanner_class="WorksetScanner"
            )
            
            # Store workset_id in category for scanner to use
            category.workset_id = workset_id
            category.workset_name = workset_name
            
            categories.append(category)
            print("  -> ADDED as category")
        
        print("DEBUG: Created {} workset categories".format(len(categories)))
        print("=" * 80)
            
    except Exception as e:
        print("ERROR getting worksets: {}".format(str(e)))
        import traceback
        print(traceback.format_exc())
        print("=" * 80)
    
    return categories


# Static empty list for compatibility - will be populated dynamically
WORKSET_CLEANUP = []


# ============================================================================
# GROUP 3: MODEL DEEP CLEANUP (5 categories)
# ============================================================================

MODEL_DEEP_CLEANUP = [
    AdvancedPurgeCategory(
        id="orphaned_elevation_markers",
        name="Orphaned Elevation Markers",
        description="Elevation markers with no views",
        scanner_class="ModelDeepScanner"
    ),
    AdvancedPurgeCategory(
        id="unused_scope_boxes",
        name="Unused Scope Boxes",
        description="Scope boxes not used in any view",
        scanner_class="ModelDeepScanner"
    ),
    AdvancedPurgeCategory(
        id="room_separation_lines",
        name="Room Separation Lines",
        description="All room separation lines",
        is_dangerous=True,
        scanner_class="ModelDeepScanner"
    ),
    AdvancedPurgeCategory(
        id="unnamed_reference_planes",
        name="Unnamed Reference Planes",
        description="Reference planes without names",
        scanner_class="ModelDeepScanner"
    ),
    AdvancedPurgeCategory(
        id="all_reference_planes",
        name="All Reference Planes",
        description="All reference planes in the model",
        is_dangerous=True,
        scanner_class="ModelDeepScanner"
    ),
]


# ============================================================================
# GROUP 4: COLLABORATION CLEANUP (0 categories - All require special implementation)
# ============================================================================

# NOTE: All collaboration cleanup categories require special executor implementation
# - BIM360 cache: File system operations
# - Data schema: ExtensibleStorage API (dangerous)
# - View-specific constraints: Complex iteration
# Removed for initial release

COLLABORATION_CLEANUP = []


# ============================================================================
# GROUP 5: DANGEROUS OPERATIONS (2 categories)
# ============================================================================

DANGEROUS_OPERATIONS = [
    AdvancedPurgeCategory(
        id="area_separation_lines",
        name="All Area Separation Lines",
        description="Remove all area separation lines (DANGEROUS!)",
        is_dangerous=True,
        scanner_class="DangerousOpsScanner"
    ),
    AdvancedPurgeCategory(
        id="all_groups",
        name="All Groups",
        description="Delete all groups in the model (DANGEROUS!)",
        is_dangerous=True,
        scanner_class="DangerousOpsScanner"
    ),
]


# ============================================================================
# ALL CATEGORIES LIST
# ============================================================================

ALL_CATEGORIES = (
    UNREFERENCED_VIEWS + 
    WORKSET_CLEANUP + 
    MODEL_DEEP_CLEANUP + 
    COLLABORATION_CLEANUP + 
    DANGEROUS_OPERATIONS
)