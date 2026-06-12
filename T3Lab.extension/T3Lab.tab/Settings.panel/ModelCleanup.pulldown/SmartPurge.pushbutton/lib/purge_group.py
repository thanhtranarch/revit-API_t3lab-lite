# -*- coding: utf-8 -*-
"""
Purge Group Management
Manages groups of purge categories
Copyright (c) 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"


class PurgeGroup:
    """Group of related purge categories"""
    
    def __init__(self, id, name, icon, description, default_checked=True):
        """
        Initialize purge group
        
        Args:
            id: Unique group ID (e.g., "element_types")
            name: Display name (e.g., "Element Types")
            icon: Unicode icon for display
            description: Group description
            default_checked: Whether group is checked by default
        """
        self.id = id
        self.name = name
        self.icon = icon
        self.description = description
        self.default_checked = default_checked
        
        # Runtime state
        self.categories = []           # List of PurgeCategory objects
        self.is_checked = default_checked
        self.is_scanned = False
        self.is_expanded = False       # For UI collapsible state
        
        # Statistics
        self.scanned_count = 0         # Number of categories scanned
        self.unused_count = 0          # Total unused items in group
        self.total_categories = 0      # Total categories in group
    
    def add_category(self, category):
        """Add category to this group"""
        self.categories.append(category)
        self.total_categories = len(self.categories)
        category.group_id = self.id
    
    def get_selected_categories(self):
        """Get list of checked categories in this group"""
        return [cat for cat in self.categories if cat.is_checked]
    
    def get_scanned_categories(self):
        """Get list of scanned categories in this group"""
        return [cat for cat in self.categories if cat.is_scanned]
    
    def update_statistics(self):
        """Update group statistics from categories"""
        self.scanned_count = len(self.get_scanned_categories())
        self.unused_count = sum(len(cat.unused_items) 
                               for cat in self.categories 
                               if cat.is_scanned)
    
    def get_scan_status_text(self):
        """Get scan status text for display"""
        if not self.is_scanned:
            return "Not scanned"
        
        scanned = self.scanned_count
        total = self.total_categories
        
        if scanned == 0:
            return "Not scanned"
        elif scanned < total:
            return "{}/{} scanned".format(scanned, total)
        else:
            return "All scanned"
    
    def get_unused_count_text(self):
        """Get unused items count text"""
        if not self.is_scanned:
            return ""
        
        count = self.unused_count
        if count == 0:
            return "No items"
        elif count == 1:
            return "1 item"
        else:
            return "{} items".format(count)
    
    def select_all(self):
        """Select all categories in group"""
        for cat in self.categories:
            cat.is_checked = True
    
    def select_none(self):
        """Deselect all categories in group"""
        for cat in self.categories:
            cat.is_checked = False
    
    def select_safe_only(self):
        """Select only safe categories"""
        for cat in self.categories:
            if cat.safety_level == "SAFE":
                cat.is_checked = True
            else:
                cat.is_checked = False
    
    def reset_scan_state(self):
        """Reset scan state for this group"""
        self.is_scanned = False
        self.scanned_count = 0
        self.unused_count = 0
        
        for cat in self.categories:
            cat.is_scanned = False
            cat.unused_items = []


# Group definitions
PURGE_GROUPS = {
    "element_types": {
        "id": "element_types",
        "name": "Element Types",
        "icon": u"\U0001F4E6",  # 📦
        "description": "Materials, patterns, text styles, dimension styles, and system types",
        "default_checked": True
    },
    
    "views_sheets": {
        "id": "views_sheets",
        "name": "Views & Sheets",
        "icon": u"\U0001F441",  # 👁
        "description": "View templates, filters, empty sheets, schedules, and legend views",
        "default_checked": True
    },
    
    "families": {
        "id": "families",
        "name": "Families & Types",
        "icon": u"\U0001F3DB",  # 🏛
        "description": "Detail components, unused families, family types, annotations, and profiles",
        "default_checked": False
    },
    
    "system_cleanup": {
        "id": "system_cleanup",
        "name": "System Cleanup",
        "icon": u"\U0001F9F9",  # 🧹
        "description": "Import symbols, CAD links, unused groups, design options, room separators, orphaned rooms",
        "default_checked": True
    }
}


def create_purge_groups():
    """Create all purge group objects"""
    groups = []
    
    for group_data in PURGE_GROUPS.values():
        group = PurgeGroup(
            id=group_data["id"],
            name=group_data["name"],
            icon=group_data["icon"],
            description=group_data["description"],
            default_checked=group_data["default_checked"]
        )
        groups.append(group)
    
    return groups


def get_group_by_id(groups, group_id):
    """Get group object by ID"""
    for group in groups:
        if group.id == group_id:
            return group
    return None