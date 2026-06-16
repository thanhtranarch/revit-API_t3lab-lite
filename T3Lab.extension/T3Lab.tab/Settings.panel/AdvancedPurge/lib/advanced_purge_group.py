# -*- coding: utf-8 -*-
"""
Advanced Purge Groups
Defines category groups for organization

Copyright © 2025 Dang Quoc Truong (DQT)
"""

__author__ = "Dang Quoc Truong (DQT)"

from config_advanced import Icons
from advanced_purge_categories import (
    UNREFERENCED_VIEWS,
    WORKSET_CLEANUP,
    MODEL_DEEP_CLEANUP,
    COLLABORATION_CLEANUP,
    DANGEROUS_OPERATIONS
)


class AdvancedPurgeGroup(object):
    """Represents a group of categories"""
    
    def __init__(self, id, name, icon, description, categories):
        """
        Initialize group
        
        Args:
            id: Unique identifier
            name: Display name
            icon: Unicode icon
            description: Description for tooltip
            categories: List of categories in this group
        """
        self.id = id
        self.name = name
        self.icon = icon
        self.description = description
        self.categories = categories


def create_advanced_purge_groups(doc=None):
    """
    Create all advanced purge groups
    
    Args:
        doc: Revit document (optional, required for dynamic workset categories)
    """
    
    # Get dynamic workset categories if doc provided
    workset_categories = []
    if doc:
        from advanced_purge_categories import get_workset_cleanup_categories
        workset_categories = get_workset_cleanup_categories(doc)
    
    groups = [
        AdvancedPurgeGroup(
            id="advanced_views",
            name="Advanced Views",
            icon=Icons.VIEWS,
            description="Unreferenced views by type (not placed on sheets)",
            categories=UNREFERENCED_VIEWS
        ),
        
        AdvancedPurgeGroup(
            id="workset_cleanup",
            name="Workset Cleanup",
            icon=Icons.WORKSET,
            description="Remove elements on specific worksets (requires workshared model)",
            categories=workset_categories  # Dynamic categories
        ),
        
        AdvancedPurgeGroup(
            id="model_deep",
            name="Model Deep Cleanup",
            icon=Icons.MODEL,
            description="Deep model cleanup: elevation markers, scope boxes, constraints, etc.",
            categories=MODEL_DEEP_CLEANUP
        ),
        
        # Skip Collaboration if empty
        # AdvancedPurgeGroup(
        #     id="collaboration",
        #     name="Collaboration Cleanup",
        #     icon=Icons.COLLAB,
        #     description="BIM360 cache, data schema, subcategories",
        #     categories=COLLABORATION_CLEANUP
        # ),
        
        AdvancedPurgeGroup(
            id="dangerous",
            name="Dangerous Operations",
            icon=Icons.DANGEROUS,
            description="DANGEROUS operations that can significantly modify the model",
            categories=DANGEROUS_OPERATIONS
        ),
    ]
    
    return groups


def get_group_by_id(groups, group_id):
    """Get group by ID"""
    for group in groups:
        if group.id == group_id:
            return group
    return None


def get_all_categories(groups):
    """Get all categories from all groups"""
    categories = []
    for group in groups:
        categories.extend(group.categories)
    return categories


def get_category_by_id(groups, category_id):
    """Get category by ID"""
    for group in groups:
        for category in group.categories:
            if category.id == category_id:
                return category
    return None