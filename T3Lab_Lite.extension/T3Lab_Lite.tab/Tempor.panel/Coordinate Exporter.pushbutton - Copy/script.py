# -*- coding: utf-8 -*-
"""
Revit Base Points Retriever

Simple script to retrieve Project Base Point and Survey Point coordinates.
Direct retrieval without unnecessary processing.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Revit Base Points Retriever"

# IMPORT LIBRARIES
# ==================================================

from Autodesk.Revit.DB import *


# CLASS/FUNCTIONS
# ==================================================

class BasePointsRetriever:
    
    def __init__(self, doc):
        self.doc = doc
    
    def get_base_points(self):
        result = {}
        
        # Get Project Base Point
        project_bp = BasePoint.GetProjectBasePoint(self.doc)
        if project_bp is not None:
            result['project_base_point'] = project_bp
        
        # Get Survey Point
        survey_bp = BasePoint.GetSurveyPoint(self.doc)
        if survey_bp is not None:
            result['survey_point'] = survey_bp
        
        return result


# MAIN SCRIPT
# ==================================================

def main(): 
    # Get active document
    doc = __revit__.ActiveUIDocument.Document
    
    # Retrieve base points
    retriever = BasePointsRetriever(doc)
    base_points = retriever.get_base_points()
    
    # Display results
    print("\n" + "=" * 70)
    print("BASE POINTS RETRIEVAL RESULTS")
    print("=" * 70)
    
    if 'project_base_point' in base_points:
        pbp = base_points['project_base_point']
        print("\nProject Base Point:")
        print("  Element ID: {}".format(pbp.Id))
        print("  Element Name: {}".format(pbp.Name if hasattr(pbp, 'Name') else 'N/A'))
        print("  Object: {}".format(pbp))
        print("  XYZ: {}".format(pbp.Position))
    
    if 'survey_point' in base_points:
        sp = base_points['survey_point']
        print("\nSurvey Point:")
        print("  Element ID: {}".format(sp.Id))
        print("  Element Name: {}".format(sp.Name if hasattr(sp, 'Name') else 'N/A'))
        print("  Object: {}".format(sp))
        print("  XYZ: {}".format(sp.Position))
    
    print("\n" + "=" * 70 + "\n")
    
    return base_points


# Execute main script
if __name__ == '__main__':
    main()