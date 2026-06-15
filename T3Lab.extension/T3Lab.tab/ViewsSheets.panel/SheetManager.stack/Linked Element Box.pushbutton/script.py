# -*- coding: utf-8 -*-
from pyrevit import revit, DB, UI
from pyrevit.revit import Transaction

uidoc = revit.uidoc
doc = revit.doc

try:
    # Bước 1: Chọn đối tượng trong Linked Model
    linked_ref = uidoc.Selection.PickObject(UI.Selection.ObjectType.LinkedElement, "Chọn đối tượng trong Linked Model")
    link_instance = doc.GetElement(linked_ref.ElementId)

    # Kiểm tra xem có phải là RevitLinkInstance không
    if isinstance(link_instance, DB.RevitLinkInstance):
        # Bước 2: Lấy document của Linked Model
        link_doc = link_instance.GetLinkDocument()
        linked_element = link_doc.GetElement(linked_ref.LinkedElementId)

        # Bước 3: Lấy bounding box của đối tượng trong linked
        bbox = linked_element.get_BoundingBox(None)
        if bbox:
            # Bước 4: Chuyển bounding box sang tọa độ của model chính
            transform = link_instance.GetTransform()
            min_pt = transform.OfPoint(bbox.Min)
            max_pt = transform.OfPoint(bbox.Max)

            # Mở rộng thêm 500mm (0.5m) mỗi hướng
            offset = DB.XYZ(0.5, 0.5, 0.5)
            bbox_transformed = DB.BoundingBoxXYZ()
            bbox_transformed.Min = min_pt - offset
            bbox_transformed.Max = max_pt + offset

            # Bước 5: Tìm View 3D mặc định (tên là "{3D}")
            default_3d_view = None
            views_3d = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
            for v in views_3d:
                if not v.IsTemplate and v.Name.strip() == "{3D}":
                    default_3d_view = v
                    break

            if default_3d_view:
                with Transaction("Gán Selection Box vào View 3D mặc định"):
                    default_3d_view.SetSectionBox(bbox_transformed)

                # Chuyển sang View mới sau khi Transaction đã đóng
                uidoc.ActiveView = default_3d_view
                UI.TaskDialog.Show("Thành công", "Đã gán Selection Box mở rộng 500mm vào View 3D mặc định.")
            else:
                UI.TaskDialog.Show("Lỗi", "Không tìm thấy View 3D mặc định có tên '{3D}'.")
        else:
            UI.TaskDialog.Show("Lỗi", "Không thể lấy BoundingBox của đối tượng.")
    else:
        UI.TaskDialog.Show("Lỗi", "Đối tượng được chọn không phải là RevitLinkInstance.")

except Exception as e:
    UI.TaskDialog.Show("Lỗi", "Không thể thực hiện: {}".format(str(e)))