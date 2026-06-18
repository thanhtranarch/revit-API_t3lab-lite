# -*- coding: utf-8 -*-
"""
Import các trang PDF vào các Views đang được chọn trong Project Browser.
Thứ tự: Trang 1 -> View đầu tiên, Trang 2 -> View thứ 2...
"""
__title__ = "PDF To\nSelected Views"

import clr
from pyrevit import revit, DB, forms, script

doc = revit.doc

def main():
    # 1. Lấy danh sách các phần tử đang được chọn
    selection = revit.get_selection()
    
    # Lọc ra chỉ lấy các elements là View (loại bỏ nếu lỡ chọn nhầm Element khác)
    # Đồng thời bỏ qua các View Template
    selected_views = [el for el in selection if isinstance(el, DB.View) and not el.IsTemplate]
    
    total_views = len(selected_views)
    if total_views == 0:
        forms.alert("Vui lòng giữ Ctrl/Shift và chọn các View trong Project Browser trước khi chạy lệnh.", exitscript=True)
        
    # Sắp xếp các view theo Tên (Alphabetical) để đảm bảo thứ tự import đúng từ trên xuống dưới
    selected_views.sort(key=lambda v: v.Name)

    # 2. Chọn file PDF cần import
    pdf_path = forms.pick_file(file_ext='pdf', title="Chọn PDF (Sẽ import {} trang vào {} views đang chọn)".format(total_views, total_views))
    if not pdf_path:
        script.exit()

    # 3. Tiến hành Transaction loop
    imported_count = 0
    with revit.Transaction("Import PDF to Selected Views"):
        for i, view in enumerate(selected_views):
            page_number = i + 1  # Trang 1 ứng với View 1, Trang 2 ứng với View 2...
            
            try:
                # Thiết lập tuỳ chọn Import
                options = DB.ImageTypeOptions(pdf_path, False, DB.ImageTypeSource.Import)
                options.Resolution = 300
                options.PageNumber = page_number
                
                # Tạo ImageType
                image_type = DB.ImageType.Create(doc, options)
                
                # Khởi tạo tùy chọn vị trí đặt
                placement_options = DB.ImagePlacementOptions()
                placement_options.Location = DB.XYZ.Zero
                
                # Đặt Instance vào view tương ứng trong vòng lặp
                DB.ImageInstance.Create(doc, view, image_type.Id, placement_options)
                
                imported_count += 1
                
            except Exception as e:
                # Báo lỗi nếu View đó không hỗ trợ dán ảnh (ví dụ: Schedule, 3D View...)
                print("Lỗi ở trang {} (View '{}'): {}".format(page_number, view.Name, str(e)))

    # 4. Thông báo kết quả
    if imported_count > 0:
        forms.alert("Hoàn tất! Đã import thành công {} trang PDF vào {} views.".format(imported_count, total_views))

if __name__ == '__main__':
    main()