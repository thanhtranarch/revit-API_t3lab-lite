# Text to Element Transfer Tool

Công cụ PyRevit để transfer giá trị từ Text Notes vào các element khi text giao với element đó.

## Mô tả

Tool này giúp tự động hóa việc gán thông tin từ annotation (Text Notes) vào các model elements hoặc detail components dựa trên vị trí giao nhau của chúng trong view.

## Có 2 phiên bản

### 1. Full Version (`text_to_element_transfer/`)

Phiên bản đầy đủ với giao diện WPF:

**Tính năng:**
- Giao diện đồ họa theo style pyDQT (yellow-orange theme)
- Chọn text notes bằng cách pick từ model hoặc lấy tất cả trong view
- Chọn category đích (Walls, Floors, Rooms, Doors, etc.)
- Chọn parameter đích (Comments, Mark, Description, etc.) hoặc nhập custom parameter
- Điều chỉnh tolerance cho intersection detection
- Preview trước khi transfer
- Hiển thị kết quả chi tiết

**Workflow:**
1. Step 1: Chọn Text Notes
2. Step 2: Chọn Category của elements đích
3. Step 3: Chọn Parameter để ghi vào
4. Step 4: Preview và Transfer

### 2. Quick Version (`text_to_element_transfer_quick/`)

Phiên bản nhanh, không cần WPF:

**Tính năng:**
- Làm việc trực tiếp với selection
- Tự động tìm elements giao với text notes
- Chọn parameter từ danh sách có sẵn
- Confirm trước khi transfer

**Workflow:**
1. Chọn text notes (và optionally các elements đích)
2. Chạy script
3. Chọn target parameter
4. Confirm và transfer

## Cài đặt

1. Copy folder tool vào extension của bạn:
```
YourExtension.extension/
└── YourTab.tab/
    └── YourPanel.panel/
        └── TextToElement.pushbutton/
            ├── script.py
            └── bundle.yaml
```

2. Reload PyRevit hoặc restart Revit

## Cách sử dụng

### Use Case 1: Gán tên phòng từ Text vào Room
1. Tạo Text Notes với tên phòng đặt bên trong các Room boundaries
2. Chạy tool
3. Chọn category "Rooms"
4. Chọn parameter "Comments" hoặc "Name"
5. Transfer

### Use Case 2: Gán mã chi tiết từ Text vào Detail Items
1. Tạo Text Notes với mã detail đặt gần các Detail Components
2. Chạy tool
3. Chọn category "Detail Items"
4. Chọn parameter "Mark"
5. Transfer

### Use Case 3: Gán thông tin từ Text vào Walls
1. Chọn các Text Notes có nội dung cần transfer
2. Chạy tool
3. Chọn category "Walls"
4. Nhập custom parameter name
5. Transfer

## Lưu ý kỹ thuật

### Intersection Detection
- Tool sử dụng bounding box intersection để xác định text giao với element
- Tolerance mặc định: 0.5 feet
- Chỉ xét intersection trong 2D (X, Y) - phù hợp với plan views

### Parameter Support
- Instance parameters (Comments, Mark, etc.)
- Built-in parameters (ALL_MODEL_INSTANCE_COMMENTS, ALL_MODEL_MARK)
- Custom shared parameters
- KHÔNG hỗ trợ read-only parameters hoặc type parameters

### Categories hỗ trợ
- Walls, Floors, Ceilings, Roofs
- Rooms, Areas
- Doors, Windows
- Furniture, Generic Models
- Structural Framing, Columns
- MEP Equipment
- Detail Items, Casework

## API Reference

```python
# Core functions
get_text_content(text_note)              # Lấy nội dung text
get_text_note_bounding_box(text_note, view)  # Lấy bounding box
boxes_intersect(bb1, bb2, tolerance)     # Kiểm tra intersection
set_parameter_value(element, param, value)   # Gán giá trị parameter
```

## Troubleshooting

**"No intersections found"**
- Đảm bảo text notes nằm chồng lên elements trong view hiện tại
- Thử tăng tolerance trong Full version
- Kiểm tra elements có visible trong view không

**"Parameter not found or read-only"**
- Parameter có thể là read-only
- Parameter không tồn tại trên element type đó
- Thử parameter khác như "Comments"

**"No target elements found"**
- Chọn đúng category
- Đảm bảo elements visible trong view
- Một số categories có thể không có elements trong view

## Version History

- v1.0: Initial release
  - Full version với WPF UI
  - Quick version cho workflow nhanh
  - Support 17 categories
  - Support custom parameters

## Author

DQT - pyDQT Tools
