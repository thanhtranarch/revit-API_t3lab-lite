# Contains Manager - pyDQT

![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Revit](https://img.shields.io/badge/Revit-2020--2025-green)
![License](https://img.shields.io/badge/license-Proprietary-orange)

## Gioi thieu

**Contains Manager** la cong cu tim kiem va gan tham so cho cac phan tu nam trong Rooms/Areas/Spaces. Tool nay giup tu dong hoa qua trinh kiem tra vi tri phan tu va gan thong tin phong vao cac tham so instance.

## Tinh nang chinh

### 1. Chon Spatial Elements
- **Rooms**: Phong trong mo hinh Revit
- **Areas**: Vung dien tich
- **Spaces**: Khong gian MEP

### 2. Loc theo Category
- Ho tro tat ca cac category pho bien
- Loc theo Discipline: Architecture, Structure, MEP
- Tim kiem theo ten category

### 3. Tim phan tu trong Spatial
- Su dung thuat toan **Point Containment**
- Kiem tra chinh xac phan tu nam trong Room/Area/Space nao
- Hien thi ket qua theo bang voi day du thong tin

### 4. Gan Parameter Values
- **Room Name**: Gan ten phong
- **Room Number**: Gan so phong  
- **Room Name + Number**: Ghep ten va so phong
- **Level Name**: Gan ten level
- **Custom Value**: Gia tri tuy chinh

### 5. Visualization
- Highlight cac spatial element da chon
- Select elements trong Revit

## Huong dan su dung

### Buoc 1: Chon Spatial Elements
1. Chon **Scope**: Whole Model hoac Active View
2. Chon **Spatial Type**: Rooms, Areas, hoac Spaces
3. Tick chon cac spatial elements can tim

### Buoc 2: Chon Categories
1. Chon **Discipline** de loc (tuy chon)
2. Tick chon cac categories can tim
3. Su dung Search de tim nhanh

### Buoc 3: Tim phan tu
1. Click nut **Find**
2. Ket qua hien thi trong bang Results
3. Xem thong tin Category, Family, Type, Room Name, Room Number

### Buoc 4: Gan Parameter Values
1. Chon **Source Value** (Room Name, Room Number, etc.)
2. Chon **Target Parameter** (parameter instance cua element)
3. Chon **Separator** neu can
4. Click **Set Parameter Value**

### Buoc 5: Thao tac khac
- **Visualize**: Highlight spatial elements trong Revit
- **Select**: Chon cac elements tim duoc trong Revit
- **Reset**: Xoa tat ca lua chon

## Thuat toan tim phan tu

Tool su dung phuong phap **Point Containment**:

1. **Lay vi tri phan tu**: 
   - Family Instance: Lay Location Point
   - Linear elements: Lay midpoint cua Curve
   - Fallback: Lay center cua Bounding Box

2. **Kiem tra containment**:
   - **Rooms**: Su dung `Room.IsPointInRoom()`
   - **Spaces**: Su dung `Space.IsPointInSpace()`
   - **Areas**: Su dung Ray Casting Algorithm voi boundary segments

## Ung dung thuc te

### MEP
- Gan Room Name/Number cho thiet bi co dien
- Kiem tra vi tri thiet bi theo phong
- Xuat schedule thiet bi theo location

### QS (Quantity Surveying)
- Kiem tra vat tu theo tung phong
- Thong ke so luong theo Area/Zone
- Xuat bang khoi luong theo vi tri

### Architecture
- Kiem tra furniture dat dung phong
- Tim elements bi dat sai vi tri
- Quan ly noi that theo phong

### Model QC
- Tim element nam sai zone/area
- Kiem tra vi tri element theo Level
- Xac dinh element chua duoc gan phong

## Luu y quan trong

1. **Rooms/Spaces/Areas phai duoc dat (placed)** - Cac spatial element chua dat se khong hien thi

2. **Area > 0** - Chi hien thi spatial elements co dien tich > 0

3. **Parameter target phai la Instance parameter** - Chi ho tro ghi vao instance parameters

4. **String parameters** - Chi ho tro gan gia tri vao string parameters

## Cau truc thu muc

```
ContainsManager.pushbutton/
├── script.py          # Main script
├── icon.png           # Button icon (96x96)
├── icon.svg           # Source SVG icon
├── bundle.yaml        # PyRevit configuration
└── README.md          # Documentation
```

## Yeu cau he thong

- Revit 2020 - 2025
- pyRevit 4.8+
- IronPython 2.7

## Tac gia

**Dang Quoc Truong (DQT)**

Copyright (c) 2024 DQT. All rights reserved.

## Version History

### v1.0.0 (2024)
- Initial release
- Core features: Find, Set Parameter, Visualize, Select
- Support Rooms, Areas, Spaces
- Multi-category selection
- Search and filter functionality
