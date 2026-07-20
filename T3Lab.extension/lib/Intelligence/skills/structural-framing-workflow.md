---
name: structural-framing-workflow
description: Create structural framing systems the right way
triggers: he dam, dam san, tao dam, khung ket cau, framing system, he ket cau, dam phu, dam chinh
agents: revit_action, revit_data
tools: create_structural_framing_system, create_line_based_element, list_levels, get_available_family_types, create_grid
---
# Quy trình hệ khung kết cấu (T3Lab)

## Chuẩn bị trước khi tạo
1. `list_levels` xác nhận cao độ tầng sẽ đặt dầm (dầm đặt theo cao độ mặt trên bám sàn).
2. `get_available_family_types` kiểm tra family dầm/cột có sẵn — thiếu tiết diện cần dùng thì báo người dùng load family trước (skill `family-workflow`).
3. Lưới trục phải có trước (skill `level-grid-standard`) — dầm chạy theo trục, không vẽ tự do.

## Tạo hệ dầm
- Hệ dầm đều trên một vùng sàn: `create_structural_framing_system` — nêu rõ: vùng biên, phương dầm, khoảng cách/số lượng, tiết diện.
- Dầm đơn lẻ (dầm biên, dầm đỡ cục bộ): `create_line_based_element` với điểm đầu–cuối theo giao trục.
- Kế hoạch trình trước khi tạo: | Vị trí | Phương | Tiết diện | Số lượng | Cao độ |.

## Quy ước
- Tên type theo tiết diện thật: `300x600`, `D400` — không dùng tên mơ hồ kiểu "Beam 1".
- Dầm chính nối cột–cột theo trục; dầm phụ gác lên dầm chính; chiều đặt thống nhất để Mark không lộn xộn.
- Phần mềm này hỗ trợ dựng HÌNH; thiết kế tiết diện/thép là việc của kỹ sư kết cấu — không tự quyết định tiết diện, chỉ dựng theo yêu cầu.

## Sau khi tạo
- Báo số dầm đã tạo theo tầng + nhắc kiểm tra join dầm–cột và warnings mới phát sinh (skill `warning-triage`).
