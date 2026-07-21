---
name: lod-standard
description: LOD 100-500 and LOIN - how detailed the model must be per stage
triggers: lod, lod 200, lod 300, lod 350, lod 400, lod 500, muc do chi tiet, level of detail, loin, level of development
agents: general, knowledge, revit_data
tools: analyze_model_statistics
---
# Chuẩn LOD / LOIN (T3Lab)

## Thang LOD (Level of Development)
| LOD | Nội dung hình học | Dùng cho |
|-----|-------------------|----------|
| 100 | Khối dạng ý tưởng, ký hiệu | Concept, khái toán sơ bộ |
| 200 | Hình dạng-kích thước-vị trí GẦN đúng | Thiết kế cơ sở |
| 300 | Hình học chính xác, đủ để phối hợp | Hồ sơ thiết kế kỹ thuật; từ mức này mới chạy clash detection có ý nghĩa |
| 350 | LOD 300 + chi tiết liên kết giữa các hệ | Phối hợp liên bộ môn sâu |
| 400 | Chi tiết chế tạo, lắp đặt | Shop drawing; clash ở mức này = chi phí thật |
| 500 | Nghiệm thu thực tế (as-built) | Bàn giao vận hành |

## LOIN (Level of Information Need — ISO 19650)
- LOD chỉ nói HÌNH HỌC; LOIN yêu cầu thêm: thông tin phi hình học (parameter, mã tài sản) + tài liệu đính kèm — chỉ yêu cầu thông tin THỰC SỰ CẦN cho mục đích sử dụng, tránh over-modeling.
- Mỗi loại cấu kiện có thể ở LOD khác nhau trong cùng một model — quy định trong BEP bằng bảng ma trận cấu kiện × giai đoạn.

## Khi người dùng hỏi "model này cần LOD bao nhiêu"
1. Hỏi mục đích: khái toán / hồ sơ / phối hợp / thi công / bàn giao FM.
2. Trả lời theo bảng trên + nhắc: nâng LOD là nâng chi phí modeling — chỉ nâng cấu kiện cần thiết.
3. Nếu cần đánh giá model hiện tại: `analyze_model_statistics` xem mật độ element/family để nhận định sơ bộ mức chi tiết.
