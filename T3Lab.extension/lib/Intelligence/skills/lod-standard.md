---
name: lod-standard
description: LOD 100-500 and LOIN - how detailed the model must be per stage
triggers: lod, lod 100, lod 200, lod 300, lod 350, lod 400, lod 500, muc do chi tiet, level of detail, loin, level of development, do chi tiet model
agents: general, knowledge, revit_data
tools: analyze_model_statistics
standard: project
---
# Chuẩn LOD / LOIN (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Thang LOD (Level of Development) — MẪU THAM KHẢO, chốt theo ma trận LOD trong BEP
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
