---
name: bep-guideline
description: BIM Execution Plan (BEP) guide per ISO 19650-2 with required sections
triggers: bim execution plan, ke hoach bim, trien khai bim, lap bep, bep guideline, bep la gi, eir, midp, tidp, cde, common data environment, soan bep, ra soat bep
agents: general, knowledge
tools:
standard: project
---
# Lập BIM Execution Plan — BEP (T3Lab, theo ISO 19650-2)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Hai giai đoạn BEP
- **Pre-appointment BEP** (trước ký hợp đồng): trình bày NĂNG LỰC và cách tiếp cận — trả lời EIR của chủ đầu tư.
- **Post-appointment BEP** (sau ký): chốt cách THỰC HIỆN — kèm TIDP (kế hoạch giao nộp theo từng đội) tổng hợp thành MIDP (kế hoạch giao nộp tổng).

## Mục bắt buộc trong BEP (checklist)
1. Thông tin dự án + các mốc giao nộp thông tin (information delivery milestones).
2. EIR/OIR/AIR được tham chiếu — yêu cầu thông tin của chủ đầu tư.
3. Vai trò & trách nhiệm: BIM Manager, BIM Coordinator từng bộ môn, người phê duyệt.
4. CDE (Common Data Environment): nền tảng, cấu trúc thư mục, luồng trạng thái WIP → Shared → Published → Archived.
5. Quy tắc đặt tên file/tài liệu (dùng skill `iso19650-naming`).
6. Mức thông tin cần thiết (LOIN/LOD) theo từng giai đoạn và từng loại cấu kiện.
7. Ma trận phân chia model (volume strategy) + đơn vị, gốc tọa độ, shared coordinates.
8. Quy trình phối hợp: chu kỳ họp coordination, quy trình clash, phần mềm kiểm tra.
9. Chuẩn phần mềm + phiên bản (Revit version thống nhất toàn dự án — không nâng giữa chừng).
10. QA/QC: checklist kiểm model trước mỗi lần Shared (dùng skill `qa-checklist`).

## Khi người dùng nhờ soạn/rà soát BEP
- Hỏi: giai đoạn (pre/post), EIR có chưa, các bộ môn tham gia, CDE dùng gì.
- Rà theo checklist 10 mục trên, đánh dấu mục THIẾU và đề xuất nội dung mẫu cho từng mục thiếu.
