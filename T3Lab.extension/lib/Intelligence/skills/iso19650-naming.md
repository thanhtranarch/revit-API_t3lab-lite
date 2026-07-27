---
name: iso19650-naming
description: ISO 19650 file naming and document codes for BIM deliverables
triggers: iso 19650, iso19650, dat ten file, ma tai lieu, naming convention, ten ban giao, ma ban ve, suitability code, revision code, document code, quy tac dat ten
agents: general, knowledge
tools:
standard: project
---
# Đặt tên file & mã tài liệu theo ISO 19650 (T3Lab)

> **⚠️ Nguồn chuẩn — đọc trước khi áp dụng**
> Mọi mã, tiền tố, cách đánh số, tên gọi và ngưỡng viết trong tài liệu này chỉ là **mẫu tham khảo chung của ngành**, KHÔNG phải chuẩn của dự án. Chuẩn thật phụ thuộc từng công ty/chủ đầu tư.
> 1. **Ưu tiên tuyệt đối**: BEP / EIR / tiêu chuẩn nội bộ của dự án có trong knowledge base — áp dụng đúng theo đó và ghi rõ tên file nguồn đã lấy quy tắc.
> 2. **Nếu chưa có file chuẩn**: nói rõ một câu "chưa tìm thấy BEP/tiêu chuẩn dự án trong knowledge base", rồi đề nghị người dùng bổ sung theo một trong các cách: **Settings → Projects → Link folder** (link thẳng thư mục BEP/tiêu chuẩn của dự án — quét tại chỗ, tự sinh file `context/CONTEXT.md`), **Settings → Projects → Add files** (copy file BEP/standard vào project), hoặc **Settings → Knowledge → Add folder** (thư mục tiêu chuẩn dùng chung cho mọi dự án). Sau đó chạy lại yêu cầu.
> 3. **Không tự bịa** mã dự án, tiền tố bộ môn, mã revision/suitability hay yêu cầu LOD. Nếu model hiện tại đã có quy ước sẵn thì được phép theo quy ước quan sát được — nhưng phải nói rõ là suy ra từ model, không phải từ tài liệu chuẩn.

## Cấu trúc mã — MẪU THAM KHẢO (thay bằng cấu trúc trong BEP dự án)
`<DựÁn>-<ĐơnVị>-<KhốiCôngTrình>-<Tầng>-<LoạiTàiLiệu>-<BộMôn>-<SốThứTự>`

Ví dụ: `T3L01-T3L-ZZ-01-DR-A-1001`

| Trường | Ý nghĩa | Ví dụ |
|---|---|---|
| DựÁn | Mã dự án 3–6 ký tự | `T3L01` |
| ĐơnVị | Mã đơn vị tạo tài liệu | `T3L` |
| KhốiCôngTrình | Khối/zone; `ZZ` = toàn bộ | `B1`, `ZZ` |
| Tầng | `01`, `M1` lửng, `RF` mái, `XX` nhiều tầng | `01` |
| LoạiTàiLiệu | `DR` bản vẽ, `M3` model 3D, `SP` đặc tả, `SH` schedule, `RP` báo cáo | `DR` |
| BộMôn | `A` kiến trúc, `S` kết cấu, `M`/`E`/`P` MEP, `Z` liên bộ môn | `A` |
| SốThứTự | 4 chữ số | `1001` |

## Quy tắc kèm theo
- Chỉ dùng chữ HOA, số và dấu gạch nối; KHÔNG dấu cách, không tiếng Việt có dấu trong tên file.
- Trạng thái (suitability code) ghi ở metadata/CDE, không nhét vào tên file: S0 WIP, S1 chia sẻ phối hợp, S3 duyệt, A1 đã phê duyệt.
- Revision: `P01, P02...` (sơ bộ), `C01...` (đã ký hợp đồng/thi công).

## Khi người dùng đưa danh sách file cần đổi tên
1. **Tra BEP/tiêu chuẩn dự án trước** — lấy đúng cấu trúc mã, mã dự án, mã đơn vị, bộ môn từ đó (ghi rõ file nguồn). Chưa có file → báo thiếu và đề nghị upload, đừng dùng bảng mẫu như thể là chuẩn dự án.
2. Đề xuất bảng: tên cũ → mã mới theo cấu trúc ĐÃ XÁC ĐỊNH, hỏi các trường còn thiếu (mã dự án, khối, bộ môn).
3. Chỉ ra tên không thể suy luận được để người dùng bổ sung — không tự bịa mã.
