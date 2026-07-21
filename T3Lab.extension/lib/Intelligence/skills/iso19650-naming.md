---
name: iso19650-naming
description: ISO 19650 file naming and document codes for BIM deliverables
triggers: iso 19650, iso19650, dat ten file, ma tai lieu, naming convention, ten ban giao, ma ban ve, suitability code, revision code, document code, quy tac dat ten
agents: general, knowledge
tools:
---
# Đặt tên file & mã tài liệu theo ISO 19650 (T3Lab)

## Cấu trúc mã chuẩn
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
1. Đề xuất bảng: tên cũ → mã mới theo cấu trúc trên, hỏi các trường còn thiếu (mã dự án, khối, bộ môn).
2. Chỉ ra tên không thể suy luận được để người dùng bổ sung — không tự bịa mã.
