# 04 · IDENTIFY & PRIORITIZE — issue, severity, hàng đợi

---

## 1 · Thang severity

| Mức | Nghĩa | Ví dụ trong repo này |
|-----|-------|----------------------|
| **P0 — CRITICAL** | Tool không dùng được, hoặc UI gây lỗi nghiêm trọng | `<Grid.RowDefinition/>` → crash `EMPTYPROPERTYELEMENT`; `DataGrid` bọc trong `ScrollViewer` → realize toàn bộ row → treo Revit; `x:Name` bị đổi làm script.py `AttributeError` |
| **P1 — HIGH** | Vấn đề UX nghiêm trọng hoặc vi phạm Design System lớn | Thao tác xoá không có confirm P5; nút phá huỷ là `IsDefault`; không có progress cho tác vụ 30 giây; lỗi ra stacktrace |
| **P2 — MEDIUM** | Inconsistency hoặc usability problem đáng kể | Hai primary trong một footer; thiếu empty state; hex cứng thay cho `T3.*`; font ngoài Segoe UI |
| **P3 — LOW** | Minor visual / cosmetic | Margin 10 thay vì 12; bo góc 6 thay vì 4; label thiếu tooltip |
| **P4 — OPTIONAL** | Nice-to-have | Thêm access key; gộp hai section cho gọn |

**Thứ tự sửa: P0 → P1 → P2 → P3 → P4.** Không bao giờ sửa P3 khi còn P0/P1 chưa xử lý.

### Quy tắc gán severity

- Ảnh hưởng **chạy được / không chạy được** → P0, không cần bàn.
- Ảnh hưởng **người dùng mất dữ liệu hoặc làm sai model** → P1.
- Ảnh hưởng **hiểu sai / chậm / khó chịu** → P2.
- Chỉ ảnh hưởng **cái nhìn** → P3.
- Không ai phàn nàn nếu không sửa → P4.

Khi lưỡng lự giữa hai mức → chọn mức **cao hơn** nếu liên quan đến dữ liệu model,
chọn mức **thấp hơn** nếu chỉ liên quan đến thẩm mỹ.

---

## 2 · Format ghi issue — bắt buộc đủ 9 trường

```
ISSUE-<cycle>-<số>
Tool:               ManaSheets
File:               T3Lab.extension/lib/GUI/Tools/ManaSheets.xaml:142
UI component:       Footer button row
Problem:            Có 2 button dùng T3.Button.Primary ("Create" và "Create & Open")
Why it is a problem: Người dùng không xác định được đâu là hành động chính; hai nút
                    cùng trọng lượng thị giác làm chậm quyết định và tăng nguy cơ bấm
                    nhầm nút tạo-và-mở trên model lớn
Design System rule:  T3LAB_UI_STANDARD.md §Luật bố cục 6 — "phải MỘT primary, ngoài
                    cùng phải"
Severity:           P2
Estimated impact:   Visual Consistency −1, Interaction −4  ⇒ tổng −5
Recommended solution: Đổi "Create & Open" sang T3.Button.Secondary, giữ "Create" làm
                    primary ngoài cùng phải
Status:             FIXED | NEEDS VERIFICATION | NEEDS REVIEW | MANUAL REVIEW REQUIRED | OPEN
```

### Cấm mô tả mơ hồ

❌ "UI chưa đẹp"
❌ "Layout hơi lộn xộn"
❌ "Nên cải thiện trải nghiệm"

✅ "Primary action button đang có cùng visual hierarchy với secondary action, làm người
dùng khó xác định action chính."

Một issue phải trả lời được: *ai gặp vấn đề gì, khi nào, và luật nào bị vi phạm.*

---

## 3 · Trạng thái issue

| Status | Nghĩa | Điều kiện |
|--------|-------|-----------|
| `OPEN` | Đã ghi nhận, chưa xử lý | — |
| `FIXED` | Đã sửa **và đã xác minh** | Qua đủ 7 bước verify ở `05` |
| `NEEDS VERIFICATION` | Đã sửa code, chưa xác minh được | Cần mở Revit / cần user xác nhận |
| `NEEDS REVIEW` | Thiếu context để quyết định | Ghi rõ thiếu gì |
| `MANUAL REVIEW REQUIRED` | Sửa được nhưng chạm business logic hoặc Design System | Kèm đề xuất cụ thể |
| `WONTFIX` | User quyết định không sửa | Kèm ngày + lý do |

> `FIXED` **không** được gán chỉ vì code đã đổi. Xem `05-improvement-and-safety.md` §3.

---

## 4 · PRIORITIZE — hàng đợi cho cycle sau

Sau khi audit xong, xếp mọi issue còn `OPEN` vào `PRIORITY_QUEUE.md` theo đúng thứ tự:

1. **Critical bugs** (P0) — bất kể ở tool nào.
2. **Severe UX problems** (P1).
3. **Design System violations** (P2 nhóm A/B).
4. **Repeated system-wide inconsistencies** — cùng một lỗi ở ≥3 tool; sửa ở tầng shared trước.
5. **Low-score tools** — điểm < 70.
6. **Regression** — tool bị tụt điểm so với cycle trước.
7. **Minor improvements** (P3/P4).

Trong cùng một bậc, ưu tiên:
- Issue sửa được ở **shared component** (một sửa, nhiều tool được lợi) trước issue một-file.
- Tool được dùng nhiều trước tool ít dùng.
- Issue đã tồn tại qua ≥3 cycle trước issue mới (chống tồn đọng vĩnh viễn).

## 5 · Chống tồn đọng

Nếu một issue `OPEN` sống qua **5 cycle** mà không được đụng tới: đẩy nó lên đầu report
mục *Next Priorities* với nhãn `STALE`, và nêu rõ vì sao nó liên tục bị bỏ qua (thiếu
context? cần user duyệt? hay chỉ vì luôn có việc gấp hơn?). Nếu lý do là "cần user
duyệt" → hỏi user trong report.
