# 03 · SCORE — hệ thống chấm điểm UI/UX

Mỗi tool được chấm **0–100**. Điểm phải **tái lập được**: hai lần chấm cùng một file
không đổi phải ra cùng một số. Vì vậy mỗi category được chia thành tiêu chí con có
điểm cụ thể — không chấm theo cảm giác.

---

## 1 · Thang điểm

| Category | Điểm | Nguồn kiểm tra |
|----------|------|----------------|
| Visual Consistency | **20** | `02-audit-checklist.md` §A |
| Component Consistency | **15** | §B |
| Usability | **20** | §C |
| Interaction | **15** | §D |
| Content & Clarity | **10** | §E |
| Accessibility | **10** | §F |
| Feedback & Error UX | **5** | §C4/C9–C11 + §G |
| Maintainability | **5** | §B + cấu trúc file |
| **TOTAL** | **100** | |

## 2 · Mức đánh giá

| Điểm | Mức |
|------|-----|
| 90–100 | **EXCELLENT** |
| 80–89 | **GOOD** |
| 70–79 | **ACCEPTABLE** |
| 60–69 | **NEEDS IMPROVEMENT** |
| < 60 | **POOR** |

---

## 3 · Rubric chi tiết

### 3.1 · Visual Consistency — 20

| Tiêu chí | Điểm | Trừ hết điểm khi |
|----------|------|------------------|
| Đúng một pattern P1–P5 | 4 | Pattern `MIXED` hoặc `NONE` |
| Font Segoe UI / Consolas, không font khác | 3 | Còn Hanken Grotesk / Inter / Manrope |
| Đúng thang 7 size | 3 | Có ≥1 size ngoài thang |
| Màu 100% từ `T3.*`, không hex cứng | 4 | Có ≥1 hex cứng (trừ `#F59E0B` copyright nếu còn) |
| Spacing chia hết 4, chỉ 4/8/12/16/24/32 | 3 | Có ≥3 margin lệch |
| Chiều cao & bo góc đúng bảng | 2 | Có ≥2 giá trị lệch |
| Size class S/M/L + Min* | 1 | Không có `MinWidth`/`MinHeight` |

Trừ theo tỉ lệ: `điểm = điểm_tối_đa × (1 − số_vi_phạm / ngưỡng_trừ_hết)`, làm tròn xuống.

### 3.2 · Component Consistency — 15

| Tiêu chí | Điểm |
|----------|------|
| Merge `T3Lab.Styles.xaml`, không `<Style>` tự chế trong file tool | 5 |
| Button dùng `T3.Button.*` | 3 |
| Input (TextBox/ComboBox/CheckBox/Radio) dùng `T3.*` | 3 |
| List/DataGrid dùng `T3.ListBox` / `T3.DataGrid*` | 2 |
| Progress / Callout / StatusPill dùng `T3.*` | 2 |

### 3.3 · Usability — 20

| Tiêu chí | Điểm |
|----------|------|
| Information hierarchy — 3 giây hiểu cửa sổ làm gì (C1–C3) | 4 |
| Input validation rõ ràng, chặn trước khi chạy (C7) | 3 |
| Error prevention — thao tác phá huỷ khó bấm nhầm (C5) | 3 |
| Error recovery — hỏng rồi làm lại được (C6) | 2 |
| Loading / progress cho tác vụ >2s (C8, C9) | 3 |
| Empty state ở mọi list/grid (L10) | 3 |
| Confirmation flow đúng P5 cho thao tác phá huỷ (C12) | 2 |

### 3.4 · Interaction — 15

| Tiêu chí | Điểm |
|----------|------|
| Footer đúng thứ tự L6, đúng **một** primary ngoài cùng phải | 4 |
| Có `IsDefault` **và** `IsCancel` (D3, D4) | 3 |
| Primary vs secondary phân biệt được bằng hình thức | 2 |
| Tab order theo thứ tự đọc | 2 |
| Selection behavior (All/None, Shift/Ctrl) | 2 |
| Nhất quán với các tool cùng loại | 2 |

### 3.5 · Content & Clarity — 10

| Tiêu chí | Điểm |
|----------|------|
| Tên ribbon = tên title bar; label không viết tắt tự chế | 2 |
| Một ngôn ngữ, một giọng văn trong cùng tool | 2 |
| Error message nói *cái gì sai · ở đâu · làm gì tiếp* | 3 |
| Warning / success có **số lượng** cụ thể | 2 |
| Thuật ngữ Revit dùng đúng | 1 |

### 3.6 · Accessibility — 10

| Tiêu chí | Điểm |
|----------|------|
| Không chữ nhỏ hơn thang cho nội dung chính | 2 |
| Contrast ≥4.5:1 | 2 |
| Phân cấp bằng size/weight, không chỉ bằng màu | 2 |
| Clickable area đủ lớn (≥28 cao) | 1 |
| Hover/press/focus/disabled đều nhìn thấy được | 2 |
| Panel lồng ≤2 cấp | 1 |

### 3.7 · Feedback & Error UX — 5

| Tiêu chí | Điểm |
|----------|------|
| Progress có phase + `n / total` + item hiện tại | 2 |
| Status **không bao giờ chỉ bằng màu** — luôn kèm chữ | 1 |
| Lỗi ra câu tiếng người, không stacktrace | 1 |
| Xong việc báo kết quả kèm số lượng | 1 |

### 3.8 · Maintainability — 5

| Tiêu chí | Điểm |
|----------|------|
| Không brush/style/màu định nghĩa cục bộ | 2 |
| XAML không lặp khối lớn có thể tách | 1 |
| `x:Name` đặt tên rõ, khớp với script.py | 1 |
| Không control chết / không dùng tới | 1 |

---

## 4 · Trần điểm cứng (hard cap)

Bất kể rubric cộng ra bao nhiêu, các trường hợp sau **chặn trần**:

| Điều kiện | Trần |
|-----------|------|
| Còn ≥1 issue **P0** đang mở | **59** (POOR) |
| Chưa migrate sang T3 (còn hex cứng / font cũ) | **79** (ACCEPTABLE) |
| Pattern là `MIXED` hoặc `NONE` | **84** |
| Chưa từng được QA trong Revit sau lần migrate gần nhất | **89** |

Trần điểm là cách để điểm số không "đẹp giả" khi tool vẫn còn nợ nền tảng.

## 5 · Overall Score của hệ thống

```
Overall = trung bình cộng điểm của mọi tool KHÔNG bị LOCKED
```

File `LOCKED` (§7 của `08-design-system-authority.md`) không có điểm và không tham gia
trung bình. Báo cáo phải ghi rõ mẫu số, ví dụ `Overall 74.2 (n=51)`.

---

## 6 · Cách trình bày điểm — bắt buộc đủ 8 dòng

Không được chỉ đưa tổng điểm. Mỗi tool audit sâu phải có khối:

```
Tool: Wall Numbering

Previous Score: 72
Current Score:  84
Improvement:    +12

Visual Consistency:      16/20
Component Consistency:   13/15
Usability:               17/20
Interaction:             12/15
Content:                  9/10
Accessibility:            8/10
Feedback:                 5/5
Maintainability:          4/5

Đã cải thiện:  spacing về bội số 4 (12 chỗ) · thêm empty state · gộp 2 primary còn 1
Còn tồn tại:   chưa migrate T3 (hex cứng ×23) → trần 79 ⇒ điểm bị cắt còn 79
Hard cap áp dụng: LEGACY (79)
```

Nếu trần điểm cắt điểm, phải nói rõ như dòng cuối trên — nếu không, lần sau không ai
hiểu vì sao rubric ra 84 mà bảng ghi 79.
