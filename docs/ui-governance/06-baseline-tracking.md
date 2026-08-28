# 06 · COMPARE & LOG — baseline, trend, regression, change history

---

## 1 · Baseline

Mỗi tool có một dòng baseline trong `BASELINE.md`. Baseline là **trạng thái đã biết**
của tool tại cycle gần nhất — mọi so sánh đều dựa vào nó.

### Schema một dòng

| Cột | Nghĩa |
|-----|-------|
| `Tool` | Tên tool (theo XAML) |
| `File` | Đường dẫn XAML |
| `State` | `T3` · `LEGACY` · `LOCKED` |
| `Pattern` | P1–P5 · MIXED · NONE |
| `Size` | S · M · L · OFF |
| `Prev` | Điểm cycle trước |
| `Score` | Điểm cycle này |
| `Δ` | Chênh lệch |
| `Last audit` | Ngày audit sâu gần nhất |
| `Open` | Số issue đang mở, kèm severity cao nhất (vd `3 (P1)`) |
| `Fixed` | Số issue đã đóng luỹ kế |
| `Compliance` | % tiêu chí Design System đạt |
| `Regression` | `—` · `⚠ −N` |

Tool chưa audit sâu lần nào: `Score = —`, `Last audit = —`. **Không bịa điểm.**

---

## 2 · COMPARE — so với cycle trước

Mỗi cycle, với mỗi tool đã audit:

```
Δ = Score(cycle N) − Score(cycle N−1)
```

| Δ | Ý nghĩa | Hành động |
|---|---------|-----------|
| `> 0` | Cải thiện | Ghi vào *Top Improvements* nếu ≥ +5 |
| `= 0` | Đứng yên | Không cần nói gì, trừ khi đứng yên ≥3 cycle với điểm <70 |
| `< 0` | **REGRESSION** | Bắt buộc điều tra — §3 |

Nếu tiêu chí chấm điểm thay đổi giữa hai cycle (rubric được sửa), phải **chấm lại
cycle trước bằng rubric mới** trước khi so sánh, và ghi rõ trong report:
`Điểm cycle N−1 đã được chuẩn hoá lại theo rubric mới`. Nếu không, mọi Δ đều vô nghĩa.

---

## 3 · REGRESSION DETECTION

Score giảm → **REGRESSION DETECTED**. Không được bỏ qua, không được làm tròn cho qua.

Phải xác định đủ 8 trường:

```
REGRESSION DETECTED
Tool:               ManaViews
Previous score:     84
Current score:      76
Score difference:   −8
Reason:             Commit a1b2c3d thêm cột "Discipline" vào DataGrid với Width="*",
                    làm grid có 2 cột * → vi phạm L3, và bật lại scroll ngang → vi phạm L4
Changed files:      lib/GUI/Tools/ManaViews.xaml (+14 −2)
Possible regression: Cột mới có thể được thêm để phục vụ tính năng đang làm dở —
                    kiểm tra script.py có đọc cột này không trước khi sửa
Recommended action: Đổi Width="*" của "Discipline" thành 140px, thêm lại
                    HorizontalScrollBarVisibility="Disabled"
```

### Ba loại regression

| Loại | Dấu hiệu | Xử lý |
|------|----------|-------|
| **Do người khác** | Commit ngoài routine sửa XAML | Sửa lại theo chuẩn, báo trong report |
| **Do routine** | Cycle trước routine tự sửa và làm hỏng | **Revert ngay**, ghi vào CHANGELOG như một bài học |
| **Do rubric đổi** | Không ai sửa file nhưng điểm tụt | Không phải regression — chuẩn hoá lại điểm cũ (§2) |

---

## 4 · SCORE TREND

Theo dõi mỗi cycle, ghi vào `CHANGELOG.md`:

| Chỉ số | Ghi chú |
|--------|---------|
| Overall / Average score | Kèm mẫu số `n=` |
| Highest score | Tool nào |
| Lowest score | Tool nào |
| Số tool cải thiện / tụt điểm | |
| Open issues theo severity | P0/P1/P2/P3/P4 |
| Fixed this cycle | |
| New issues | |
| Số file còn LEGACY | Chỉ số migration quan trọng nhất |

```
Previous Average: 76.4
Current Average:  81.7
Improvement:      +5.3

P0: 0   P1: 2   P2: 11   P3: 24   P4: 7
Fixed this cycle: 18      New issues: 4
LEGACY còn lại:   49/51   (−2)
```

---

## 5 · CONTINUOUS IMPROVEMENT — tìm pattern giữa các tool

Không chỉ đánh giá từng tool độc lập. Sau khi audit, hỏi:

> Lỗi này có xuất hiện ở tool khác không?

Nếu **≥3 tool** cùng một lỗi → ghi nhận **SYSTEM-WIDE UI INCONSISTENCY**:

```
SYSTEM-WIDE UI INCONSISTENCY
Pattern:        7 tool dùng chiều rộng dialog khác nhau (520/560/600/640/700/800/900)
Affected tools: ManaViews, ManaSheets, ManaFami, TagChecker, ModelAuditor,
                SplitElements, RibbonNames
Root cause:     Không có helper nào áp size class; mỗi tool tự đặt Width
Shared component: chưa có — T3Lab.Styles.xaml không có Style cho <Window>
Design System gap: Standard định nghĩa S/M/L nhưng stylesheet không cung cấp cách áp
Recommendation: Thêm 3 Style cho Window (T3.Window.S/.M/.L) vào T3Lab.Styles.xaml
                → MANUAL REVIEW REQUIRED (sửa Design System)
```

**Một thay đổi tốt ở shared component cải thiện nhiều tool cùng lúc** — luôn ưu tiên
sửa ở tầng shared nếu an toàn, thay vì sửa 7 lần ở 7 file.

---

## 6 · TREND ANALYSIS — câu hỏi bắt buộc mỗi cycle

Trả lời đủ 7 câu, ngắn gọn, đưa vào report:

1. Vấn đề nào lặp lại nhiều nhất?
2. Component nào inconsistent nhất?
3. Tool nào có score thấp nhất?
4. Tool nào cải thiện nhiều nhất?
5. Tool nào giảm điểm?
6. Design System rule nào thường xuyên bị vi phạm?
7. Có UI pattern nào cần được chuẩn hoá thành component dùng chung?

> Tìm **root cause** thay vì sửa từng triệu chứng. Nếu câu 6 luôn ra cùng một rule qua
> nhiều cycle, vấn đề không nằm ở các tool — nó nằm ở chỗ rule đó khó tuân thủ.

---

## 7 · DESIGN SYSTEM MATURITY

Mỗi cycle, đánh giá cả Design System, không chỉ tool:

| Câu hỏi | Nếu "không" |
|---------|-------------|
| Rule có rõ ràng không? | `DESIGN SYSTEM GAP` — mơ hồ ở đâu |
| Component đã được định nghĩa đầy đủ chưa? | Liệt kê component thiếu |
| Có component bị duplicate không? | Đề xuất gộp |
| Có pattern dùng phổ biến nhưng chưa vào Design System? | Đề xuất đưa vào |
| Có rule nào mâu thuẫn nhau? | Nêu cả hai rule |
| Có UI problem lặp lại vì Design System chưa giải quyết? | Đây là gap quan trọng nhất |

**Không tự ý thay đổi Design System.** Ghi nhận gap → đề xuất rule/component → đưa vào
report → chờ user duyệt.

---

## 8 · CHANGE HISTORY

Mỗi cycle ghi một entry vào `CHANGELOG.md`, đủ 11 trường:

```
DATE:                   2026-08-29
CYCLE:                  3
TOOLS AUDITED:          6  (AutoDimension, ManaSheets, TagChecker, PointCloud,
                            SplitElements, RibbonNames)
ISSUES FOUND:           21
ISSUES FIXED:           14
ISSUES REMAINING:       7   (P1×1, P2×4, P3×2)
SCORE CHANGES:          AutoDimension 62→79 (+17) · ManaSheets 71→79 (+8) ·
                        TagChecker 68→74 (+6) · PointCloud 64→64 (0) ·
                        SplitElements 68→72 (+4) · RibbonNames — →70 (mới)
REGRESSIONS:            không
SYSTEM-WIDE PATTERNS:   dialog width không theo size class (7 tool)
DESIGN SYSTEM GAPS:     #1 vị trí deploy T3Lab.Styles.xaml chưa chốt ·
                        #2 thiếu Style cho <Window> theo size class
NEXT PRIORITIES:        P1 viết dev/audit_t3.py · P2 migrate ManaSheets + ManaViews
COMMIT:                 <sha>
```

## 9 · Thứ tự cập nhật file cuối cycle

1. `BASELINE.md` — điểm, trạng thái, issue count của mọi tool đã đụng tới.
2. `CHANGELOG.md` — thêm entry cycle mới **ở trên cùng**.
3. `PRIORITY_QUEUE.md` — xoá việc đã làm, thêm việc mới, sắp lại thứ tự.
4. Commit cả 3 file cùng với thay đổi code của cycle — **một commit, một cycle**.
