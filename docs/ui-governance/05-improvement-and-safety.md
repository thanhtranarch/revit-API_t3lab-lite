# 05 · IMPROVE & VERIFY — luật sửa, luật an toàn, luật xác minh

---

## 1 · Nguyên tắc cải thiện

Khi sửa được **an toàn**, agent phải **chủ động sửa** — không hỏi lại cho từng việc nhỏ.

| Nguyên tắc | Nghĩa cụ thể |
|------------|--------------|
| Follow Design System | Mọi giá trị lấy từ `T3Lab.Styles.xaml`, không tự nghĩ |
| Reuse existing components | Có `T3.Button.Ghost` rồi thì không tạo style nút mờ mới |
| Reuse existing helper | Có helper message/progress trong `lib/Snippets/` thì dùng lại |
| Prefer minimal changes | Sửa đúng dòng gây lỗi, không format lại cả file |
| Preserve current behavior | Sau khi sửa, tool làm đúng việc cũ, ra đúng kết quả cũ |
| Preserve business logic | Không đụng `script.py` trừ khi lỗi nằm ở UI-layer của nó |
| Preserve Revit API logic | Collector/transaction/filter y nguyên |
| Preserve user workflow | Không đổi số bước, không đổi thứ tự thao tác |

**Không redesign một tool chỉ vì vấn đề cosmetic.**

### Thứ tự ưu tiên khi cải thiện

```
Consistency → Usability → Clarity → Feedback → Accessibility → Cosmetic refinement
```

---

## 2 · pyRevit / Revit safety rules

### 2.1 · Tuyệt đối không thay đổi

- Business logic
- Revit API behavior
- Element filtering
- Parameter mapping
- Selection logic
- Category logic
- Document modification logic
- Transaction behavior & transaction boundaries
- Element creation / deletion
- Existing output behavior

### 2.2 · Ranh giới UI ↔ logic trong repo này

| Được sửa | Không được sửa |
|----------|----------------|
| `.xaml` — layout, style, text, thứ tự control | Bất kỳ `FilteredElementCollector` nào |
| Handler thuần UI (`minimize_button_clicked`, đổi tab, toggle panel) | `with Transaction(doc, ...)` — nội dung và biên |
| Chuỗi thông báo hiển thị cho người dùng | Điều kiện `if` quyết định element nào được xử lý |
| Thêm `x:Name` **mới** | Đổi / xoá `x:Name` **đang có** |
| Thêm empty state, progress text | Giá trị parameter được ghi vào model |

### 2.3 · Các lệnh cấm khác

- **Không refactor lớn** nếu không cần cho UI.
- **Không xoá code chỉ vì trông như không được dùng** — pyRevit nạp động, `x:Name` được
  tra bằng chuỗi; "không thấy reference" ≠ "không dùng".
- **Không thêm dependency mới.** IronPython 2.7 + stock WPF, hết.
- **Không đổi behavior để lấy điểm cao hơn.** Điểm là thước đo, không phải mục tiêu.

```
Stability  > Refactoring
Correctness > Cosmetic improvement
Existing Design System > Personal preference
```

---

## 3 · VERIFY — 7 bước sau mỗi thay đổi

Chạy đủ 7 bước cho **mỗi file** đã sửa:

| # | Bước | Cách làm |
|---|------|----------|
| 1 | **Syntax** | Parse XAML: `python3 -c "import xml.dom.minidom as m; m.parse('<file>')"` |
| 2 | **Imports** | `python3 -m py_compile` cho script.py liên quan (nếu có sửa) |
| 3 | **References** | Mọi `x:Name` script.py dùng vẫn còn: `grep -o "self\.[a-z_]*" script.py` đối chiếu XAML |
| 4 | **Resource** | Mọi `{StaticResource T3.*}` có key thật trong `T3Lab.Styles.xaml` |
| 5 | **Logic liên quan** | `git diff` — xác nhận không có dòng nào chạm §2.1 |
| 6 | **UI consistency** | Chấm lại điểm theo `03-scoring.md`, điểm phải tăng hoặc giữ |
| 7 | **Regression risk** | `python3 dev/audit_tools.py --quiet` phải xanh |

### Script kiểm tra nhanh

```bash
# 1 · Syntax mọi XAML đã sửa
for f in $(git diff --name-only -- '*.xaml'); do
  python3 -c "import xml.dom.minidom,sys; xml.dom.minidom.parse('$f')" \
    && echo "OK   $f" || echo "FAIL $f"
done

# 3 · x:Name trước/sau
git show HEAD:<file>.xaml | grep -o 'x:Name="[^"]*"' | sort > /tmp/before.txt
grep -o 'x:Name="[^"]*"' <file>.xaml | sort > /tmp/after.txt
diff /tmp/before.txt /tmp/after.txt    # phải rỗng, hoặc chỉ có dòng THÊM

# 4 · StaticResource mồ côi
grep -oh '{StaticResource \(T3\.[^}]*\)}' <file>.xaml | sed 's/.*T3/T3/;s/}//' | sort -u \
  | while read k; do grep -q "x:Key=\"$k\"" "pyRevit UI Design System/T3Lab.Styles.xaml" \
    || echo "MISSING KEY: $k"; done

# 7 · Gate code tĩnh
python3 dev/audit_tools.py --quiet
```

> Bước 3 là bước quan trọng nhất khi migration. Một `x:Name` bị đổi = `AttributeError`
> lúc runtime = tool chết = P0 do chính routine gây ra.

### Khi nào được đánh `FIXED`

Chỉ khi **đã xác nhận vấn đề được giải quyết**, không phải khi code đã được chỉnh sửa.

| Loại thay đổi | Điều kiện `FIXED` |
|---------------|-------------------|
| Sửa giá trị tĩnh (margin, màu, size, attribute) | Đủ 7 bước verify → `FIXED` |
| Thêm empty state / đổi text | Đủ 7 bước → `FIXED` |
| Gỡ `ScrollViewer` / bật virtualization | Cần Revit → `NEEDS VERIFICATION` |
| Full migration một file | Cần Revit → `NEEDS VERIFICATION` |
| Bất kỳ thứ gì đụng hành vi runtime | Cần Revit → `NEEDS VERIFICATION` |

### Checklist Revit cho user

Mọi mục `NEEDS VERIFICATION` gom thành một checklist ở cuối report, dạng:

```
CẦN TEST TRONG REVIT (cycle N)
[ ] ManaSheets — mở tool, chọn 200 sheet: list cuộn mượt, không đơ >1s
[ ] ManaSheets — bấm Create: kết quả giống trước khi sửa (số sheet tạo ra)
[ ] AutoDimension — cửa sổ hiện đúng size M, không cắt chữ ở 125% scaling
```

**Không bao giờ tự tick các mục này.** Tick khi user báo lại kết quả.

---

## 4 · Decision rules

| Tình huống | Hành động |
|------------|-----------|
| Không chắc chắn | **Do not guess.** → `NEEDS REVIEW` + ghi rõ thiếu gì |
| Thiếu context | **DO NOT MAKE UNSAFE CHANGES** |
| Sửa được nhưng có nguy cơ ảnh hưởng business logic | **DO NOT AUTO-FIX** → đề xuất + `MANUAL REVIEW REQUIRED` |
| Design System không nói gì về case này | `DESIGN SYSTEM GAP` — ghi nhận + đề xuất rule, **không tự thêm luật** |
| Cần thêm style mới vào `T3Lab.Styles.xaml` | `MANUAL REVIEW REQUIRED` — sửa stylesheet là sửa Design System |
| Verify thất bại sau khi sửa | **Revert thay đổi**, ghi issue lại thành `NEEDS REVIEW` |

## 5 · Golden rules

```
Consistency            >  Cosmetic improvement
Usability              >  Decoration
Clarity                >  Complexity
Stability              >  Refactoring
Existing Design System >  Personal preference
Minimal change         >  Large redesign
Root cause             >  Temporary workaround
Reusable solution      >  One-off solution
Verified fix           >  Assumed fix
Continuous improvement >  One-time redesign
```
