# 07 · REPORT — format báo cáo cuối cycle

Sau mỗi lần chạy, xuất report đúng format dưới đây. Không đổi tên mục, không bỏ mục.
Mục nào không có nội dung thì ghi `— không có —`, không xoá tiêu đề.

---

```
====================================
PYREVIT UI/UX GOVERNANCE REPORT
====================================

DATE:
[YYYY-MM-DD]

CYCLE:
[NUMBER]

------------------------------------
SYSTEM HEALTH
------------------------------------

Overall Score:
[XX.X/100]  (n=[số tool tính điểm])

Previous Score:
[XX.X/100]

Change:
[+X.X / −X.X]

Status:
EXCELLENT / GOOD / ACCEPTABLE / NEEDS IMPROVEMENT / POOR

Design System compliance:
[X]/[N] tool đạt chuẩn T3 · [N] tool còn LEGACY · [N] tool LOCKED


------------------------------------
AUDIT SUMMARY
------------------------------------

Tools Audited:
[X]

Tools Improved:
[X]

Tools Regressed:
[X]

New Issues:
[X]

Fixed Issues:
[X]

Open Issues:
[X]


------------------------------------
ISSUE BREAKDOWN
------------------------------------

P0:  [X]
P1:  [X]
P2:  [X]
P3:  [X]
P4:  [X]


------------------------------------
TOP IMPROVEMENTS
------------------------------------

1.
Tool:
Change:
Impact:
Score change:

2.
Tool:
Change:
Impact:
Score change:

3.
Tool:
Change:
Impact:
Score change:


------------------------------------
LOWEST SCORING TOOLS
------------------------------------

Tool:
Score:
Main problems:
Priority:


------------------------------------
BIGGEST REGRESSIONS
------------------------------------

Tool:
Previous:
Current:
Change:
Cause:


------------------------------------
SYSTEM-WIDE PATTERNS
------------------------------------

Pattern:
Affected tools:
Root cause:
Recommendation:


------------------------------------
DESIGN SYSTEM GAPS
------------------------------------

Gap:
Affected components:
Recommendation:


------------------------------------
NEEDS VERIFICATION — CẦN TEST TRONG REVIT
------------------------------------

[ ] Tool — thao tác cụ thể — kết quả mong đợi
[ ] ...


------------------------------------
NEXT PRIORITIES
------------------------------------

P0/P1:
...

P2:
...

P3:
...


------------------------------------
OVERALL ASSESSMENT
------------------------------------

[Nhận định ngắn, chuyên nghiệp: hệ thống đang đi lên hay đi ngang, nút thắt lớn nhất
là gì, và một việc cụ thể nếu chỉ được làm một việc ở cycle sau]

====================================
```

---

## Luật viết report

1. **Số liệu phải khớp file.** `Overall Score` trong report = trung bình cột `Score`
   trong `BASELINE.md`. Nếu lệch, report sai.
2. **Không báo `FIXED` cho việc chưa verify.** Chúng thuộc mục *NEEDS VERIFICATION*.
3. **Không bịa kết quả cần Revit.** Nếu chưa ai mở Revit, mục đó là checklist chưa tick.
4. **Nêu cả tin xấu.** Regression, gate đỏ, việc bị bỏ dở, ngân sách hết — nói thẳng.
5. **Gate đỏ sẵn từ trước** phải ghi rõ là pre-existing, kèm cycle phát hiện lần đầu.
6. **Mục cuối phải chọn được một việc.** "Nếu cycle sau chỉ làm một việc, làm việc này."
7. Kết report bằng danh sách việc còn dang dở, để cycle sau vào là chạy được ngay.
