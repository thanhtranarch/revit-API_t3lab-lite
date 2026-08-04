# T3Lab Assistant — kiến trúc agent dạng đồ thị + NLU song ngữ

Tài liệu này mô tả hai lớp mới của T3Lab Assistant:

- `T3Lab.extension/lib/Intelligence/language/` — phân tích từ ngữ & ngữ cảnh
  tiếng Việt / tiếng Anh
- `T3Lab.extension/lib/Intelligence/graph/` — hệ thống agent dạng đồ thị

Cả hai đều **thuần Python**, không import Revit API hay WPF, nên chạy được
dưới IronPython 2.7 (trong Revit) lẫn CPython 3 (test trong `dev/`).

---

## 1. Vì sao

Trước đây mỗi lượt chat là một **đường thẳng**:

```
tin nhắn → phân loại 1 lần → 1 specialist → 1 vòng agent loop → 1 câu trả lời
```

Hình dạng đó không diễn đạt được một yêu cầu rất bình thường như
*"thống kê tường rồi xuất pdf sheet A-101"* — hai mục tiêu, có phụ thuộc. Hệ
thống làm mục tiêu đầu rồi bỏ phần còn lại.

Về ngôn ngữ, phân tích từ ngữ chỉ gồm bỏ dấu (`knowledge/vi_text.py`) + bảng
keyword phẳng trong `agents/dispatcher.py`. Ba ca hỏng thường gặp nhất trong
văn phòng BIM Việt Nam:

| Người dùng gõ | Trước đây | Vì sao |
|---|---|---|
| `xuat pdf sheet A` | trả lời **tiếng Anh** | không dấu → `_is_viet_text` đọc thành tiếng Anh |
| `export cái sheet này ra pdf` | tuỳ nửa nào thắng | trộn Anh–Việt |
| `ktra warning ko` | rơi về `general` | teencode, không khớp bảng nào |

---

## 2. Lớp ngôn ngữ — `Intelligence/language/`

Một lần phân tích, mọi nơi dùng chung:

```python
from Intelligence.language import analyze

utt = analyze(u"đừng xóa text note trong view này")
utt.lang              # 'vi'
utt.modality          # 'command'
utt.negation          # {'negated': True, 'prohibitive': True, ...}
utt.scope             # 'view'
utt.risk['level']     # 'destructive'
utt.categories        # ['TextNotes']
utt.is_safe_to_act    # False  ← câu này CẤM hành động
utt.to_prompt_hint()  # khối phân tích ngắn, chèn vào prompt LLM
```

| Module | Trách nhiệm |
|---|---|
| `lang_detect.py` | Chấm điểm bằng chứng VI/EN. Thuật ngữ BIM (`sheet`, `export`, `level`…) tính là **trung lập**. Trả `'unknown'` khi không đủ bằng chứng thay vì đoán bừa. |
| `normalizer.py` | Teencode, viết tắt domain, ký tự lặp, telex thừa dấu (`xoas`→`xoa`), sửa typo 1-edit (Damerau, có bắt hoán vị) với lexicon đóng + guard list. |
| `entities.py` | Từ điển song ngữ → tên Revit: category, màu, tầng, sheet, số đo (quy về mm), số lượng. |
| `semantics.py` | Phủ định, modality, scope, risk, và bảng sequencer để tách nhiều mục tiêu. |
| `analyzer.py` | Gộp thành `Utterance`, sinh `to_prompt_hint()`, và `resolve_ellipsis()` cho câu rút gọn qua lượt. |

### 2.1 Quy tắc dấu tiếng Việt (quan trọng)

Bỏ dấu làm **gộp** các từ khác nghĩa. Mỗi ca dưới đây từng là một false
positive thật:

| Viết đúng | Bỏ dấu | Bị nhầm thành |
|---|---|---|
| `từ` (from) | `tu` | `tủ` → category **Casework** |
| `đến` (to) | `den` | `đèn` → **LightingFixtures** |
| `của` (of) | `cua` | `cửa` → **Doors** |
| `chớ` (don't) | `cho` | `cho` (give) → phủ định giả |
| `dựng` (build) | `dung` | `đừng` (don't) → phủ định giả |
| `chữa` (fix) | `chua` | `chưa` (not yet) → phủ định giả |
| `tìm` (search) | `tim` | `tím` → màu tím |

Vì vậy **alias đơn âm tiếng Việt được khớp trên văn bản CÓ DẤU**; alias nhiều
âm tiết (`cua so`, `mau do`, `cau thang`) vẫn khớp trên văn bản bỏ dấu, nên gõ
không dấu vẫn nhận diện được. Đây đúng là cách bảng `_ACTION_COLOR_RAW` sẵn có
trong `dispatcher.py` đang xử lý `"tô đỏ"`.

**Hệ quả đã biết**: `"cua"` trần (không dấu, không bổ ngữ) không ra category
nào — không phân biệt được với `"của"`. Viết `cua di` / `cua so`, hoặc gõ dấu.

### 2.2 Nguyên tắc an toàn

> Văn bản đã chuẩn hoá **chỉ dùng cho tín hiệu định tuyến**. Thứ gửi cho model,
> lưu vào history và hiển thị trong transcript **luôn là nguyên văn user gõ**.

Nên một lần sửa typo sai chỉ có thể làm lệch route — không bao giờ đặt chữ vào
miệng người dùng.

---

## 3. Lớp đồ thị — `Intelligence/graph/`

Ánh xạ trực tiếp từ sơ đồ Graph Engineering:

```
        USER GOAL
            │
            ▼
   GRAPH ORCHESTRATOR ──────────────┐   orchestrator.py
   plan · dispatch · execute        │   (5 graph operations)
   observe · adapt                  │
            │                       │
            ▼                       │
    GRAPH COMPONENTS                │
    agents · tools · data · goals   │   planner.py · router.py
            │                       │
            ▼                       │
     EXECUTION LAYER                │   executor.py
     song song · context · lỗi      │
            │                       │
            ▼                       │
  OBSERVABILITY & FEEDBACK ─────────┘   observability.py
  spans · traces · verdicts             verifier.py
```

| Module | Primitive / operation trong sơ đồ |
|---|---|
| `primitives.py` | NODE, EDGE, STATE, CONTEXT, MEMORY |
| `topology.py` | 8 topology pattern + `layers()` — bộ lập lịch |
| `router.py` | ROUTER |
| `planner.py` | operation **Plan** |
| `reducer.py` | REDUCER |
| `verifier.py` | VERIFIER |
| `executor.py` | operation **Execute** (EXECUTION LAYER) |
| `observability.py` | operation **Observe** |
| `orchestrator.py` | operation **Dispatch** + **Adapt**, chạy cả vòng |

### 3.1 Planner — bảo thủ có chủ đích

Planner thà trả về **một** node (đúng hành vi cũ) còn hơn tách sai một câu nó
không chắc.

1. **Sequencer mạnh** + dấu câu ngắt mệnh đề (`rồi`, `sau đó`, `then`, `;`,
   danh sách đánh số) → tách.
2. **Sequencer yếu** (`và`, `and`) → chỉ tách khi **cả hai vế** độc lập khớp
   keyword stage của dispatcher.
   - `thống kê tường **và** cột` → **1 mục tiêu** (`cột` không phải câu lệnh)
   - `thống kê tường **và** xuất pdf` → **2 mục tiêu**
3. Mảnh không có động từ riêng được **ghép lại** vào vế trước:
   `thống kê tường rồi cửa sổ` = một động từ, hai đối tượng.
4. Trần `MAX_NODES = 6`; phần dư gộp vào node cuối chứ không bị bỏ.

**Thứ tự**: node **đọc** độc lập → fan-out chạy song song. Node **ghi** phụ
thuộc vào mọi node user nói trước nó, giữ đúng thứ tự đã yêu cầu.

### 3.2 Executor — 4 tính chất

| Tính chất | Nghĩa |
|---|---|
| **Writer serialisation** | Không bao giờ có 2 node sửa model chạy cùng lúc — 2 transaction Revit đan nhau làm hỏng undo stack. |
| **Fault containment** | Node lỗi không làm hỏng cả đồ thị: node phụ thuộc bị `SKIPPED`, nhánh không liên quan chạy tiếp. |
| **Redundant paths** | Node lỗi có cạnh `fallback` sẽ chạy đường dự phòng; nếu thành công, kết quả được ghi **dưới id node gốc** — đồ thị đã phục hồi. |
| **Cooperative cancel** | `state.cancel()` dừng giữa các layer; không giết luồng bằng vũ lực. |

### 3.3 Giới hạn đã biết — `max_parallel=1` khi gắn UI

`script.py` ghim `max_parallel=1`: runner stream thẳng vào transcript chat, hai
lượt agent viết chung một transcript sẽ đan vào nhau thành vô nghĩa.

Đường song song là **thật** và được test (`dev/test_graph_agents.py` kiểm tra 3
node 150 ms xong dưới 500 ms), dành cho caller headless — báo cáo nền, batch —
không viết vào transcript đang sống.

### 3.4 Adapt

Node **đọc** thất bại verification, và đường `fallback` cũng không cứu được, sẽ
được hạ về specialist `general` (đủ tool, không thu hẹp) và chạy lại **đúng một
lần**. Node **ghi** không bao giờ được adapt — chạy lại một lệnh sửa model là
cách để nó được áp dụng hai lần.

---

## 4. Trải nghiệm người dùng

Với `thống kê tường rồi xuất pdf sheet A-101`:

```
Bước 1/2 — thống kê tường
  [câu trả lời của node 1 stream vào chat]
Bước 2/2 — xuất pdf sheet A-101
  [câu trả lời của node 2 stream vào chat]
✔ Đã hoàn thành cả 2 bước.
```

Khi có bước lỗi, dòng cuối liệt kê **đúng những bước không chạy được** (reducer
chỉ được gọi trên các node thất bại — các node thành công đã hiển thị rồi, in
lại là thừa).

---

## 5. Kill switch

| Setting | Mặc định | Tác dụng |
|---|---|---|
| `agents.graph_mode` | `true` | Tắt → lượt nhiều mục tiêu quay về đúng đường single-specialist cũ |
| `agents.multi_agent` | `true` | Tắt → tắt luôn cả graph (graph cần dispatcher) |

File: `%APPDATA%/T3LabAI/settings.json`

```json
{ "agents": { "graph_mode": false } }
```

Lượt **một mục tiêu** không bao giờ đi vào lớp graph, nên tắt/bật không ảnh
hưởng gì tới đại đa số cuộc chat.

---

## 6. Đọc trace

Mỗi lượt graph ghi một dòng JSONL:

```
%APPDATA%/T3LabAI/telemetry/graph/<YYYY-MM-DD>.jsonl
```

```python
from Intelligence.graph import observability
rows = observability.read_traces(path)
rows[0]['pattern']      # 'diamond'
rows[0]['spans']        # [{node, specialist, status, ms, attempts, verdict}, ...]
rows[0]['adaptations']  # [{node, change, detail}, ...]
```

Câu hỏi trace trả lời được mà `telemetry.py` (per-turn) không:

- planner chọn hình dạng nào, và có đúng không?
- node nào chậm — hay chậm ở reducer?
- đường fallback thực sự cứu được node bao nhiêu lần?
- specialist nào hay trượt verification, và trượt trên loại câu hỏi nào?

---

## 7. Chạy test

```bash
python3 dev/test_language.py          # 134 check
python3 dev/test_graph_agents.py      # 104 check
python3 dev/test_assistant_routing.py # regression dispatcher/routing
python3 dev/test_knowledge.py
python3 dev/audit_tools.py --quiet
python3 dev/audit_ui.py --quiet
python3 dev/sync_wpf_styles.py --check
```

Không file XAML nào bị sửa — 3 file UI-frozen (`DWGManagement`,
`ExportManager`, `T3LabAssistant`) không bị đụng tới.

---

## 8. Mở rộng

**Thêm từ vựng** → sửa bảng trong `language/entities.py`. Alias đơn âm tiếng
Việt phải viết **có dấu**; alias nhiều âm tiết viết ASCII (khớp cả không dấu).
Thêm test vào `dev/test_language.py`.

**Thêm specialist** → khai báo trong `agents/specialists.py` như cũ, rồi thêm
vào `graph/router.py`: `WRITER_SPECIALISTS` nếu nó sửa model, và
`_FALLBACK_ROUTE` cho đường dự phòng.

**Dùng graph ngoài chat** (headless, chạy song song thật):

```python
from Intelligence.graph.executor import GraphExecutor
from Intelligence.graph.primitives import GraphState

state = GraphState(goal=u"...", global_context=u"...")
GraphExecutor(my_runner, max_parallel=3).run(my_graph, state)
```

`my_runner(node, context, state)` trả về chuỗi hoặc `NodeResult`.
