# T3Lab Assistant — knowledge, skills, trí thông minh & tốc độ

> Cập nhật 2026-08-02 (mục 7, 8). Các file liên quan:
> `lib/Intelligence/telemetry.py`, `conversation.py`, `feedback.py`,
> `knowledge/query_builder.py`, `knowledge/rerank.py`,
> `agent_loop.py`, `agents/dispatcher.py`, `skills_engine.py`,
> `nlu_engine.py`, `core/server.py`, `Services/revit_context.py`,
> `Intelligence/skill_installer.py`,
> `T3Lab.tab/Support.panel/T3LabAssistant.pushbutton/script.py`,
> `dev/test_assistant_perf.py`, `dev/test_assistant_routing.py`,
> `dev/test_skill_installer.py`, `dev/bench_assistant.py`.

Đợt này **không thêm tính năng mới**. Cả 4 trục đều đã có sẵn cơ chế — vấn đề
nằm ở chỗ chúng bị **nối sai**, và mỗi mục dưới đây là một khiếm khuyết đã xác
minh trong mã nguồn, kèm cách sửa.

---

## 1. Đo lường trước đã

Trước đó **không có chỗ nào** đọc token usage hay bấm giờ một lượt, nên không
thể biết một thay đổi về tốc độ có tác dụng hay không.

`Intelligence/telemetry.py` ghi mỗi lượt một dòng JSONL vào
`%APPDATA%/T3LabAI/telemetry/<ngày>.jsonl`: time-to-first-token, tổng thời
gian, số vòng lặp, **số tool call so với số lần sang Revit**, và bộ đếm token
(`cache_read` / `cache_creation` / `input` / `output`).

Điểm mấu chốt: handler SSE của Claude bỏ qua sự kiện `message_start` — sự kiện
**duy nhất** mang `cache_read_input_tokens`. Đã bắt lại.

Đọc số liệu:

```bash
python3 dev/bench_assistant.py --days 7 --by-provider
```

Bộ đếm bắt đầu/kết thúc ở `_set_busy()` — nút giao **duy nhất** mà mọi nhánh
định tuyến đều đi qua. Việc ghi file được đẩy sang pool thread như `_log_activity`.

> Lưu ý: `_begin_turn_timer` cố tình **không** gọi `has_local_llm()` — hàm đó
> probe qua HTTP, tức là sẽ tự cộng đúng loại độ trễ mà nó đang đo.

---

## 2. Tốc độ

### 2.1 Prompt cache gần như không bao giờ trúng

Ngữ cảnh Revit (view đang mở, số element đang chọn — timer 2 giây làm mới) được
nội suy **vào giữa** system prompt. Khối system là một breakpoint cache; khối
không bao giờ lặp lại ⇒ **cache không trúng giữa các lượt**. Và vì `messages`
nằm **sau** `system` trong prefix cache, transcript cũng trượt theo. Chỉ còn
`tools` được cache.

Đã tách theo đúng đường ranh giới thật:

| | Nội dung | Đi đâu |
|---|---|---|
| **TĨNH** | base prompt, vai specialist, few-shot, project instructions, skills, persistent memory, tóm tắt lăn | `system` — giống hệt nhau qua từng lượt |
| **BIẾN ĐỘNG** | ngữ cảnh Revit, trích đoạn tài liệu theo câu hỏi | đi kèm **lượt user hiện tại**, có rào `<<<T3LAB_LIVE_CONTEXT` |

Lịch sử chỉ lưu text thô nên khối biến động **tự nhiên là ephemeral** — không rò
vào lịch sử, không phình transcript. System prompt vẫn là **chuỗi phẳng**, nên
4 provider ngoài Claude không đổi gì.

`build_agent_system_prompt()` và `build_specialist_prompt()` vẫn nhận
`revit_context` ở vị trí cũ nhưng **bỏ qua** nó — caller cũ không vỡ.

### 2.2 Mỗi tool call là một lần sang Revit

Mỗi call: xếp task → `ExternalEvent.Raise()` → **chặn chờ** Revit rảnh (read
chờ 2 giây rồi mới chạy trên worker). Prompt lại yêu cầu model gộp nhiều read
vào một lượt ⇒ 3 read = 3 lần chờ.

Nút thắt **không** ở handler: `Execute` vốn đã vét sạch hàng đợi trong một lần
gọi (ExternalEvent gộp nhiều `Raise()`). Nút thắt ở `_execute_tool`, nó chặn
chờ chính task của mình trước khi người gọi kịp xếp task tiếp theo.

`server.execute_tools_batch()` xếp cả cụm rồi `Raise()` **một lần**. Chỉ nhận
call **chỉ-đọc**; `agent_loop.leading_read_run()` dừng ở call ghi đầu tiên,
launcher, memory tool, hoặc call trùng — đó chính là thứ giữ cho một read không
bị nhấc lên trước một write làm đổi kết quả của nó. Vòng lặp tuần tự giữ nguyên
(vẫn hiện **từng** thẻ tool, vẫn đủ mọi guard), chỉ là dùng kết quả đã lấy sẵn.

Cả hai seam là tham số **tuỳ chọn**: không truyền thì chạy y như cũ.

### 2.3 Round-trip phân loại nằm chắn trước lượt thật

`dispatcher._llm_stage` gọi `provider.chat()` **không có timeout**, thừa hưởng
trần của từng provider: 60s (cloud) và **180s** (Ollama / LM Studio). Một model
local chậm có thể chắn 3 phút trước khi lượt thật bắt đầu.

- Cả 5 provider giờ đọc `timeout_ms` từ kwargs (mặc định đúng giá trị cũ);
  lời gọi phân loại truyền **6 giây**.
- Khi **một skill trúng dứt khoát** (điểm cao, không có ứng viên bám sát),
  dispatcher định tuyến luôn theo frontmatter của skill và **bỏ hẳn** lời gọi.
  Trúng yếu hoặc lưỡng lự thì vẫn hỏi model.

---

## 3. Skills — hết chọn theo bảng chữ cái

`match()` duyệt `sorted(self._skills)` và caller lấy phần tử 0 ⇒ **chữ cái đầu**
quyết định skill nào tới được model. "Đặt tên sheet theo chuẩn ISO 19650" khớp
cả `iso19650-naming` lẫn `sheet-naming-standard`, và model **luôn** nhận cái đầu
— chỉ vì "i" đứng trước "s".

`match_scored()` chấm theo bằng chứng:

| Tín hiệu | Vì sao |
|---|---|
| số **từ** của trigger | cụm dài là bằng chứng mạnh hơn từ đơn |
| độ phủ câu truy vấn | trigger phủ nhiều hơn thì sát chủ đề hơn |
| số trigger phân biệt trúng | nhiều trigger cùng trúng = củng cố |
| trigger **khai báo** > trigger suy ra | `derive_triggers()` đoán từ văn xuôi, cố ý lỏng |
| project > user > builtin | playbook riêng của dự án phải thắng cái mặc định nó ghi đè |

`build_skills_block` vốn đã inject tối đa **2** body nhưng chỉ từng được đưa **1**
id. Nay cả đường native lẫn đường JSON-intent cũ đều nhận **top-2 đã xếp hạng**
(skill ép bằng `/slash` vẫn chạy một mình).

---

## 4. Knowledge

### 4.1 Câu hỏi nối tiếp truy hồi hụt

Truy hồi chạy trên **câu thô**. "chiều cao lan can theo tiêu chuẩn?" tìm đúng
tài liệu; "còn cầu thang thì sao?" **không mang** chữ nào đã tìm ra tài liệu đó
⇒ BM25 chấm 2 âm tiết với cả kho và trả về nhiễu.

`knowledge/query_builder.py` mở rộng **chỉ** những câu đọc ra là nối tiếp —
câu tự đủ nghĩa đã là truy vấn tốt nhất cho chính nó, nhồi thêm lịch sử chỉ làm
giảm độ chính xác. Phép thử "có phải nối tiếp không" **dùng lại**
`nlu_engine._is_pronoun_query` (vốn đã phân biệt "cái này là gì" với "this
project needs cleanup") thay vì viết một logic thứ hai rồi mâu thuẫn nhau.

Chữ của user **luôn đứng trước và không bao giờ bị bỏ**, nên mở rộng chỉ có thể
tăng recall. Term lấy theo **thứ tự xuất hiện**, ưu tiên câu hỏi trước đó của
chính user, loại bỏ khung câu hỏi ("bao nhiêu", "what about").

### 4.2 Ba trích đoạn cùng một trang

`search()` cắt `top_k` thẳng từ bảng xếp hạng, mà xếp hạng là **độc lập theo
chunk** ⇒ vài chunk mạnh nhất thường là các chunk liền nhau của cùng một trang.
Với `top_k` chỉ 2–3, model thấy một sự kiện ba lần và **không thấy phần còn lại**.

`knowledge/rerank.py` giới hạn phần ngân sách mà một tài liệu được chiếm và bỏ
ứng viên có nội dung đã được phủ bởi cái đã chọn. Giới hạn là **ưu tiên chứ
không phải luật cứng**: nếu không đủ lấp `top_k` thì các ứng viên bị loại tốt
nhất được đưa lại — trả ít trích đoạn còn tệ hơn trả trích đoạn hơi lặp.

Cả hai đều thuần từ vựng, thuần Python ⇒ **không cộng thêm độ trễ** và đường
tới hạn vẫn là BM25.

---

## 5. Trí thông minh

### 5.1 Hội thoại dài quên mất phần đầu

Cửa sổ hội thoại được viết **hai lần** và lệch nhau: `_add_to_history` cắt còn
**16**, `_sanitize_history` lấy **24**. Caller cắt trước ⇒ con số 24 **không
bao giờ có hiệu lực**, và docstring nói nó "cho agent tính liên tục đa lượt
mạnh hơn" mô tả một thứ không thể xảy ra.

`Intelligence.conversation.HISTORY_LIMIT` là nguồn duy nhất; cả hai cùng đọc.

Lượt rơi khỏi cửa sổ giờ được **gộp vào tóm tắt lăn** thay vì vứt đi: giữ mọi
yêu cầu của user, và những câu trả lời có **kết quả cụ thể** (số liệu, id, tên).
Câu tường thuật kiểu "Đang kiểm tra…" bị loại.

Tóm tắt là **trích rút, tất định** — không gọi model ⇒ không tốn độ trễ, chạy
được offline, và **không thể bịa** ra một sự kiện vào lịch sử. Khi vượt ngân
sách thì bỏ dòng **cũ nhất**: quên phần đầu phiên dài vẫn hơn quên việc vừa xong.

Tóm tắt nằm ở **cuối phần TĨNH** của system prompt (chỉ đổi vài lượt một lần
⇒ cache vẫn trúng phần lớn thời gian) và được **gắn nhãn là bối cảnh, không
phải việc cần làm lại**.

Lưu vào khoá `summary` cạnh `messages`; file cũ đọc bằng `.get()` nên **tương
thích ngược**.

### 5.2 👍/👎 không dẫn tới đâu cả

Nút vote có từ lâu, mỗi vote ghi **một dòng** vào activity log rồi thôi. Cơ chế
để hành động cũng đã có (`learn_pattern` / `find_learned_match`) nhưng **chưa ai
nối hai đầu**, và **không có cách nào** ghi nhận một tuyến định tuyến là SAI.
Một mapping sai, một khi đã học, lặp lại mãi mãi.

`Intelligence/feedback.py` giữ danh sách tuyến bị người dùng bác, khoá được
chuẩn hoá qua chính `_normalize_key` của learned-pattern store (nên **không phụ
thuộc thứ tự từ** và khớp đúng entry cần chặn).

- 👍 → củng cố qua đúng `learn_pattern` sẵn có (vốn đã từ chối small talk và
  intent không phải tool), **và gỡ** chặn cũ nếu có.
- 👎 → chặn đúng cặp (cách diễn đạt, intent) đó.
  `routing.learned_pattern_wins()` tra danh sách này trước khi cho learned
  pattern thắng.

Phạm vi **cố ý hẹp**: không chặn intent ở mọi nơi, không chặn skill ở mọi nơi.
Một câu trả lời tệ là bằng chứng về **một** tuyến; vòng phản hồi khái quát quá
tay còn hại hơn không có.

Hàng trả lời mang theo **ảnh chụp cách định tuyến của chính nó**, chụp lúc dựng
hàng — nên một vote bấm muộn vẫn trỏ đúng lượt của nó, không phải lượt đang chạy.

---

## Kiểm chứng

```bash
python3 dev/test_assistant_perf.py      # telemetry, batching, classify, feedback, conversation
python3 dev/test_assistant_routing.py   # gồm nhóm "prompt cache"
python3 dev/test_knowledge.py           # gồm query_builder, rerank, xếp hạng skill
python3 dev/test_assistant_llm.py
python3 dev/test_skill_installer.py
python3 dev/test_llm_config.py
python3 dev/test_spell_checker.py
python3 dev/audit_tools.py --quiet
python3 dev/audit_ui.py --quiet
python3 dev/sync_wpf_styles.py --check
```

> `dev/test_tool_registry.py` từng báo **3 lỗi có từ trước** đợt này. Đã sửa —
> xem mục 6 bên dưới.

**Không có file XAML nào bị sửa** — toàn bộ đợt này là logic.
`T3LabAssistant.xaml` vẫn UI-locked.

---

## 6. Dọn nốt 4 lỗi có sẵn từ trước

Ba lỗi đỏ của `dev/test_tool_registry.py` **không cùng một loại** — mỗi lỗi sai
ở một phía khác nhau:

### 6.1 `revit_list_views` trả lời sai (lỗi thật)

Nhánh dispatch so `str(v.ViewType)` — tên enum của Revit (`"FloorPlan"`,
`"ThreeD"`) — trong khi schema **không khai báo enum** nào, và `create_view`
lại dạy model một bộ từ vựng khác hẳn (`floor_plan`, `3d`…).

Model vừa tạo view xong, gọi `revit_list_views(view_type='floor_plan')` và nhận
**`count: 0`** ⇒ đọc ra thành *"dự án không có mặt bằng nào"*. Im lặng, không
báo lỗi. `revit_list_views` nằm trong `ESSENTIAL_TOOL_NAMES` nên **model local
cũng dùng** — chính nhóm hay đoán sai từ vựng nhất.

`_VIEW_TYPE_FILTER_MAP` (module-level trong `core/server.py`) ánh xạ từ vựng
snake_case → tên `ViewType`. Chín khoá đầu **đúng bằng** từ vựng của
`create_view` để hai tool nói cùng thứ tiếng; phần còn lại là loại chỉ liệt kê
được (`schedule`, `detail`, `rendering`, `walkthrough`). Giá trị không khớp map
**vẫn rơi về so sánh cũ**, nên caller đang truyền `"FloorPlan"` không bị vỡ.

### 6.2 `collect_spellcheck_text` — nguồn sự thật đặt sai chỗ

Là collector nội bộ, chỉ `script.py::_run_spellcheck` gọi thẳng, **cố ý không**
quảng cáo cho LLM (nó đổ *mọi* chuỗi trong model ra). Cùng loại với
`__begin_action_group` / `__end_action_group` — nhưng danh sách loại trừ lại
nằm **cứng trong test**, nên cũ đi mỗi lần thêm pseudo-tool.

Nguồn sự thật chuyển về `server._INTERNAL_TOOLS`, test đọc bằng
`_frozenset_literal()` (helper đã có sẵn).

### 6.3 8 enum "mồ côi" của `create_view` — **test sai**

`create_view` uỷ nhiệm cho `core.advanced_view_manager`, nơi
`SUPPORTED_VIEW_TYPES` chính là 9 giá trị của schema. Nó **thật sự tạo được**
section/elevation/3d/legend. Test lại quét chuỗi thô trong thân nhánh nên chỉ
thấy `floor_plan` (giá trị duy nhất xuất hiện nguyên văn, vì là default).

`_effective_branch_source()` lần thêm **một bước**: source của module trong
extension mà chính nhánh đó import, cộng hằng số module-level của `server.py`
có tên xuất hiện trong nhánh. Một quy tắc, xử lý cả uỷ nhiệm (`create_view`)
lẫn bảng tra (`revit_list_views`).

Đã kiểm chứng quy tắc mới **không** quá lỏng: chèn một giá trị enum bịa
(`teleport_view`) thì test vẫn bắt được. Hai test mới khoá thêm hai chiều —
khoá của `_VIEW_TYPE_FILTER_MAP` phải khớp enum của schema, và `create_view`
phải giữ đủ 9 giá trị.

### 6.4 `Snippets/_revisions.py` không parse được

`revision_type = RevisionNumberType.None` trong chữ ký hàm — `None` là **từ
khoá**, `X.None` là SyntaxError ⇒ file **chưa bao giờ import được**. Không file
nào import nó nên chưa ai vấp phải. Nay default là `None`, phân giải trong thân
hàm bằng `getattr(RevisionNumberType, 'None', None)` — cũng tránh phân giải lúc
định nghĩa hàm, vì `NumberType` đã obsolete từ RVT 2023.

Lỗi này sống lâu vì **không gì kiểm tra**: `audit_tools.py` chỉ soi pushbutton,
`lib/Snippets/` không có test nào. Thêm `test_every_extension_file_parses` —
`ast.parse` mọi file `.py` trong `T3Lab.extension/` (chỉ parse, **không**
import: phần lớn cây cần Revit thật).

---

## 7. Rà soát tiếp theo (2026-08-02) — 2 lỗi đã sửa, 4 khoản còn để lại

Đợt trước (mục 1–6) đã khớp lại tốc độ, skill scoring, retrieval và bộ nhớ hội
thoại. Đợt này soát riêng tầng **quyết định** (`nlu_engine.py`, `routing.py`,
`agent_loop.py`) và tầng **tích hợp dữ liệu ngoài** (`Services/revit_context.py`,
`api_learner.py`, `api_updater.py`) tìm khiếm khuyết còn sót — không lặp lại
phạm vi đợt trước.

### 7.1 Đã sửa

**a. `_last_tool_from_history` không phân biệt vai trò (`nlu_engine.py`)**

Nhánh khớp `"Opening X..." / "đang mở X"` để tìm tool assistant vừa mở lẽ ra
chỉ nên đọc dòng **của assistant**, nhưng vòng lặp quét cả `user` lẫn
`assistant`. Một tin nhắn user tình cờ chứa cụm đó ("đang mở dwg management
giúp tôi với") bị hiểu nhầm thành assistant vừa mở DWG Management, và một câu
"mở nó" ngay sau sẽ mở nhầm tool đó dù chưa tool nào thực sự chạy. Đã thêm
điều kiện `entry.get("role") == "assistant"` trước khi thử khớp dòng này —
nhánh khớp tên tool theo `resolve_tool(raw, exact_only=True)` ngay dưới vẫn
đọc cả hai vai vì nó chỉ khớp khi CHÍNH tin nhắn là một yêu cầu mở tool, không
phân biệt ai gửi. Test: `test_history_referent_is_a_real_tool_mention` (thêm
case mới) trong `dev/test_assistant_routing.py`.

**b. Race khi `Raise()` thất bại có thể xoá nhầm task của người khác (`revit_context.py`)**

`run_in_api_context` xếp `(func, on_done)` vào `_TASKS` rồi gọi
`ExternalEvent.Raise()`; nếu `Raise()` ném lỗi, code cũ gọi `_TASKS.get_nowait()`
để "lấy lại task vừa xếp" — nhưng `Queue` là FIFO, và assistant **hỗ trợ nhiều
cửa sổ cùng lúc**, nên nếu một cửa sổ khác đã xếp task trước, `get_nowait()`
lấy nhầm task đó (đứng đầu hàng đợi) chứ không phải task của lệnh gọi đang lỗi.
Task bị lấy nhầm bị vứt bỏ im lặng — `on_done` của nó không bao giờ chạy, và
lệnh gọi kia cứ treo chờ mãi dù đã nhận `'queued'`.

Mỗi lần xếp hàng giờ kèm một token định danh (`object()` riêng); khi `Raise()`
lỗi, `_remove_task(token)` lọc đúng entry mang token đó ra khỏi hàng đợi (giữ
nguyên thứ tự mọi entry khác) thay vì lấy bất kỳ thứ gì đứng đầu. Test:
`test_raise_failure_only_discards_its_own_task` (mới).

Cả hai test đều **fail trước bản sửa này** và pass sau đó; không đổi hành vi
khi chỉ có một task/một cửa sổ.

### 7.2 Còn để lại — ưu tiên theo impact/effort

| # | Vấn đề | File | Impact | Effort |
|---|---|---|---|---|
| 1 | `chat_stream()` khi mất kết nối SSE giữa chừng thì âm thầm gọi lại `chat()` không streaming từ đầu — user thấy câu trả lời đang stream bị thay bằng một câu trả lời khác được sinh độc lập, tốn gấp đôi latency/token mà không đảm bảo khớp nhau | `claude_provider.py`, `openai_provider.py` | Trung bình–Cao | Nhỏ |
| 2 | `assistant_memory.add_fact()` chỉ dedupe theo so khớp chuỗi thường tuyệt đối, không có `update_fact`/`forget_fact` — user sửa lại một quy ước đã lưu ("prefix WH-" → "đổi thành EA-") thì cả hai fact cùng tồn tại và cùng được nhồi vào system prompt mỗi lượt, không có tín hiệu cái nào còn hiệu lực | `assistant_memory.py` | Trung bình–Cao | Nhỏ–Vừa |
| 3 | `api_updater.py` (RevitAPIUpdater / auto_check_and_update / APIUpdateNotifier) không được import ở bất kỳ đâu khác trong repo — tính năng "tự phát hiện phiên bản Revit API mới" quảng cáo trong docstring không chạy | `api_updater.py` | Thấp–Trung bình | Nhỏ |
| 4 | `feedback.skill_score()` tính điểm 👍/👎 theo từng skill nhưng không route/ranking nào đọc lại số đó — dữ liệu thu thập rồi bỏ không | `feedback.py` | Thấp | Nhỏ |

Không đề xuất động vào `T3LabAssistant.xaml` (UI-locked) hay lặp lại phạm vi
mục 1–6. `knowledge/embeddings.py` đã kiểm tra riêng: **không phải dead code**
— được `knowledge_store.search()`/`embed_pending()` dùng thật qua nhánh
knowledge-agent; đường inject ngữ cảnh của agent chính (`_build_knowledge_reference`)
cố tình chỉ dùng BM25 để giữ tốc độ/tính tất định, đây là đánh đổi đã ghi nhận
chứ không phải thiếu sót.

---

## 8. Rà soát vòng hai (2026-08-02) — phạm vi rộng hơn

Vòng này soát các file chưa được vòng 7 chạm tới: `skills_engine.py`,
`t3lab_agent.py`, `t3lab_assistant.py`, `agents/{dispatcher,task_manager,
specialists}.py`, `Services/mcp_service.py`, `knowledge/context_digest.py`,
`telemetry.py`, `config/project_store.py`, `skill_installer.py`, và các bảng
tra/dispatch trong `core/server.py`. `dispatcher.py`, `specialists.py` và
phần lớn `skills_engine.py` đã được làm chắc ở các đợt trước — không thấy gì
mới, thật ở đó.

### Đã sửa

**`/skills install` / `update` có thể ghi đè một skill người dùng tự viết (`skill_installer.py`)**

`remove_installed()` từ chối xoá một thư mục skill không có stamp
`.t3lab-source.json` — coi đó là "không phải của mình, không được đụng vào".
Nhưng `_install_one()` (đường ghi, dùng chung bởi cả install lẫn update) lại
không có rào tương ứng: nếu người dùng tự tạo một skill thủ công và tình cờ
trùng id/slug với một skill trong repo đang cài, lần install/update kế tiếp
ghi đè thẳng `SKILL.md` và toàn bộ resource của họ, không cảnh báo.

Đã thêm cùng một điều kiện ở đường ghi: thư mục đích đã tồn tại **và** không
có stamp thì bị bỏ qua, đưa vào `report['skipped']` kèm lý do, không ghi gì
vào đó. Test: `test_install_does_not_overwrite_a_hand_authored_skill` trong
`dev/test_skill_installer.py` (fail trước bản sửa, pass sau).

### Còn để lại — ưu tiên theo impact/effort

| # | Vấn đề | File | Impact | Effort |
|---|---|---|---|---|
| 1 | Rescan tăng dần của `context_digest.py` chỉ coi là "đổi" khi **số lượng** tài liệu đổi; một tài liệu đã theo dõi mà nội dung đổi trạng thái đọc-được ↔ không-đọc-được (số lượng không đổi) bị bỏ sót, khiến "Project standards summary" cache cũ mà không ai biết | `knowledge/context_digest.py` | Trung bình | Nhỏ |
| 2 | `get_material_quantities`'s `QTY_CATEGORY_MAP` tra thẳng dict phân biệt hoa/thường, không có alias tiếng Việt — không nhất quán với `_resolve_bic()` (không phân biệt hoa/thường + alias) mà mọi tool loại khác đều dùng; dễ báo "Unsupported category" giả | `core/server.py` | Trung bình | Nhỏ |
| 3 | `Intelligence/agents/task_manager.py` (`AgentTaskManager`/`get_task_manager`) không có nơi nào trong repo import nó — cùng loại dead code với `api_updater.py` đã ghi ở mục 7, chỉ khác file | `agents/task_manager.py` | Thấp | Nhỏ |
| 4 | Dữ liệu `telemetry.py` ghi lại trung thực mỗi lượt nhưng nơi duy nhất đọc lại là script offline `dev/bench_assistant.py` — không có gì ở runtime hành động dựa trên nó | `telemetry.py` | Thấp | Vừa (để dùng thật) |

Không đề xuất động vào XAML (UI-locked). Không lặp lại phạm vi mục 1–7.
`t3lab_agent.py`, `t3lab_assistant.py`, `dispatcher.py`, `specialists.py`,
`mcp_service.py` đã được xem qua nhưng không tìm thấy khiếm khuyết mới đủ rõ để
báo cáo — hoặc đã được làm chắc từ trước, hoặc là lớp truyền dẫn mỏng không có
logic quyết định đáng kể để sai. Chưa soát hết toàn bộ 7361 dòng của
`server.py` dòng-theo-dòng — vòng này chỉ nhắm bảng tra và dispatch/error-handling
theo đúng phạm vi được giao.
