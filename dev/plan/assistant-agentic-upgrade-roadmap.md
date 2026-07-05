# T3Lab Assistant — Agentic Upgrade Roadmap (mượt mà)

> Standalone roadmap — **không** thuộc bảng `gd<N> revitapi` trong `dev/plan/README.md`.
> Mục tiêu: đưa T3Lab Assistant (chat trong Revit) lên ngang khả năng các assistant
> Revit ngoài kia (revit-mcp + Claude Desktop / Cursor): hội thoại agentic nhiều bước,
> tự gọi tool tạo/sửa/truy vấn model — nhưng **không giật, không khóa UI, không treo Revit**.
>
> File liên quan:
> - Chat window:   `T3Lab.extension/T3Lab.tab/Support.panel/T3LabAssistant.pushbutton/script.py`
> - Providers:     `T3Lab.extension/lib/Intelligence/{claude,openai,deepseek,ollama,lmstudio}_provider.py`
> - Router:        `T3Lab.extension/lib/Intelligence/llm_router.py`
> - Tool registry: `T3Lab.extension/lib/core/server.py` (`_register_default_tools`, `_WRITE_TOOLS`, `_execute_tool_in_context`)
> - MCP service:   `T3Lab.extension/lib/Services/mcp_service.py`

## Hiện trạng (điểm xuất phát — đã có nhiều hơn tưởng)

Đã có sẵn:
- ✅ Vòng lặp tool trong `do_nlp()` (max 5 iteration): model trả JSON `{intent, params, message}`,
  nếu `intent` trùng tên tool trong `srv._tools` thì gọi `srv._execute_tool()` rồi feed kết quả lại.
- ✅ ~75 MCP tool (read + write) trong `core/server.py`; write tool marshal lên main thread
  qua `_WRITE_TOOLS` + ExternalEvent (đã tạo up-front trên UI thread — đúng convention).
- ✅ Streaming turn 1 (`_stream_llm_turn` + `StreamingJSONExtractor`), typing indicator,
  multi-provider (Claude/OpenAI/DeepSeek/Ollama/LM Studio), RAG đính kèm PDF/ảnh,
  fast-path DB answer, NLU offline, learned patterns, lịch sử theo document.

Điểm yếu so với assistant chuẩn (revit-mcp + Claude Desktop):
1. **Không dùng native function-calling.** Toàn bộ 75 schema JSON bị nhồi vào system prompt
   ở **mỗi iteration** → tốn token khủng khiếp, chậm, model nhỏ (qwen 1.5b) gãy JSON thường xuyên.
2. **Ép cả câu trả lời thành JSON** (`response_format=json_object`) → streaming phải "bóc" JSON
   sống bằng regex; câu trả lời hội thoại thuần cũng chịu overhead này.
3. **1 tool / iteration, max 5 iteration** → chuỗi việc dài (tạo lưới + level + tường + tag) bị đứt giữa chừng.
4. **Tool feedback thô**: bubble text `🔧 [Tool Call] ... (params: {json})` — không có card
   trạng thái chạy/xong/lỗi, không xem lại được kết quả gọn.
5. **Không có nút Stop** — user không hủy được agent loop đang chạy; input bị khóa cứng toàn thời gian.
6. **Mỗi tool = 1 Transaction riêng** → "tạo 5 tường + tag" = 6 lần Undo. Chưa có TransactionGroup.
7. **Không confirm hành động nguy hiểm** (delete_element, purge_unused…): model gọi là chạy luôn.
8. **Dispatcher.Invoke mỗi token** khi stream → giật với model nhanh/mạng tốt.

---

## Trạng thái triển khai (2026-07-04)

**Toàn bộ GĐ A + B + C + D1/D2 đã code xong** (`py_compile` sạch, audit_tools/audit_ui/
sync_wpf_styles --check đều OK) — **chưa smoke test trong Revit** (xem D3):
- Mới: `lib/Intelligence/tool_schema.py`, `lib/Intelligence/agent_loop.py`
- Sửa: `llm_provider.py` (helper OpenAI-format + `openai_chat_agent`),
  `claude_provider.py` (chat_agent stream + prompt caching), `openai_provider.py`,
  `deepseek_provider.py`, `ollama_provider.py`, `core/server.py`
  (pseudo-tool `__begin/__end_action_group`, `select_elements` +`show`,
  `export_image` trả `files`), `T3LabAssistant/script.py`
  (nhánh native trong `do_nlp`, `_run_native_agent`, tool card + element link,
  confirm card, vision view capture, doc guard, nút Stop, throttle,
  catalog tiết kiệm + `describe_tool` cho nhánh JSON cũ).
- Giới hạn v1 (chủ đích): (1) request có file đính kèm (RAG/vision) vẫn đi đường cũ;
  (2) OpenAI/DeepSeek/Ollama gọi blocking (không SSE) trong agent mode — text hiện
  nguyên khối, chỉ Claude stream thật; (3) nút Stop hiệu lực với nhánh native,
  nhánh JSON cũ chạy nốt bước hiện tại; (4) model Ollama không hỗ
  trợ tools sẽ fail lượt 1 → tự rơi về nhánh JSON cũ (đúng thiết kế);
  (5) C2 vision view chỉ chạy trên nhánh agent với Claude (agent call của các
  provider khác không convert Claude-format image block); (6) D1 doc-guard chỉ
  áp dụng nhánh native (nhánh JSON cũ tối đa 5 iteration, rủi ro thấp);
  (7) B4 group không mở cho tool không đổi model (select/export/view/UI).

## GĐ A — Native tool calling (nền tảng, làm trước)

- [x] **A1. `chat_agent()` trên provider** — thêm method mới (không đụng `chat`/`chat_stream` cũ):
  - `claude_provider.py`: Messages API với `tools=[...]`, xử lý `stop_reason=tool_use`,
    trả `tool_result` block ở lượt sau. Bật **prompt caching** (`cache_control` trên block
    system + tools) → tool catalog chỉ tốn token 1 lần / 5 phút.
  - `openai_provider.py` / `deepseek_provider.py` (OpenAI-compatible): `tools` + `tool_calls`.
  - `ollama_provider.py`: `/api/chat` với `tools` (qwen2.5 ≥1.5b, llama3.1+ hỗ trợ).
  - `lmstudio_provider.py` + model không hỗ trợ tools: **fallback giữ nguyên** JSON protocol hiện tại.
  - Interface thống nhất trả về: `{"text": ..., "tool_calls": [{name, args, id}], "stop": ...}`.
- [x] **A2. Bộ chuyển schema** — `Intelligence/tool_schema.py`: đọc `srv._handle_tools_list()`
  → convert 1 lần sang format Anthropic / OpenAI / Ollama (cache theo session).
  KHÔNG nhồi schema vào system prompt nữa khi provider hỗ trợ native tools.
- [x] **A3. Tách agent loop ra module riêng** — `Intelligence/agent_loop.py`
  (script.py đã 3600 dòng): class `AgentLoop` chạy trên background thread,
  nhận callbacks `on_text_delta / on_tool_start / on_tool_done / on_finish / is_cancelled`.
  - Max iteration nâng lên 10, có budget thời gian tổng (vd 120s) + budget token.
  - Vẫn gọi tool qua `srv._execute_tool()` — giữ nguyên cơ chế ExternalEvent cho write tool
    (KHÔNG đổi; đây là phần đã đúng và đã ổn định).
  - Hỗ trợ nhiều tool_calls trong 1 lượt trả lời (chạy tuần tự — Revit single-thread).
- [x] **A4. Bỏ ép JSON cho câu trả lời hội thoại** khi dùng native tools: text ra là text,
  stream thẳng vào bubble, không cần `StreamingJSONExtractor` trên nhánh này.

## GĐ B — Mượt mà (UI/UX) — làm song song ngay sau A3

- [x] **B1. Tool-call card** thay bubble text: card gọn trong chat gồm tên tool + tóm tắt args,
  trạng thái ⏳ đang chạy → ✓ xong (kèm thời gian) / ✗ lỗi (kèm message), click mở rộng xem
  JSON kết quả. Xây bằng code-behind WPF như `_make_cmd_card` hiện có; palette Lumina.
- [x] **B2. Nút Stop**: khi `_busy=True`, nút Gửi biến thành nút ⏹ Stop.
  Stop = set cancel flag → agent loop kiểm tra giữa các iteration + abort stream reader.
  Tool đang chạy trong ExternalEvent thì để chạy nốt (không thể hủy an toàn giữa Transaction),
  nhưng không phát iteration mới.
- [x] **B3. Throttle stream UI**: gom token, `Dispatcher.BeginInvoke` tối đa ~25 lần/giây
  (tick 40ms hoặc buffer + flush) thay vì Invoke mỗi delta. Dùng `BeginInvoke` (không chặn
  worker) cho delta; chỉ `Invoke` ở các mốc bắt buộc đồng bộ.
- [x] **B4. TransactionGroup / 1 yêu cầu**: agent loop mở `TransactionGroup` (qua một tool-context
  wrapper trong `server.py`) trước tool write đầu tiên của 1 request, `Assimilate()` khi kết thúc
  → toàn bộ chuỗi hành động = **1 lần Undo** đúng nghĩa. Nút ↺ Undo hiện có hưởng lợi trực tiếp.
  Lưu ý: group phải mở/đóng trên main thread — thực hiện bên trong ExternalEvent handler,
  điều phối bằng request-id từ agent loop.
- [x] **B5. Confirm hành động phá hủy**: danh sách `_DESTRUCTIVE_TOOLS = {delete_element,
  purge_unused, ...}` — trước khi chạy, hiện card Confirm/Cancel trong chat (chờ user bấm,
  agent loop block trên event có timeout). Mặc định `purge_unused` luôn `dry_run=True` lượt đầu.
- [x] **B6. Trạng thái typing theo ngữ cảnh thật**: typing indicator hiển thị
  "Đang chạy `create_wall` (2/5)…" từ callback on_tool_start (thay text đoán theo giây hiện tại).

## GĐ C — Ngang khả năng assistant ngoài (context giàu hơn)

- [x] **C1. Element link bấm được**: kết quả tool có element id → render Hyperlink trong bubble,
  click = `uidoc.Selection.SetElementIds` + `uidoc.ShowElements` (marshal qua ExternalEvent read-safe).
- [x] **C2. Vision view hiện tại**: lệnh "nhìn view này/kiểm tra bố cục" → chụp active view
  (`ExportImage` thu nhỏ ~1280px, chạy qua ExternalEvent) rồi gửi như attachment ảnh
  (tái dùng `rag_processor.build_vision_content_blocks`). Chỉ với provider vision (Claude/GPT-4o).
- [x] **C3. Selection-aware mặc định**: khi user nói "các element này / đang chọn",
  agent tự đổ selected ids vào args (ContextScout đã có dữ liệu — chỉ cần khai báo
  quy ước trong system prompt + tool `revit_get_selected_elements` ưu tiên).
- [x] **C4. Catalog tool tiết kiệm cho model local**: nhánh fallback JSON (LM Studio / model nhỏ)
  chỉ đưa danh sách `tên — mô tả 1 dòng`, thêm meta-tool `describe_tool(name)` để model
  xin schema chi tiết khi cần → giảm ~80% token prompt cho máy yếu, đỡ gãy JSON.

## GĐ D — An toàn & kiểm thử

- [x] **D1. Guard đổi document**: lưu `doc_key` lúc bắt đầu request; nếu user đổi document giữa
  chừng → hủy loop, báo trong chat (tránh write nhầm model).
- [x] **D2. `py_compile` + regression**: `python -m py_compile` các file sửa;
  `python3 dev/audit_tools.py --quiet` · `python3 dev/audit_ui.py --quiet` ·
  `python3 dev/sync_wpf_styles.py --check`.
- [ ] **D3. Smoke test trong Revit** (checklist cho user — không tự bịa kết quả):
  1. Chat thường (không tool) — stream mượt, không giật khi trả lời dài.
  2. "tạo 3 tường 5m tại Level 1 rồi tag" — thấy tool card chạy tuần tự, xong = 1 lần Undo.
  3. Bấm Stop giữa chừng chuỗi tool — dừng sau tool hiện tại, input mở khóa.
  4. "xóa element ID x" — hiện card Confirm, Cancel thì không xóa.
  5. Model local (qwen2.5:1.5b) — nhánh fallback vẫn hoạt động, không crash;
     hỏi 1 lệnh tool lạ → model tự xin schema qua `describe_tool` rồi gọi đúng.
  6. Kéo/resize window trong lúc agent đang chạy — UI không đơ (không có blocking Invoke dài).
  7. Đổi document giữa chừng — loop tự hủy, có thông báo "đã chuyển document".
  8. Tool card kết quả có element id → click `#id` chọn & zoom đúng element trong Revit.
  9. (Claude + key vision) "nhìn view này và nhận xét bố cục" — assistant chụp active view
     rồi mô tả đúng nội dung ảnh.
  10. Sau chuỗi "tạo 3 tường + tag": bấm Undo 1 lần trong Revit → mất TOÀN BỘ chuỗi
      (TransactionGroup assimilate đúng), không phải undo từng bước.

---

## Thứ tự triển khai đề xuất

| Bước | Việc | Phụ thuộc | Ghi chú |
|------|------|-----------|---------|
| 1 | A2 + A3 (tách agent_loop + schema converter) | — | Refactor thuần, chưa đổi hành vi |
| 2 | A1 Claude trước (native tools + prompt caching) | 1 | Provider chính, lợi nhất |
| 3 | B1 + B2 + B3 (card, Stop, throttle) | 1 | Cảm giác "mượt" đến từ đây |
| 4 | A1 OpenAI/DeepSeek/Ollama + A4 | 2 | |
| 5 | B4 + B5 (TransactionGroup, confirm) | 3 | Đụng `server.py` — cẩn thận ExternalEvent |
| 6 | C1 → C4 | 3 | Từng cái độc lập, ship dần |
| 7 | D1 → D3 | 5 | Chốt bằng smoke test trong Revit |

## Ràng buộc kỹ thuật phải nhớ

- IronPython 2.7: không `async/await`, HTTP bằng `System.Net`/`urllib2` — SSE streaming đã
  chứng minh chạy được trong `chat_stream`, native tools chỉ là payload JSON khác → khả thi.
- ExternalEvent **phải** được tạo trên UI thread (đã làm up-front trong `__init__` — giữ nguyên).
- Mọi write tool đi qua `_WRITE_TOOLS` + ExternalEvent — tuyệt đối không mở Transaction
  từ worker thread (xem memory: TagChecker precedent).
- UI update từ worker: delta dùng `BeginInvoke` + throttle; mốc trạng thái dùng `Invoke`.
- `T3LabAssistant.xaml` không thuộc UI-frozen list, nhưng thay đổi UI phải theo Lumina
  (`.claude/rules/ui-design-standard.md`); style block auto-sync không được sửa tay.
