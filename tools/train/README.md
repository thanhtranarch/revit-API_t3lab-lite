# T3Lab local fine-tune (`tools/train/`)

The out-of-Revit half of the self-improving assistant. The in-Revit self-study
loop (`Intelligence/learning/`) curates a chat-SFT dataset while Revit is idle;
this pipeline turns that dataset into a **local Ollama model** the assistant
serves. Weight training cannot run inside Revit/IronPython — it needs CPython 3 +
a GPU — so it lives here and runs separately.

```
in-Revit idle loop  ──►  %APPDATA%/T3LabAI/training/dataset.jsonl
                                     │
        tools/train/finetune_local.py (CPython3 + GPU)
                                     │
                     GGUF + Modelfile  ──►  ollama create t3lab-assistant
                                     │
                 assistant OllamaProvider points at t3lab-assistant
```

## Files
- `validate_dataset.py` — pre-flight: enough valid, deduped examples to train?
  Pure Python, no ML deps. `python3 tools/train/validate_dataset.py`.
- `finetune_local.py` — LoRA fine-tune (Unsloth). `--dry-run` validates + writes
  a Modelfile without training (works with no GPU/deps).
- `Modelfile.tmpl` — the Ollama Modelfile written next to the trained GGUF.

## Requirements (real run only)
- NVIDIA GPU (≥8 GB for an 8B 4-bit LoRA), CUDA, Python 3.10+.
- `pip install "unsloth[cu121] @ git+https://github.com/unslothai/unsloth.git" trl datasets transformers torch`
- Ollama installed (`ollama --version`).

## Recipe
1. **Check the corpus:**
   ```
   python3 tools/train/validate_dataset.py
   ```
   Needs ≥ 30 unique examples (raise this in practice — a few hundred is better).
2. **Dry run** (anywhere, no GPU): validates + writes `out/Modelfile` + `out/sft.jsonl`.
   ```
   python3 tools/train/finetune_local.py --dry-run
   ```
3. **Train** (GPU box). Pick a `--base` HF model that matches the Ollama base you
   chat with. The assistant auto-selects **qwen3:14b** as its default agentic
   model, so distil INTO a Qwen base (default `unsloth/Qwen2.5-14B-Instruct-bnb-4bit`;
   use a Qwen3 checkpoint if available). `--base` overrides for other families:
   ```
   python3 tools/train/finetune_local.py \
       --base unsloth/Qwen2.5-14B-Instruct-bnb-4bit \
       --tag t3lab-assistant --epochs 1
   ```
   Produces `out/gguf/*.gguf` + `out/Modelfile` and writes
   `%APPDATA%/T3LabAI/training/last_train.json`.
4. **Register with Ollama:**
   ```
   ollama create t3lab-assistant -f tools/train/out/Modelfile
   ```
5. **Point the assistant at it:** in *LLMs Setting* choose provider **Ollama** and
   model **t3lab-assistant**, or run `/model use t3lab-assistant` in the chat.

## Cadence
Fine-tuning is heavy — run it periodically (e.g. weekly, or when the dataset has
grown by a few hundred examples), not continuously. The idle loop can auto-launch
this as a detached process when `agents.self_train_auto` is enabled (Phase 3),
but only when a GPU + trainer are present and Revit is not mid-operation.

## Privacy
The dataset and all artifacts stay on your machine (`%APPDATA%/T3LabAI`), same as
telemetry. Nothing is uploaded.

## Distillation from Opus 5 (optional teacher)
The corpus is mostly the office's own successful commands (zero-cost enrichers).
Two ways to also distil **Opus-5-quality** behaviour into it:

**a) Text answers (`opus_teacher`).** Enable:
```json
{ "agents": { "self_study": true, "opus_teacher": true } }
```
with a Claude API key set in *LLMs Setting*. While Revit is idle, Opus writes
gold answers to `dataset.jsonl` (`meta.source = "opus_teacher"`, `quality =
"teacher"`) for curated Revit/BIM questions plus the office's own terse commands.
The ONE enricher that spends API tokens — off by default — and it runs even while
the assistant itself chats on local Qwen.

**b) Tool-use trajectories (`mcp_teacher`).** Teach the model to *drive Revit*,
not just answer. In **MCP Control**, turn on **Teaching Capture** and mark a
scratch `.rvt` as the **Sandbox** (model writes are then blocked everywhere else).
Connect **Claude Desktop** (Opus) to Revit via the MCP bridge and perform tasks;
each `t3lab_begin_teaching(goal)` … tool calls … `t3lab_end_teaching(summary)`
sequence is stored as an agentic trajectory (`meta.source = "mcp_teacher"`) with
`user`/`assistant(tool_calls)`/`tool` turns. The session miner (idle enricher)
also sweeps any raw session files into the dataset. These trajectory rows teach
the local Qwen to call the right tool with the right arguments — exactly what
small models get wrong. `export_sft` and the trainer keep the tool turns
(Qwen2.5/3 chat templates support the `tool` role), so no extra flags are needed.

## Notes
- Ollama base models aren't HF checkpoints; Unsloth trains a HF base then exports
  GGUF. Match the family/instruct-format of your Ollama base for best results.
- Start with 1 epoch, LoRA r=16. Overfitting a tiny corpus makes the model worse
  — that's why `validate_dataset` enforces a floor and `--force` is required to
  train below it.
