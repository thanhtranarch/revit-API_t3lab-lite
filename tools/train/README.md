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
   chat with (e.g. Llama-3.1-8B-Instruct ↔ `unsloth/llama-3.1-8b-instruct-bnb-4bit`):
   ```
   python3 tools/train/finetune_local.py \
       --base unsloth/llama-3.1-8b-instruct-bnb-4bit \
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

## Notes
- Ollama base models aren't HF checkpoints; Unsloth trains a HF base then exports
  GGUF. Match the family/instruct-format of your Ollama base for best results.
- Start with 1 epoch, LoRA r=16. Overfitting a tiny corpus makes the model worse
  — that's why `validate_dataset` enforces a floor and `--force` is required to
  train below it.
