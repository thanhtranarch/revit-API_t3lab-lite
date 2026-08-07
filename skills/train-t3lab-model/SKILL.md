---
name: train-t3lab-model
description: Teach and fine-tune the T3Lab local Revit assistant. Use after connecting Claude to Revit via the T3Lab MCP server, when the user asks to "train the model", "teach the assistant", "dạy model", or "huấn luyện" the local T3Lab assistant. Drives the whole loop — enable teaching capture, mark a scratch model as the sandbox, demonstrate representative Revit tasks, then launch the local fine-tune.
---

# Train the T3Lab local model (teach → fine-tune)

You are teaching T3Lab's **local** Revit assistant (a small Qwen model) by
demonstrating real Revit work through the T3Lab MCP tools. Each demonstration is
captured as a training example; when enough are collected you launch a local
fine-tune. You (a strong model) are the teacher; the local model is the student.

**Prerequisite:** the T3Lab MCP server must be connected (the user opened Revit
and started the server in *MCP Control*). If tool calls fail with a
"Revit down" style error, stop and tell the user to open Revit and start the
T3Lab server in MCP Control, then re-run.

## Safety — read first
- Teaching mode restricts **model-modifying** tools to a single **sandbox**
  document. Only ever mark a **scratch / throwaway `.rvt`** as the sandbox —
  **never** a real project model. Writes to any other document are blocked while
  teaching is on.
- If the user has not opened a scratch model, ask them to open one before you
  perform any create/modify tasks. Read-only demonstrations are always safe.

## Steps

1. **Confirm the connection.** Call `list_revit_instances` (or
   `revit_get_project_info`). If nothing responds, stop and give the
   open-Revit-and-start-server instruction above.

2. **Enable capture.** Call `t3lab_set_teaching_mode` with `enabled: true`.

3. **Mark the sandbox.** Call `list_open_documents`. Identify the scratch model
   (ask the user which one if it is not obvious — do not guess a real project).
   Call `t3lab_mark_sandbox` with `document` set to that document's title or
   path. Confirm to the user which document is now the sandbox.

4. **Demonstrate tasks.** Do several **varied, representative** Revit tasks the
   office actually performs — a mix of read tasks (list levels, count walls,
   inspect a view, read parameters) and, on the sandbox only, write tasks
   (create a level, place a wall, tag walls, set a parameter, override colors).
   For **each** task:
   1. Call `t3lab_begin_teaching` with `goal` = the task phrased as a natural
      user request (e.g. "tag all walls on level 1", "đổi màu tất cả cửa sang đỏ").
   2. Perform the task with the normal T3Lab Revit tools — one clean, correct
      sequence of tool calls (this is exactly what the student will learn).
   3. Call `t3lab_end_teaching` with a one-line `summary` of the result.
   Aim for at least ~10–20 varied tasks; more and more diverse is better.

5. **Check readiness.** Call `t3lab_training_status`. It reports the example
   count, sources, last training, and whether a run is in progress.

6. **Train.** If there are enough examples, call `t3lab_train_model`. If it
   reports the dataset is below the minimum, either demonstrate more tasks
   (preferred) or, only if the user insists, call `t3lab_train_model` with
   `force: true`. Training runs **outside Revit** on a GPU and can take hours.

7. **Make it portable (no GPU).** Call `t3lab_build_exemplars`. This distils the
   captured teaching into a small git-tracked file (`teacher_exemplars.json`)
   that makes even a plain local model answer in the taught style on **any**
   machine — no fine-tune or GPU required. Tell the user to commit that file to
   share the improvement. (`t3lab_train_model` also refreshes it automatically;
   use `t3lab_build_exemplars` alone when you want the portable layer without a
   full fine-tune.)

8. **Report.** Tell the user how many trajectories were captured, whether a
   fine-tune was launched, how many portable exemplars were built, and that the
   fine-tune result becomes an Ollama model the assistant can point at (see
   `tools/train/README.md`). Remind them they can watch teaching state in
   *MCP Control*.

## Two ways the training travels to other machines
- **Fine-tuned model (best quality):** the weights live only in the training
  machine's Ollama. Distribute via `ollama push`/`pull` (a registry) or by
  copying the GGUF, then select it in *LLMs Setting* on each machine.
- **Portable exemplars (works everywhere immediately):** commit
  `teacher_exemplars.json`; any machine running a plain `qwen3:14b` picks it up
  from the extension and answers in the taught style — no re-train.

## Notes
- Do all model-editing demonstrations on the **sandbox** document. If a write is
  blocked, you targeted the wrong document — re-check `list_open_documents` and
  `t3lab_mark_sandbox`.
- Keep each demonstration a clean, minimal, correct tool sequence. Sloppy or
  failed sequences teach the student bad habits.
- One `t3lab_begin_teaching` … `t3lab_end_teaching` pair per task; don't nest.
