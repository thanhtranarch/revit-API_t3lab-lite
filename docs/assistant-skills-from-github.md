# Installing skills from a GitHub repo

The T3Lab Assistant reads **skills** — markdown instruction packs that
activate on a request and steer the answer. Besides the ones shipped in
`lib/Intelligence/skills/`, it can install skills straight from a GitHub
repo, including repos written for **Claude** (Claude Code / claude.ai skill
packs), without any conversion by hand.

## Using it

**From the chat** — paste a repo link and ask:

```
cài skill từ https://github.com/anthropics/skills
install the claude skills from github.com/owner/repo
install skills from https://github.com/owner/repo/tree/main/document-skills
```

The command family works too:

| Command | What it does |
|---------|--------------|
| `/skills` | List the skills installed from a repo |
| `/skills install <github link>` | Download and install every skill in the repo |
| `/skills update` | Re-pull every repo-installed skill (one download per repo) |
| `/skills remove <id>` | Delete one repo-installed skill |

**From Settings** — *Assistant settings → Skills → Install from GitHub*
(and *Update*) run exactly the same installer.

All three signals must be present for a chat message to install anything: an
install verb, the word *skill* / *kỹ năng*, and a GitHub link. Pasting a repo
link while merely discussing it never writes to disk.

A `/tree/<branch>/<folder>` link narrows the install to that folder; a plain
repo link installs everything and follows the default branch, whatever it is
named.

## What gets installed

Skills land in `%APPDATA%\T3LabAI\skills\<id>\` as:

```
<id>/
├── SKILL.md            ← normalised frontmatter + the original body
├── scripts/ …          ← bundled resources, copied as-is
└── .t3lab-source.json  ← provenance: repo, ref, path, install time
```

The stamp is what makes `/skills update` possible — it re-downloads the same
repo and rewrites the same folder. Nothing inside a skill is ever executed;
the assistant only reads `SKILL.md`.

Two layouts are recognised, matching Claude's own conventions:

- a folder containing `SKILL.md` (`skills/pdf-processing/SKILL.md`)
- a loose `*.md` inside a `skills/` folder (`skills/quick-tip.md`)

`.git/`, `node_modules/`, `README.md` and friends are ignored, and binaries
and Windows auto-run scripts (`.exe`, `.bat`, `.ps1`, `.vbs`, …) are never
written — they are listed in the install report as skipped, along with
anything over 1 MB.

## Claude → T3Lab frontmatter

A Claude skill declares only what Claude needs:

```yaml
---
name: pdf-processing
description: >-
  Extracts text and tables from PDF files, fills forms and merges documents.
  Triggers: "extract pdf", "merge pdf", "fill pdf form".
allowed-tools:
  - Read
  - Bash
---
```

T3Lab's `SkillsEngine` reads a slightly different header, so the installer
bridges the gap:

| Claude | T3Lab | Handling |
|--------|-------|----------|
| `name`, `description` | same | kept; folded (`>-`) and literal (`\|`) scalars are joined |
| — | `triggers` | **derived** — see below |
| `allowed-tools` | `allowed-tools` | preserved verbatim, **never** written to `tools:` |
| — | `agents` | left empty = the skill applies to every specialist |
| — | `tools` | left empty = a reference playbook, answered from its body |

`allowed-tools` names *Claude Code* tools (Read, Bash, …), not T3Lab Revit
tools. Promoting them to `tools:` would make a reference playbook look
executable to the router, so they are kept for reference only.

### Trigger derivation

Claude leaves activation to the model reading `description`; T3Lab matches
literal phrases, so a skill with no `triggers` would sit in the registry and
never fire. `derive_triggers()` fills that in, conservatively:

1. an explicit `Triggers: …` line in the description or body — quoted
   phrases win over comma splitting (`Trigger`, `Từ khóa`, `Kích hoạt` are
   recognised too);
2. failing that, the skill's own name and id as a phrase.

Phrases shorter than 4 or longer than 60 characters are dropped, generic
words are filtered, and the list is capped at 12. An over-eager trigger
silently hijacks unrelated requests, which is worse than a skill that only
answers to `/slash` — when nothing usable is found, that is exactly what the
install report says.

Derivation also runs at scan time, so a Claude `SKILL.md` copied into the
skills folder by hand behaves the same as an installed one. An authored
`triggers:` line always wins.

## Code map

| File | Role |
|------|------|
| `lib/Intelligence/skill_installer.py` | Request parsing, download, discovery, install, update, reporting |
| `lib/Intelligence/skills_engine.py` | Frontmatter parser (YAML subset), `derive_triggers()`, registry |
| `T3LabAssistant.pushbutton/script.py` | `_try_skills_command()` — the deterministic chat path |
| `lib/GUI/LLMSettingDialog.py` | Settings → Skills buttons |
| `dev/test_skill_installer.py` | CPython-3 tests (no network — fetchers are injected) |

Run the tests with:

```bash
python3 dev/test_skill_installer.py
```
