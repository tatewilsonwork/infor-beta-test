# Stage envelope — Agent prompt template

The conductor renders this template once per stage and passes it as the prompt to the `Task` (Agent) tool. The sub-agent exports the handoff paths from the prompt body, reads `STAGE_INPUTS`, does its work, and writes `STAGE_OUTPUTS` before finishing. The conductor parses `outputs.json`, not the sub-agent's reply text.

Placeholders enclosed in `{{double_braces}}` are substituted by the conductor at render time.

> **The `Task`/`Agent` tool cannot set environment variables on the sub-agent.** So the conductor does **not** pass `STAGE_INPUTS` / `STAGE_OUTPUTS` / `DEAL_DIR` / `CLAUDE_PLUGIN_ROOT` as env vars — there is no parameter for that. Instead it renders their **absolute paths into the prompt body** (via the `{{…path}}` / `{{deal_dir}}` / `{{plugin_root}}` placeholders), and the sub-agent's **first step** is to `export` them itself so its SKILL.md reference commands (which read `os.environ[...]`) find them.

---

## Template

```
You are running stage `{{stage_id}}` of the `{{deliverable_type}}` plan for deal `{{codename}}`.

# First step — export your handoff paths (do this before anything else)

The Task tool could not set environment variables for you, so set them yourself
first. Run this exact block in your shell (bash/zsh):

    export STAGE_INPUTS="{{stage_inputs_path}}"
    export STAGE_OUTPUTS="{{stage_outputs_path}}"
    export DEAL_DIR="{{deal_dir}}"
    export CLAUDE_PLUGIN_ROOT="{{plugin_root}}"

On Windows PowerShell use instead:

    $env:STAGE_INPUTS  = "{{stage_inputs_path}}"
    $env:STAGE_OUTPUTS = "{{stage_outputs_path}}"
    $env:DEAL_DIR      = "{{deal_dir}}"
    $env:CLAUDE_PLUGIN_ROOT = "{{plugin_root}}"

Your SKILL.md reference commands read these from the environment (`os.environ[...]`),
so they must be exported in the same session before you run them.

# Deal context

- Codename: {{codename}}
- Deal directory: {{deal_dir}}
- Deliverable: {{deliverable_type}}
- Subject company: {{subject_company_summary}}
  (full company JSON in `{{deal_dir}}/deal.json` under `subject_company`)

# Your task

Load and follow the workflow in:

    {{plugin_root}}/skills/{{skill_name}}/SKILL.md

The conductor has already collected analyst inputs and resolved any references.
Your resolved inputs are at `{{stage_inputs_path}}` (also `$STAGE_INPUTS` once exported).

When you finish, write your structured outputs as JSON to `{{stage_outputs_path}}`
(also `$STAGE_OUTPUTS`). The conductor REQUIRES this file to exist before it
will run the next stage. Output keys must match the named outputs declared by the
plan for this stage:

{{declared_outputs_block}}

If you cannot complete the work, write `outputs.json` anyway with an `error: <reason>` key
explaining why, then exit. Do NOT exit silently — the conductor cannot resume past a
missing outputs.json.

Write any deliverable artefacts under `{{deal_dir}}/artefacts/` (NOT cwd) so the analyst
finds them in a predictable place.

# Constraints

- Voice, brand, and source-trust rules live in the skill's own SKILL.md / references.
  Follow those.
- Do not invoke other INFOR skills via Task yourself — the conductor handles cross-skill
  composition. If your SKILL.md says "then invoke skill X", check whether X is also a
  stage in this plan and skip the invocation if `$STAGE_OUTPUTS` is set.
```

---

## Placeholder substitution table

| Placeholder | Source |
|---|---|
| `{{stage_id}}` | Current `Stage.id` |
| `{{deliverable_type}}` | `Plan.deliverable_type` |
| `{{codename}}` | `DealContext.codename` |
| `{{deal_dir}}` | `DealContext.deal_dir` (absolute) |
| `{{subject_company_summary}}` | One-line render of `DealContext.subject_company` (legal_name + ticker if public, else "private") |
| `{{plugin_root}}` | `CLAUDE_PLUGIN_ROOT` env var (or `./infor-beta` default) |
| `{{skill_name}}` | `Stage.skill` |
| `{{stage_inputs_path}}` | Absolute path to `<run_dir>/stages/<stage_id>/inputs.json` |
| `{{stage_outputs_path}}` | Absolute path to `<run_dir>/stages/<stage_id>/outputs.json` |
| `{{declared_outputs_block}}` | Bulleted list of `OutputSpec` entries for this stage, e.g. `- deck_path: Path`. Empty bulleted line if no outputs declared. |

The conductor renders all four handoff paths **into the prompt body** (the export block in
the template) — it does **not** set them as environment variables on the `Task`/`Agent`
invocation, because that tool has no parameter for env vars. Render absolute paths:

- `{{stage_inputs_path}}`  → `<run_dir>/stages/<stage_id>/inputs.json` (absolute)
- `{{stage_outputs_path}}` → `<run_dir>/stages/<stage_id>/outputs.json` (absolute)
- `{{deal_dir}}`           → `DealContext.deal_dir` (absolute)
- `{{plugin_root}}`        → `CLAUDE_PLUGIN_ROOT` (or the `./infor-beta` default), absolute

The sub-agent exports them itself as its first step (see the template), so its SKILL.md
reference commands that read `os.environ[...]` resolve correctly.

---

## Why a file-based handoff

Per locked decision H2 (note 12):

- The sub-agent does not need to negotiate JSON formatting in its reply text.
- The conductor does not need a parser that handles every way an LLM might wrap a JSON block.
- `inputs.json` / `outputs.json` are machine-readable, version-controlled (in the run directory), and inspectable when something goes wrong.
- This matches note 5's run-log layout (`runs/<run-id>/stages/<id>/{inputs.json, outputs.json, log.txt}`).
