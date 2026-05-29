# Stage envelope — Agent prompt template

The conductor renders this template once per stage and passes it as the prompt to the `Task` (Agent) tool. The sub-agent reads `STAGE_INPUTS`, does its work, and writes `STAGE_OUTPUTS` before finishing. The conductor parses `outputs.json`, not the sub-agent's reply text.

Placeholders enclosed in `{{double_braces}}` are substituted by the conductor at render time.

---

## Template

```
You are running stage `{{stage_id}}` of the `{{deliverable_type}}` plan for deal `{{codename}}`.

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
Your resolved inputs are at the path in `$STAGE_INPUTS` (also `{{stage_inputs_path}}`).

When you finish, write your structured outputs as JSON to the path in `$STAGE_OUTPUTS`
(also `{{stage_outputs_path}}`). The conductor REQUIRES this file to exist before it
will run the next stage. Output keys must match the named outputs declared by the
plan for this stage:

{{declared_outputs_block}}

If you cannot complete the work, write `outputs.json` anyway with an `error: <reason>` key
explaining why, then exit. Do NOT exit silently — the conductor cannot resume past a
missing outputs.json.

# Environment variables already set for you

- `$STAGE_INPUTS`  — absolute path to your inputs.json
- `$STAGE_OUTPUTS` — absolute path where you must write outputs.json
- `$DEAL_DIR`      — absolute path to the deal directory; write any deliverable artefacts
                     under `$DEAL_DIR/artefacts/` (NOT cwd) so the analyst finds them in
                     a predictable place.

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

The conductor must also set the following environment variables on the Agent invocation:

- `STAGE_INPUTS=<absolute path>`
- `STAGE_OUTPUTS=<absolute path>`
- `DEAL_DIR=<absolute path>`
- `CLAUDE_PLUGIN_ROOT=<absolute path>` (if not already inherited)

---

## Why a file-based handoff

Per locked decision H2 (note 12):

- The sub-agent does not need to negotiate JSON formatting in its reply text.
- The conductor does not need a parser that handles every way an LLM might wrap a JSON block.
- `inputs.json` / `outputs.json` are machine-readable, version-controlled (in the run directory), and inspectable when something goes wrong.
- This matches note 5's run-log layout (`runs/<run-id>/stages/<id>/{inputs.json, outputs.json, log.txt}`).
