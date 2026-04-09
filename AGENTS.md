# AGENTS.md

## AI Factory Policy

For this project, AI Factory is active and mandatory.

Every future agent working in this repository must use AI Factory rules as the default execution model.

## Required Skill Usage

Always activate `aif` first for any non-trivial task in this repository.

In addition to `aif`, the agent must also use the task-specific AI Factory skill that matches the work:

- Planning or decomposition: `aif-plan`
- Code exploration or impact analysis: `aif-explore`
- Architecture decisions: `aif-architecture`
- Implementation: `aif-implement`
- Bug fixing: `aif-fix`
- Quality improvements: `aif-improve`
- Verification and validation: `aif-verify`
- Code review: `aif-review`
- Documentation updates: `aif-docs`
- Roadmap and backlog work: `aif-roadmap`
- Commit preparation: `aif-commit`
- CI or automation changes: `aif-ci` or `aif-build-automation`
- Security-sensitive work: `aif-security-checklist`

## Enforcement

- Do not skip AI Factory for convenience.
- Do not perform substantial repository changes without selecting the relevant `aif-*` skill.
- If multiple task types apply, use `aif` plus the minimal required combination of `aif-*` skills.
- If a requested change does not clearly match one skill, start with `aif-plan`, then continue with the appropriate execution skill.

## Project Notes

- English-only VBA code and comments for all new or edited implementation.
- All new user-facing Russian text must be routed through localization modules or dictionaries, not hardcoded into VBA logic/comments.
- UserForms must be designed for variable screen resolutions: fit within the visible screen, center on open, and provide scrollable fallback on smaller displays.
- Project type: Excel VBA / UserForms / Ribbon XML / workbook automation
- File encoding for VBA exports: Windows-1251
- Prefer minimal, reversible changes
- Treat exported `.frm`, `.frx`, `.bas`, `.cls`, Ribbon XML, and workbook module files as source of truth for reviewable changes
