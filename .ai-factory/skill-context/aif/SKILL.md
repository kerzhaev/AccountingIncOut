# AccountingIncOut AI Factory Skill Context

These rules override generic AI Factory defaults for this repository.

## Project-Specific Priorities

- Preserve workbook behavior over cosmetic code normalization.
- Prefer source exports over binary workbook changes whenever possible.
- Treat exported VBA modules as the reviewable implementation surface.
- Minimize edits to `.frx` and workbook binary files.

## File And Encoding Rules

- Assume exported VBA files are Windows-1251 encoded.
- Do not assume shell-visible mojibake means the file is invalid.
- Be conservative when editing `.frm` files; validate the exact target file path before modifying.
- Avoid workflows that silently re-encode VBA source to UTF-8 unless the file is markdown or an explicitly UTF-8 project artifact.

## Change Selection Rules

- If the worktree is already dirty, do not sweep unrelated changes into the same commit.
- Separate source-code fixes from workbook-binary churn unless the user explicitly requests a bundled commit.
- Prefer minimal commits that map to one user-visible fix.

## Known Regression Hotspots

- `UserFormVhIsh.frm`
- `RecordOperations.bas`
- `LocalizationManager.bas`
- `RibbonCallbacks.bas`
- `workbook-modules/ЭтаКнига.cls`
- `TableEventHandler.bas`

## Mandatory Checks Before Finalizing

- Confirm whether changes affect dictionary-backed controls.
- Confirm whether runtime state is duplicated between a form and a standard module.
- Confirm whether callback names in Ribbon XML still match VBA procedures.
- Confirm whether localization lookup names still match actual sheet and table names.

## Documentation Expectation

When project context changes, update `.spec/PROJECT_CONTEXT.md` or another project-facing document in the same turn when feasible.

## UserForm Responsiveness Rule

- When changing Excel VBA UserForms, always account for the end user's screen resolution.
- UserForms must open within the visible screen area, center on screen, and avoid clipping critical controls on low-resolution displays.
- Prefer responsive resizing with scrollable fallback over fixed large layouts.
- For `UserFormPackageDocuments`, do not constrain `lstPackageItems` width to the right edge of the left-side button block. This made the package list look artificially narrow in runtime even when the form had wider usable space.
- Runtime probe on `UserFormPackageDocuments` showed `lstPackageItems.Width` can already be wide in runtime while the list still looks narrow to the user. In that case, inspect `ColumnWidths` before changing the control width again.
- For `UserFormPackageDocuments`, if records make `lstPackageItems` look visually narrow while the empty list looks fine, treat that as a `ColumnWidths` problem first, not a `Width` problem.
- For `UserFormPackageDocuments`, `lstPackageItems.ColumnWidths` may revert to designer-era values during form startup. If a runtime probe shows old widths despite updated code, re-apply `ColumnWidths` after `Show`/final resize instead of only during `Initialize`.
- When debugging UserForm layout regressions, distinguish between:
  - designer geometry stored in `.frm/.frx`
  - runtime geometry set by layout procedures such as `ApplyFormLayout`
  - viewport clipping/scrollbar behavior applied by responsive resize logic
- Do not treat a failed runtime layout hypothesis as resolved just because the exported `.frm` or workbook import contains the patch. Re-check the visible runtime outcome before keeping the change.
- Do not keep obsolete public VBA launch procedures after routing changes. They stay visible in `Alt+F8`, still participate in compilation, and stale variables inside them can break unrelated workflows. Replace them with a hidden compatibility wrapper or remove them.


## Code Language Rule

- All new or edited VBA code identifiers and comments must remain English-only.
- Do not add new Russian comments or Russian string literals directly in implementation code unless routed through localization lookups.
- User-facing text should go through LocalizationManager.GetText() or other localization data sources whenever feasible.

## Excel Automation Test Rule

- For `.xlsm` automated testing, use Python with `xlwings` or `pywin32` as the primary tool for direct Excel COM interaction.
- Do not use `openpyxl` or `pandas` to validate or simulate VBA macro behavior.
- Use PowerShell only for system-level tasks such as backups, process launch, or cleanup, not as the main end-to-end test runner for VBA logic.
- COM test scripts must catch COM exceptions, account for blocking modal dialogs or UserForms, and terminate stuck `EXCEL.EXE` processes when required.
- Test output should be emitted in a `pytest`-style console format whenever practical.

