# Repository Guidelines

## Project Structure & Module Organization
- `dashboard.py` rebuilds KPIs, charts, and formatting in `Bet_Tracker.xlsx`; treat it as the automation entry point.
- `log_new_bets.py` drives the CLI for ingesting wagers; extend logging functions here and keep bet calculations consistent.
- Excel workbooks (`Bet_Tracker.xlsx`, backups) stay in the repository root; version experiments with suffixes such as `bet_tracker_2024-01-15.xlsx`.
- `requirements.txt` defines dependencies; update it whenever a new library is added.

<!-- ## Build, Test, and Development Commands
- `python -m venv .venv && source .venv/bin/activate` prepares the local environment.
- `pip install -r requirements.txt` installs Excel automation libraries.
- `python log_new_bets.py` records bets and should run before dashboard refreshes during QA.
- `python dashboard.py` regenerates KPIs and visualizations in the workbook. -->

## Coding Style & Naming Conventions
- Follow PEP 8: 4-space indentation, snake_case for names, and UPPER_SNAKE_CASE for constants like `FILE_PATH`.
- Add clear inline comments to explain Excel formulas, chart logic, and conditional formatting.
- Extract reusable workbook logic into helpers with docstrings (e.g. odds conversion, bet logging, KPI refresh).
- Scripts should remain import-safe with `if __name__ == "__main__":` guards.

## Excel Workflow Requirements
- Always use **openpyxl** (no pandas unless explicitly requested).
- Append rows for new bets — never overwrite existing bet data.
- Preserve headers, conditional formatting, formulas, and charts.
- Keep separation of concerns:
  - `log_new_bets.py`: bet logging, odds conversion, KPI updates.
  - `dashboard.py`: dashboards, charts, formatting, and summary KPIs.
- Handle missing/corrupted workbooks gracefully by creating new files with headers.


## Testing & Validation
- Manual testing for now: stage a copy of `Bet_Tracker.xlsx` and verify:
  - KPI totals,
  - chart data,
  - conditional formatting,
  - pending bet handling.
- Document edge cases: bonus bets, void bets, malformed dates.
- Prefer `pytest` when tests are introduced; place them under `tests/` and run with `pytest tests`.

<!-- ## Testing Guidelines
- No automated suite exists yet; stage a copy of the workbook and manually verify KPI totals, chart data, and pending-bet handling before merging.
- Prefer `pytest` when tests are introduced; place them under `tests/` and run with `pytest tests`.
- Document new edge cases (bonus stakes, void bets, malformed dates) in fixture data or step-by-step manual scenarios. -->

## Workflow with Codex
- When I say “log bet” or “extend logging,” edit `log_new_bets.py`.
- When I say “update KPI” or “add chart,” edit `dashboard.py`.
- Before changing Excel logic, check headers and formulas already in use.
- Ask for clarification if unsure which file to edit.


## Commit & Pull Request Guidelines
- Use concise, imperative commit subjects ("Add sportsbook ROI breakdown").
- Note why changes were needed in the body.
- PRs should list affected sheets/scripts and before/after screenshots of dashboards.

## Data Safety & Configuration Tips
- Treat `Bet_Tracker.xlsx` as sensitive; do not commit personal data. Extend `.gitignore` for local exports.
- Back up the workbook before schema changes (new columns, renamed sheets).
- Mention migrations (e.g. new KPI formulas) in PR descriptions.


<!-- ## Commit & Pull Request Guidelines
- Use concise, imperative commit subjects ("Add sportsbook ROI breakdown") and note why changes were needed in the body.
- PRs should list affected sheets/scripts, reproduction steps, and before/after screenshots whenever visuals change.
- Link issues or task IDs when available, and call out spreadsheet migrations so reviewers can back up their copy.

## Data Safety & Configuration Tips
- Treat `Bet_Tracker.xlsx` as sensitive; do not commit personal data and extend `.gitignore` for local exports.
- Back up the workbook before schema changes (new columns, renamed sheets) and mention required migrations in PR descriptions. -->
