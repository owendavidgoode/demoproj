# Repository Guidelines

## Project Structure & Module Organization
- Keep runtime code in `src/` grouped by domain (e.g., `src/api`, `src/ui`, `src/jobs`); keep entrypoints lean (`src/index.*` or `src/main.*` wires modules together).
- Mirror code with tests in `tests/`; keep fixtures in `tests/fixtures/`.
- House scripts in `scripts/`, config templates in `config/` (copy to `.env.local` for secrets), docs in `docs/`, and static assets in `public/` or `assets/`.
- Avoid root sprawl: add new tooling under a folder (e.g., `config/eslint.config.*`, `scripts/format.sh`) instead of alongside `README.md`.

## Build, Test, and Development Commands
- `make install` — install dependencies (wraps the project’s package manager).
- `make dev` — run the app in watch mode.
- `make test` — run the suite; set `COVERAGE=1` to enforce coverage.
- `make lint` — run linters/formatters; keep zero warnings before committing.
- `make build` — produce the production artifact; ensure it passes after dependency changes.

## Coding Style & Naming Conventions
- Indentation: 2 spaces for JS/TS/JSON/YAML; 4 spaces for Python; lines ≤100 chars.
- Naming: kebab-case for web assets (`user-card.tsx`), snake_case for scripts (`sync_data.py`), singular domain nouns for modules (`user.ts` vs `usersHelper.ts`).
- Prefer small, pure functions with minimal exports; document only non-obvious decisions with short comments.
- Run formatters before pushing; if adding a formatter (e.g., Prettier, Ruff), place config under `config/` and wire it into `make lint`.

## Testing Guidelines
- Use a unit-first approach: `tests/unit/` for pure logic, `tests/integration/` for flows that touch IO.
- Name tests after behavior (e.g., `user_signs_in.spec.ts`); avoid external network calls by mocking or recording.
- Target ≥80% coverage on changed code; justify exclusions in code comments.
- Update `make test` when adding frameworks, markers, or flags (e.g., `make test TEST_PATTERN=...`).

## Commit & Pull Request Guidelines
- Commits: one logical change, imperative subject ≤72 chars (e.g., `Add user session renewal`), avoid WIP, and reference issues in the body (`Refs #123`).
- Pull requests: include a short summary, testing notes (`make test`, `make lint`), screenshots for UI changes, and rollout/rollback steps when relevant. Request review and wait for CI green before merging.

## Security & Configuration
- Never commit secrets; keep sample values in `.env.example` and load real values from environment or a secret store.
- Vet new dependencies; prefer standard libraries, pin versions, and run `make audit` (or equivalent) when dependencies change.
- Anonymize fixtures and logs; remove debug endpoints or temporary routes before release.
