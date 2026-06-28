# Repository Guidelines

## Project Structure & Module Organization

This is a small Windows-focused Python print server. `main.py` contains the FastAPI application, embedded web interface, ZPL generation, printer routing, and diagnostics. `test_print_routing.py` holds the `unittest` suite, while `testdata/` stores golden ZPL output used by regression tests. Operational launch and recovery scripts live at the repository root (`start_zpl_server.bat`, `Reset-Label-Print-Server.ps1`, and `Reset Label Printer.bat`). Static assets, such as `favicon.ico`, also remain at the root.

## Build, Test, and Development Commands

- `uv sync` installs the locked Python 3.14 dependencies from `uv.lock`.
- `uv run main.py` starts the server at `http://127.0.0.1:8787`.
- `uv run python -m unittest -v` runs all automated tests.
- `uv run python -m unittest test_print_routing.MarkedDownPriceTests` runs one test class.
- `.\Reset-Label-Print-Server.ps1 -NoPause` restarts the local server and checks its health.

No separate build step is required.

## Coding Style & Naming Conventions

Use four-space indentation and standard Python conventions: `snake_case` for functions and variables, `PascalCase` for classes, and uppercase names for configuration constants. Keep type hints on public helpers and use descriptive FastAPI/Pydantic field names. Prefer standard-library features and existing helpers before adding dependencies or abstractions. Keep ZPL layout constants near the top of `main.py`. No formatter or linter is configured, so match the surrounding style and keep imports grouped by standard library, third party, and local modules.

## Testing Guidelines

Tests use Python's built-in `unittest` and `unittest.mock`. Name test methods `test_<expected_behavior>`. Add focused coverage for routing decisions, validation, and label layout changes. If rendered ZPL intentionally changes, update `testdata/peach_ringz_2x1.zpl` and verify the diff carefully. Tests must not contact a real printer; mock Windows and direct-TCP delivery functions.

## Commit & Pull Request Guidelines

Recent commits use short, imperative descriptions such as `Add configurable retail label layout and markdown support`. Keep each commit scoped to one behavior. Pull requests should explain the user-visible change, list tests run, and note printer or environment assumptions. Include screenshots for web-interface changes and sample ZPL diffs for label-layout changes. Link the relevant issue when one exists.

## Configuration & Safety

Configure printers through the `ZPL_*` environment variables documented in `README.md`; do not commit machine-specific addresses or secrets. Preserve the no-fallback-after-partial-direct-send behavior, which prevents duplicate labels.
