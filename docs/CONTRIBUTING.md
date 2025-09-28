# Contributing Guide

Thanks for considering a contribution!

## Development Setup
```
uv venv
uv pip install -e .[dev]
```
Optional (schema validation):
```
uv pip install jsonschema
```

## Project Principles
- Single-file core (`main.py`) stays readable.
- Deterministic output for identical JSON input.
- Schema-first: update `ppt.schema.json` when adding new fields.
- Graceful degradation: never crash the whole build for one failed element.

## Pull Requests
1. Fork & branch: `feat/..`, `fix/..`, `docs/..`
2. Add/adjust tests if behaviour changes.
3. Update README / schema where relevant.
4. Keep changes focused (avoid large unrelated refactors).
5. Ensure linters pass (add a pre-commit config if desired).

## Tests (proposed minimal pattern)
- Generate PPT from `examples/minimal.json` -> assert file exists & size > 0.
- (Optional) Round-trip via `ppt_to_json.py` for sanity.

## Coding Style
- Prefer pure functions where practical.
- Avoid hidden global state (except controlled caches like `_IMAGE_CACHE`).
- Log warnings instead of raising where recoverable.

## Adding A New Element Type
1. Extend schema (`ppt.schema.json`).
2. Add handler in `build()` dispatch.
3. Implement renderer (follow existing naming like `add_<type>`).
4. Document usage in README + example JSON.

## License
By contributing you agree your code is released under the project MIT License.
