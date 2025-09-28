<!-- English is the default. Chinese version separated: see README.zh-CN.md / 中文独立文件：README.zh-CN.md -->

# JSON2PPT

[![Status – Experimental](https://img.shields.io/badge/status-experimental-orange)](./LIMITATIONS.md)
[![Deterministic / Idempotent](https://img.shields.io/badge/output-deterministic%20%2F%20idempotent-blue)](#determinism--idempotency-)
[![LLM Friendly](https://img.shields.io/badge/LLM-friendly-7B3FE4)](docs/PROMPT_FOR_LLM.md)
[![Single File Core](https://img.shields.io/badge/core-single--file-lightgrey)](./main.py)

English | [中文说明 Chinese](./README.zh-CN.md)

Declarative, LLM‑friendly PowerPoint generator. Provide a strictly structured JSON ➜ receive a fully editable `.pptx`. No more manual alignment, font tweaking, or copy‑pasting numbers. Designed for automation pipelines & AI Agents requiring reproducible, auditably stable slide output.

> Vision: author “structured narrative data”, not slides. Rendering is a side effect. ✨

## Why (Motivation) 🚀
Traditional approaches are either brittle mail‑merge style templates or rasterized screenshots. JSON2PPT deliberately focuses on three pillars:

1. LLM First: JSON Schema is explicit, constrained and machine‑verifiable – large models quickly learn to emit valid payloads.
2. Fully Controllable: Position / size / style / charts / tables all explicitly declared → reproducible & traceable.
3. Minimal Surface: A single readable core file (`main.py`) → fast onboarding & easy for you (or a model) to extend.

## Determinism & Idempotency 🔁
Core principle: the same validated JSON input always produces the same PowerPoint (byte‑level stable except for timestamps that Office may add when reopened). Achieved by:

- Pure, order‑preserving traversal of `slides` & `elements`.
- Explicit units & coordinate resolution (`px` → EMU; percentage resolved against slide width/height deterministically).
- No random seeds, hashing tricks, or heuristic auto‑layout.
- Style merging is associative + predictable (`merge_styles(base, override)`).
- Image handling: optional caching layer prevents duplicate downloads; for strict reproducibility prefer `base64:` or `file:` sources (remote `url:` can change outside our control).

Result: the pipeline is idempotent – re‑running the generator on the exact same JSON does not drift or silently mutate formatting. This is critical for:

- Regression tests in CI (snapshot compare of produced `.pptx`).
- Multi‑agent planning where downstream steps rely on stable slide indexes & element ordering.
- Auditable reporting (e.g. monthly KPI decks regenerated on new data only).

## Feature Snapshot 🧰
| Domain | Capability |
| ------ | ---------- |
| Text 📝 | Single or multi paragraph (bullets / numbered) with rich run & paragraph styles |
| Images 🖼️ | `file:` / `base64:` / `url:` (with simple in‑memory cache) |
| Shapes 🧱 | Rectangle / RoundRect / Ellipse with fill colors |
| Tables 📊 | Header/body style separation, zebra stripes, per‑column widths |
| Charts 📈 | Bar / Line / Pie (easy to extend via mapping) |
| Background 🎨 | Solid color or full‑bleed image |
| Coordinates 📐 | Mixed `px` & `%` (converted to EMU) |
| Validation ✅ | Draft‑7 JSON Schema (`ppt.schema.json`) before rendering |

## Quick Start ⚡
1. Install (Python 3.13+):
   ```powershell
   uv venv
   uv pip install python-pptx requests
   ```
2. Prepare a JSON payload complying with the schema (see Minimal Example below or open `ppt.schema.json`).
3. Call `handler(args)` with `args.input.meta` set to the JSON string – or simply run:
   ```powershell
   python main.py
   ```
   to view the built‑in demo.

> Only local / embedded images? You may omit `requests`.

## Minimal Example (excerpt, 3 slides) 🧪
```jsonc
{
  "ppt": {
    "defaultUnit": "px",
    "slides": [
      {"title": "Cover", "elements": [
        {"type": "text", "text": "AI Market Watch", "box": {"x": 120, "y": 180, "w": 1040, "h": 160}, "style": {"fontSize": 64, "align": "center", "bold": true}}
      ]},
      {"title": "Structured Highlights", "elements": [
        {"type": "text", "box": {"x": 100, "y": 120, "w": 880, "h": 380}, "paragraphs": [
          {"text": "Weekly Highlights", "style": {"bold": true, "fontSize": 36}},
          {"text": "3 Product Launches", "listType": "bullet"},
          {"text": "5 Funding Events", "listType": "bullet"}
        ], "style": {"fontSize": 28}}
      ]},
      {"title": "Revenue Comparison", "elements": [
        {"type": "chart", "chartType": "bar", "box": {"x": 80, "y": 120, "w": 960, "h": 400},
          "data": {"categories": ["Q1", "Q2", "Q3"], "series": [
            {"name": "Actual", "values": [820, 860, 910]},
            {"name": "Target", "values": [800, 840, 900]}
          ]}}
      ]}
    ]
  }
}
```

More composition patterns:
- Full Schema: `ppt.schema.json`
- LLM Prompt & constraints: `PROMPT_FOR_LLM.md`

## LLM Workflow 🤖
1. Provide domain data (metrics / outline / descriptions)
2. Model generates JSON using constraints in `PROMPT_FOR_LLM.md`
3. (Optional) Validate with `jsonschema`
4. Feed into `handler` → `.pptx`

### Optional Validation Sample 🧪
```python
import json
from pathlib import Path
from jsonschema import validate
from main import handler

schema = json.loads(Path("ppt.schema.json").read_text(encoding="utf-8"))
payload = json.loads(Path("demo/example.json").read_text(encoding="utf-8"))
validate(payload, schema)

class A: ...
args = A(); args.input = A(); args.input.meta = json.dumps(payload)
handler(args)
```

## Entry Point 🛣️
`handler(args)` in `main.py` is the only public entry. If remote images (`url:`) are present be sure `requests` is installed. Output is written to the working directory (or a temp dir depending on how you integrate it).

## Suitable Scenarios ✅
- Batch deck generation (daily/weekly ops, KPI boards)
- A/B narrative template comparison
- Automated pipeline: natural language → structured metrics → slides
- External LLM plugins / agents producing PowerPoints directly

## Not Yet Suitable ⛔
- Complex master slides / animations / SmartArt / video / audio
- Highly bespoke advanced chart types beyond bar/line/pie

See `LIMITATIONS.md` for details.

## Extension Points 🔌
Search in `main.py`:
- `ELEMENT_TYPES` – register a new element type
- `add_chart` / `map_chart_type` – new chart mappings
- `add_shape` – extend shape enums
- `apply_run_style` / `apply_paragraph_style` – add style properties

## Troubleshooting 🔍
- Missing `python-pptx`: raises `RuntimeError`
- Remote image missing: check network & `requests`
- Color not applied: ensure valid hex
- Blank chart: ensure each `series.values` length matches `categories`

## Security / Hardening 🔒
Runtime limits (can be overridden via environment variables):

| Env Var | Default | Meaning |
| ------- | ------- | ------- |
| `JSON2PPT_MAX_SLIDES` | 200 | Maximum slides allowed |
| `JSON2PPT_MAX_ELEMENTS` | 600 | Max elements per slide |
| `JSON2PPT_MAX_IMG_BYTES` | 5242880 (5MB) | Max single image size |
| `JSON2PPT_ALLOW_REMOTE` | 1 | Allow `url:` images (0 disables) |
| `JSON2PPT_ALLOW_FILE` | 1 | Allow `file:` images |
| `JSON2PPT_REMOTE_DOMAINS` | (unset) | Comma list of allowed remote domains (whitelist) |
| `JSON2PPT_ASSET_ROOT` | (unset) | If set, `file:` paths must reside under this root |

Notes:
1. Remote images are denied if `requests` is missing or domain not in whitelist (when configured).
2. Oversized images or disallowed schemes are skipped (warning logged, build continues).
3. File images are restricted to `JSON2PPT_ASSET_ROOT` when provided.
4. Schema validation (if `jsonschema` installed) enforces structural integrity before rendering.

Recommended Production Setup:
- Run generation in a container / isolated worker.
- Set an asset root & whitelist remote domains or disable remote entirely.
- Keep `ppt.schema.json` versioned when extending fields.
- Add monitoring on warnings for early abuse signals.

## License & Contribution 🤝
Simple & pragmatic: Fork and tailor into your internal “auto report generator”. PRs welcome for broadly reusable enhancements (new chart types, layout helpers, richer text styling).

See `SECURITY.md` for vulnerability reporting and hardening guidance.

## Business / Hosted Service 📦
This project intentionally stays MIT to maximise adoption. Revenue strategy focuses on a hosted API + premium template packs + collaboration features (see `COMMERCIAL_STRATEGY.md`).

Planned hosted capabilities:
- Scalable render API (sync for small decks, async for large)
- Usage metering & team workspaces
- Template marketplace with signed packs
- Deterministic file hashing & history registry

Nothing in MIT prevents commercial usage; sustainability will derive from:
1. Operational convenience (no infra / scaling worries)
2. Rich curated templates & industry packs
3. Team governance (roles, logs, retention)
4. Faster feature cadence than forks

If you are interested in early hosted access or partnership: open a discussion or email (placeholder) biz@example.com.

---
Want the model to learn fast? Read `PROMPT_FOR_LLM.md`.

---
Looking for Chinese docs? → [README.zh-CN.md](./README.zh-CN.md)

