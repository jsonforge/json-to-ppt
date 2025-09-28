# JSON-to-PPT Generator - AI Coding Instructions

## Project Overview
This is a **declarative PowerPoint generator** that transforms JSON configurations into complete `.pptx` presentations using `python-pptx`. The core architecture follows a **single-file design** with functional programming patterns.

## Core Architecture

### Entry Points & Data Flow
- **Main handler**: `handler(args)` in `main.py` - expects `args.input.meta` (JSON string) or `args.input.mets`
- **Processing pipeline**: JSON validation → `build()` → element rendering → file output
- **Key validation**: Always check for required `ppt.slides` structure before processing

### Element Type System
The generator supports 5 element types with distinct rendering functions:
- `text` → `add_text()` - Single paragraph text boxes with rich styling
- `image` → `add_image()` - Supports `base64:`, `url:`, `file:` source formats
- `shape` → `add_shape()` - Basic shapes (rect, roundRect, ellipse) with fill colors
- `chart` → `add_chart()` - Bar, line, pie charts with `CategoryChartData`
- `table` → `add_table()` - Grid layouts with per-cell styling

### Coordinate System & Units
- **Default unit**: `px` (pixels), converted to EMU via `unit_to_emu()` (1px = 9525 EMU)
- **Percentage support**: Relative to slide dimensions via `resolve_box()`
- **Box model**: Every element has optional `{x, y, w, h, unit}` positioning

### Image Handling Patterns
- **Caching mechanism**: `_IMAGE_CACHE` prevents re-downloading
- **Source prefixes**: `base64:`, `url:` (requires `requests`), `file:`
- **Error tolerance**: Failed image loads log warnings but don't crash

## Development Conventions

### Code Organization
- **Single module**: All functionality in `main.py` (915 lines)
- **Helper functions**: Prefixed utilities (`hex_to_rgb`, `map_align`, etc.)
- **Style application**: Separate `apply_run_style()`, `apply_paragraph_style()` functions

### Error Handling Patterns
```python
# Graceful degradation - log but continue
try:
    content = get_image_bytes(source, logger)
    if content:
        # proceed with image
except Exception as e:
    logger and logger.warning(f"image failed: {e}")
    # continue without image
```

### Style Merging System
Use `merge_styles(base, override)` for cascading styles from defaults to element-specific overrides.

## Schema & Validation

### JSON Structure
- **Root requirement**: `{"ppt": {"slides": [...]}}` 
- **Validation**: `validate(meta)` checks structure before processing
- **Schema file**: `ppt.schema.json` (570 lines) defines complete structure

### Layout Resolution
```python
# Layout can be string name or integer index
"layout": "blank"          # → prs.slide_layouts[6]
"layout": "titleOnly"      # → prs.slide_layouts[5] 
"layout": 3                # → prs.slide_layouts[3]
```

## Testing & Debugging

### Running the Generator
```powershell
python main.py  # Uses built-in sample data in __main__
```

### Dependencies
- **Required**: `python-pptx` (presentation generation)
- **Optional**: `requests` (for `url:` image sources)
- **Python**: 3.13+ (per pyproject.toml)

### Common Patterns for Extensions
1. **New element type**: Add to `ELEMENT_TYPES` switch in `build()`
2. **Style properties**: Extend `apply_*_style()` functions
3. **Chart types**: Add mappings in `map_chart_type()`
4. **Shape types**: Extend `pick_shape()` enum mapping

## Known Limitations (from LIMITATIONS.md)
- **Elements**: Only text/image/shape/chart/table (no video, SmartArt, audio)
- **Layouts**: Fixed blank layout, no template/master support  
- **Charts**: Limited to bar/line/pie with basic configuration
- **Text**: Single paragraph only, no bullet lists or rich text mixing
- **Backgrounds**: Solid colors or single full-page images only

## File Structure Context
- `main.py` - Complete generator implementation
- `ppt.schema.json` - JSON schema validation
- `README.md` - Chinese documentation with examples
- `LIMITATIONS.md` - Constraint documentation
- `pyproject.toml` - Minimal Python 3.13+ project config