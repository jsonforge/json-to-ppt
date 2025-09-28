# Template Packaging Specification (Draft)

## Purpose
Define a portable, signed bundle format for distributing premium or community templates.

## Bundle Structure
```
my-template-pack/
  manifest.json
  templates/
    cover.json
    kpi_dashboard.json
  assets/
    logo.png
    palette.json
  README.md (optional)
  SIGNATURE.txt
```

## manifest.json
```
{
  "packId": "tpl_pack_finance_v1",
  "version": "1.0.0",
  "name": "Finance Essentials",
  "author": "Acme Studio",
  "license": "Commercial-Use-Limited",
  "createdAt": "2025-09-29T12:00:00Z",
  "templates": [
    {"id": "cover", "file": "templates/cover.json", "tags": ["cover", "brand"], "slides": 1},
    {"id": "kpi_dashboard", "file": "templates/kpi_dashboard.json", "tags": ["kpi","metrics"], "slides": 3}
  ],
  "assets": [
    {"type": "image", "file": "assets/logo.png"},
    {"type": "palette", "file": "assets/palette.json"}
  ],
  "requiresCoreVersion": ">=0.1.0",
  "hashAlgo": "sha256"
}
```

## Signature
- Concatenate canonical JSON of `manifest.json` (no whitespace) + sorted file hashes.
- Produce a detached signature stored in `SIGNATURE.txt`.
- Optionally include public key reference: `key:fpr:ABCD1234`.

## Integrity Workflow
1. User uploads `.zip` to marketplace.
2. Server unpacks -> computes file hashes -> verifies signature -> stores metadata.
3. At install/publish time a second integrity check occurs.

## Hash File Example
```
sha256  templates/cover.json  5e3d3c...
sha256  templates/kpi_dashboard.json  9a1031...
sha256  assets/logo.png  33ffab...
```

## Distribution Format
- Distributed as `.zip` archive.
- Maximum unpacked size (free tier): 5 MB.
- Reject archives with symlinks or paths containing `..` traversal.

## Runtime Use
- Loader verifies `requiresCoreVersion` to avoid incompatible fields.
- Merges palette/theme into user JSON payload if requested.

## Future
- Encrypted asset section for DRM-lite (optional).
- Differential updates (delta packs).
- Template dependency graph (inherit base layout pack).
