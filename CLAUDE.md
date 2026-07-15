# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

A growing collection of standalone Python utility scripts. Each utility lives in its own subfolder. Dependencies are managed with uv.

## Structure

uv workspace — każdy util ma własny `pyproject.toml`, root agreguje wszystkich members.

## Running scripts

Z rootu (workspace):
```bash
uv run --package <name> <subfolder>/<script>.py [args]
```

Z folderu utilsa (standalone):
```bash
cd <subfolder>
uv run <script>.py [args]
```

## Syncing

```bash
uv sync --all-packages   # instaluje deps wszystkich members
```

## Adding a new utility

1. Utwórz folder i skrypt
2. Dodaj `pyproject.toml` z własnymi deps
3. Dopisz folder do `members` w root `pyproject.toml`
4. `uv sync --all-packages`

## Adding a dependency to an existing utility

```bash
cd <subfolder>
uv add <package>
```

## Utilities

### `merge_pptx/merge_pptx.py`

Merges multiple PPTX files into one by manipulating the underlying ZIP/XML structure directly (no python-pptx). Handles slide renumbering, media/embedding deduplication, notes slides, relationship files, and content type registration.

**Dependency:** `lxml`

```bash
uv run merge_pptx/merge_pptx.py <folder> [output.pptx]
```

- `<folder>` — directory with `.pptx` files to merge (sorted alphabetically)
- `[output.pptx]` — optional output path, defaults to `merged.pptx`

**Architecture:** Reads all ZIPs into memory. Uses the first file as the base (its `ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels`, `[Content_Types].xml` are mutated). For each additional source: renames media/embeddings with `_sN` suffix, renumbers slides and notes slides sequentially, rewrites `.rels` files, appends `<sldId>` entries, registers new parts in `[Content_Types].xml`. Strips revision/change-tracking metadata from the base. Preserves original ZIP compression flags from the base; uses `ZIP_DEFLATED` level 6 for all new entries.

---

### `chromedriver/download_chromedriver.py`

Downloads ChromeDriver matching a given Chrome version. Supports Chrome ≥115 (Chrome for Testing API) and legacy <115 (storage.googleapis.com). Auto-detects platform (win32/win64/linux64/mac-x64/mac-arm64). No external dependencies.

```bash
uv run chromedriver/download_chromedriver.py <chrome_version> [output_dir]
# Example:
uv run chromedriver/download_chromedriver.py 120.0.6099.109 chromedriver/chromedriver-win64
```

Downloaded binaries land in `chromedriver/chromedriver-*/` which is gitignored.

---

### `chromedriver_pypac/download_chromedriver.py`

Same as `chromedriver/download_chromedriver.py`, but all HTTP requests go through a `pypac.PACSession` so an enterprise/system PAC file (Proxy Auto-Config) is honoured. When no PAC file is discovered, pypac falls back to requests' normal proxy handling (`HTTP_PROXY` / `HTTPS_PROXY`). Same CLI and behaviour otherwise.

**Dependency:** `pypac`

```bash
uv run chromedriver_pypac/download_chromedriver.py <chrome_version> [output_dir]
# Example:
uv run chromedriver_pypac/download_chromedriver.py 120.0.6099.109 chromedriver_pypac/chromedriver-win64
```

Downloaded binaries land in `chromedriver_pypac/chromedriver-*/` which is gitignored.
