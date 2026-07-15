# utils

Kolekcja standalone skryptów Python do codziennych zadań. Każdy util ma własny folder z `pyproject.toml`. Środowiskiem zarządza [uv](https://docs.astral.sh/uv/).

```bash
# uruchomienie z rootu
uv run --package <name> <folder>/<script>.py [args]

# lub standalone
cd <folder> && uv run <script>.py [args]
```

---

## merge_pptx

Scala wiele plików PPTX w jeden, operując bezpośrednio na strukturze ZIP/XML — bez zewnętrznych bibliotek do obsługi PPTX. Zachowuje slajdy, notatki, media, osadzenia i relacje.

**Zależności:** `lxml`

```bash
uv run --package merge-pptx merge_pptx/merge_pptx.py <folder> [output.pptx]
```

| Argument | Opis |
|---|---|
| `<folder>` | Folder z plikami `.pptx` (scalane alfabetycznie) |
| `[output.pptx]` | Opcjonalna ścieżka wyjściowa, domyślnie `merged.pptx` |

---

## chromedriver

Pobiera ChromeDriver dopasowany do podanej wersji Chrome. Obsługuje Chrome ≥ 115 (Chrome for Testing API) oraz starsze wersje (legacy endpoint). Platforma wykrywana automatycznie.

**Zależności:** brak (tylko stdlib)

```bash
uv run --package chromedriver-utils chromedriver/download_chromedriver.py <chrome_version> [output_dir]
```

| Argument | Opis |
|---|---|
| `<chrome_version>` | Wersja Chrome, np. `136.0.7103.93` |
| `[output_dir]` | Folder docelowy, domyślnie `.` |

Pliki binarne ChromeDrivera trafiają bezpośrednio do `output_dir`.

---

## chromedriver_pypac

To samo co `chromedriver`, z jedną różnicą: wszystkie żądania HTTP idą przez `pypac.PACSession`, więc respektowany jest systemowy/firmowy plik PAC (Proxy Auto-Config). Gdy PAC nie zostanie znaleziony, pypac wraca do standardowej obsługi proxy z `requests` (`HTTP_PROXY` / `HTTPS_PROXY`).

**Zależności:** `pypac`

```bash
uv run --package chromedriver-pypac-utils chromedriver_pypac/download_chromedriver.py <chrome_version> [output_dir]
```

| Argument | Opis |
|---|---|
| `<chrome_version>` | Wersja Chrome, np. `136.0.7103.93` |
| `[output_dir]` | Folder docelowy, domyślnie `.` |

Pliki binarne ChromeDrivera trafiają bezpośrednio do `output_dir`.
