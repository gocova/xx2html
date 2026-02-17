# xx2html
[![PyPI Version](https://img.shields.io/pypi/v/xx2html.svg)](https://pypi.org/project/xx2html/)
[![License](https://img.shields.io/badge/License-MIT%20%2F%20Apache%202.0-green.svg)](https://opensource.org/licenses/)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-Support-orange?logo=buy-me-a-coffee&style=flat-square)](https://buymeacoffee.com/gocova)

`xx2html` converts Excel workbooks (`.xlsx`) into a single HTML file while preserving:

- Cell formatting and styles
- Conditional formatting classes (via `condif2css`)
- Worksheet link behavior
- Embedded worksheet images and in-cell rich-value images

## Installation

```bash
pip install xx2html
```

## Usage

```python
from xx2html import apply_openpyxl_patches, create_xlsx_transform

# Explicit entrypoint. Patches are also applied automatically on import.
apply_openpyxl_patches()

transform = create_xlsx_transform(
    sheet_html=(
        '<section id="{enc_sheet_name}" data-sheet-name="{sheet_name}">'
        "{table_generated_html}"
        "</section>"
    ),
    sheetname_html='<a class="sheet-nav" href="#{sheet_name}.A1">{sheet_name}</a>',
    index_html=(
        "<!doctype html><html><head>"
        "{fonts_html}{core_css_html}{user_css_html}{generated_css_html}"
        "{generated_incell_css_html}{conditional_css_html}"
        "</head><body>{sheets_names_generated_html}{sheets_generated_html}"
        "{safari_js}</body></html>"
    ),
    fonts_html="",
    core_css="",
    user_css="",
    safari_js="",
    apply_cf=True,
)

ok, err = transform("input.xlsx", "output.html", "en_US")
if not ok:
    raise RuntimeError(err)
```

## Template Placeholders

`sheet_html` requires:
- `{enc_sheet_name}`
- `{sheet_name}`
- `{table_generated_html}`

`sheetname_html` requires:
- `{enc_sheet_name}`
- `{sheet_name}`

`index_html` requires:
- `{sheets_generated_html}`
- `{sheets_names_generated_html}`
- `{source_filename}`
- `{fonts_html}`
- `{core_css_html}`
- `{user_css_html}`
- `{generated_css_html}`
- `{generated_incell_css_html}`
- `{safari_js}`
- `{conditional_css_html}`

## Monkey Patching Behavior

`xx2html` relies on an `openpyxl` monkey patch to carry rich-value metadata used for in-cell images.

- The patch is applied automatically when `xx2html.core` is imported.
- The explicit API entrypoint is `apply_openpyxl_patches()`.

## Development

```bash
uv sync --group dev
python3 tests/scripts/generate_fixtures.py
python3 -m compileall src tests
python3 -m pytest
```

## License

`xx2html` is dual-licensed under MIT or Apache-2.0.
