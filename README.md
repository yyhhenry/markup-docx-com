# markup-docx-com

Use Markdown or Typst in Word documents

## Usage

```txt
usage: markup-docx.py [-h] [--from FROM_FORMAT]

options:
  -h, --help          show this help message and exit
  --from FROM_FORMAT  typst, markdown_mmd, html (default: typst)
```

## Development

```bash
# Run
uv run markup-docx.py

# Build
uv run -- pyinstaller --onefile markup-docx.py
```
