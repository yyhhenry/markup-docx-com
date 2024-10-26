# markup-docx-com

Use Markdown or Typst in Word documents

## Usage

```txt
usage: markup-docx.py [-h] [--from FROM_FORMAT] [--word-title WORD_TITLE]

options:
  -h, --help            show this help message and exit
  --from FROM_FORMAT    typst, markdown_mmd, html (default: typst)
  --word-title WORD_TITLE
                        The title of the Word window (default: {doc} - Word)
```

## Development

```ps1
# Run
uv run markup-docx.py

# Build
uv run -- pyinstaller --onefile markup-docx.py
# Then put the executable in the PATH, e.g.:
# cp dist/markup-docx.exe $HOME/src/bin/markup-docx.exe
```
