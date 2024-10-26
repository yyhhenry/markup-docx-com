# markup-docx-com

Use Markdown or Typst in Word documents

## Usage

```txt
usage: markup-docx.py [-h] [--from FROM_FORMAT] [--word-title WORD_TITLE] [--force-straight-quotes]

options:
  -h, --help            show this help message and exit
  --from FROM_FORMAT    typst, markdown_mmd, html (default: typst)
  --word-title WORD_TITLE
                        The title of the Word window (default: {doc} - Word)
  --force-straight-quotes
                        Replace curly quotes with straight quotes
```

提示 (zh-CN):

- 当 Word 在前台时，按 Ctrl+# 可以将选中内容作为标记语言并替换为编译结果
- 确保 pandoc 已安装并在 PATH 中
- 打开 选项-校对-自动更正选项，检查不适合代码的自动更正
  - 在自动套用格式和键入时自动套用格式中，关闭“直引号”自动更正
  - 关闭首字母大写自动更正，以方便代码块输入

## Development

```ps1
# Run
uv run markup-docx.py

# Build
uv run -- pyinstaller --onefile markup-docx.py
# Then put the executable in the PATH, e.g.:
# cp dist/markup-docx.exe $HOME/src/bin/markup-docx.exe
```
