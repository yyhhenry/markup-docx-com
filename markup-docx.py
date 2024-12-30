import argparse
import os
import subprocess
import tempfile
from typing import Literal

import keyboard
import pyperclip
import win32com.client
from pydantic import BaseModel
from pythoncom import CoInitialize
from win32api import MessageBox
from win32com.client.dynamic import CDispatch
from win32con import MB_ICONERROR
from win32gui import GetForegroundWindow, GetWindowText


class Args(BaseModel):
    from_format: Literal["typst", "md", "markdown_mmd", "html"]
    wps: bool
    title: str | None
    force_straight_quotes: bool


arg_parser = argparse.ArgumentParser()
arg_parser.add_argument(
    "-f",
    "--from-format",
    dest="from_format",
    default="typst",
    help="typst, md (markdown_mmd), html (default: typst)",
)
arg_parser.add_argument(
    "--wps",
    dest="wps",
    action="store_true",
    help="Use WPS Office instead of Word",
)
arg_parser.add_argument(
    "--title",
    dest="title",
    default=None,
    help="The title of the Word window (default: `{doc} - Word` or `{doc} - WPS Office` if --wps is set)",
)
arg_parser.add_argument(
    "--force-straight-quotes",
    dest="force_straight_quotes",
    action="store_true",
    help="Replace curly quotes with straight quotes",
)
args_namespace = arg_parser.parse_args()
args = Args.model_validate(args_namespace.__dict__)

print(f"Auto-converting from {args.from_format} to docx")


def get_app_name(wps: bool) -> str:
    return "WPS Office" if wps else "Word"


app_name = get_app_name(args.wps)
if args.title is None:
    args.title = f"{{doc}} - {app_name}"

if args.from_format == "md":
    args.from_format = "markdown_mmd"


def ext_from_format(format: str) -> str:
    return {
        "typst": "typ",
        "markdown_mmd": "md",
        "html": "html",
    }[format]


def is_pandoc_in_path():
    path = os.environ.get("PATH")
    if path is None:
        return False

    for dir in path.split(os.pathsep):
        pandoc_path = os.path.join(dir, "pandoc.exe")
        if os.path.isfile(pandoc_path):
            return True
    return False


def get_selection_text(word: CDispatch) -> tuple[str, bool] | None:
    selection = word.Selection
    text = str(selection.Text)
    if text.strip() == "":
        print("No text selected")
        return None
    else:
        inline_block = text.strip().find("\r") == -1
        if inline_block and selection.End == selection.Paragraphs.Last.Range.End:
            # Don't include the last line break
            selection.End = selection.End - 1
        return text, inline_block


def get_clipboard_text() -> tuple[str, bool] | None:
    text = pyperclip.paste()
    if text.strip() == "":
        print("No text in clipboard")
        return None
    else:
        inline_block = text.strip().find("\n") == -1
        return text, inline_block


def text_filter(text: str) -> str:
    # Normalize line endings
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Shift+Enter in Word, text of soft line break is '\x0b' (Vertical Tab in ASCII)
    text = text.replace("\x0b", "\n")

    if args.force_straight_quotes:
        text = text.replace("“", '"').replace("”", '"')
        text = text.replace("‘", "'").replace("’", "'")

    return text


def call_pandoc(
    input_file: str,
    output_file: str,
    from_format: str,
    to_format: str,
):
    if not is_pandoc_in_path():
        raise Exception("Pandoc not found in PATH")

    result = subprocess.run(
        ["pandoc", "-f", from_format, "-t", to_format, input_file, "-o", output_file],
        capture_output=True,
        encoding="utf-8",
    )

    if result.returncode != 0:
        message = result.stdout + result.stderr
        raise Exception(message)


def convert_to_docx(text: str, temp_dir: str):
    print(f"Converting {args.from_format} to docx...")

    ext = ext_from_format(args.from_format)
    source_code = os.path.join(temp_dir, "source." + ext)
    with open(source_code, "w", encoding="utf-8") as f:
        f.write(text)

    docx_file = os.path.join(temp_dir, "temp.docx")
    call_pandoc(source_code, docx_file, args.from_format, "docx")
    return docx_file


def insert_into_docx(word: CDispatch, docx_file: str, inline_block: bool):
    selection = word.Selection
    style = None
    if inline_block:
        # Get the style of the current selection
        style = selection.Style()
        print(f"{style=}")

    selection.InsertFile(docx_file)

    # Remove additional line break at the end of the inserted text
    if inline_block:
        selection.MoveLeft()
        text = selection.Text
        assert text == "\r", f"Expected '\\r', got '{text}'"
        selection.Delete()
        selection.Style = style


def connect_to_word(hwnd: int) -> CDispatch:
    CoInitialize()
    try:
        word = win32com.client.GetObject(None, "Word.Application")
        doc = word.ActiveDocument
        print(f"\nActive document: {doc}")
    except Exception:
        raise Exception("Please open a Word document.")

    title = GetWindowText(hwnd)
    assert args.title is not None, "Title is not set"
    expected_title = args.title.format(doc=doc.Name)
    if title != expected_title:
        raise Exception(f"Foreground window is not Word. {title=}, {expected_title=}")
    return word


def on_triggered():
    hwnd = GetForegroundWindow()
    try:
        word = connect_to_word(hwnd)
    except Exception as e:
        message = str(e)
        MessageBox(hwnd, message, "Error", MB_ICONERROR)
        print(f"\n{message}")
        return

    text, inline_block = (
        get_selection_text(word) or get_clipboard_text() or (None, False)
    )

    if text is None:
        print("No text to convert")
        return

    text = text_filter(text)

    print("Code:")
    for idx, line in enumerate(text.strip().split("\n")):
        print(f"{idx:3}| {line}")

    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            docx_file = convert_to_docx(text, temp_dir)
            insert_into_docx(word, docx_file, inline_block)

        print("Done.")
    except Exception as e:
        message = str(e)
        MessageBox(hwnd, message, "Error", MB_ICONERROR)
        print(message)


keyboard.add_hotkey("ctrl+shift+3", on_triggered)

print("Press Ctrl+# (Ctrl+Shift+3) to convert selected text to docx")


print("\n提示 (zh-CN):")
print(f"- 当 {app_name} 在前台时，按 Ctrl+# 可以将选中内容作为标记语言并替换为编译结果")
print(f"- 请勿同时打开 {get_app_name(not args.wps)}，否则可能会导致错误")
print("- 确保 pandoc 已安装并在 PATH 中")
print("- 打开 选项-校对-自动更正选项，检查不适合代码的自动更正")
print("  - 在自动套用格式和键入时自动套用格式中，关闭“直引号”自动更正")
print("  - 关闭首字母大写自动更正，以方便代码块输入")


keyboard.wait()
