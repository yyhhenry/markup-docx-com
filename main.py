import argparse
import os
import tempfile
from typing import Callable, Literal

import keyboard
import pyperclip
import win32com.client
from pydantic import BaseModel
from win32com.client.dynamic import CDispatch


class Args(BaseModel):
    from_format: Literal["typst", "markdown_mmd", "html"]


arg_parser = argparse.ArgumentParser()
arg_parser.add_argument(
    "--from",
    dest="from_format",
    default="typst",
    help="typst, markdown_mmd, html (default: typst)",
)
args_namespace = arg_parser.parse_args()
args = Args.model_validate(args_namespace.__dict__)

print(f"Auto-converting from {args.from_format} to docx")


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

        print(f"{text=}")
        return text, inline_block


def get_clipboard_text() -> tuple[str, bool] | None:
    text = pyperclip.paste()
    if text.strip() == "":
        print("No text in clipboard")
        return None
    else:
        inline_block = text.strip().find("\n") == -1
        print(f"{text=}")
        return text, inline_block


def call_pandoc(
    input_file: str,
    output_file: str,
    from_format: str,
    to_format: str,
):
    if not is_pandoc_in_path():
        raise Exception("Pandoc not found in PATH")

    command = f"pandoc -f {from_format} -t {to_format} {input_file} -o {output_file}"
    result = os.system(command)
    if result != 0:
        print(f"$ {command}")
        raise Exception(f"Failed to convert {input_file} to {output_file}")


def convert_to_docx(text: str, insert_into_word: Callable[[str], None]):
    print(f"Converting {args.from_format} to docx...")
    if not is_pandoc_in_path():
        raise Exception("Pandoc not found in PATH")

    with tempfile.TemporaryDirectory() as temp_dir:
        ext = ext_from_format(args.from_format)
        source_code = os.path.join(temp_dir, "source." + ext)
        with open(source_code, "w", encoding="utf-8") as f:
            f.write(text)

        docx_file = os.path.join(temp_dir, "temp.docx")
        call_pandoc(source_code, docx_file, args.from_format, "docx")
        insert_into_word(docx_file)


def insert_into_docx(word: CDispatch, docx_file: str, inline_block: bool):
    selection = word.Selection
    if inline_block:
        # Get the style of the current selection
        style = selection.Style()
        print(f"{style=}")

    selection.InsertFile(docx_file)
    print(f"{selection.Text=}")

    # Remove additional line break at the end of the inserted text
    if inline_block:
        selection.MoveLeft()
        text = selection.Text
        assert text == "\r", f"{text=}"
        selection.Delete()
        selection.Style = style


def on_triggered():
    from pythoncom import CoInitialize

    CoInitialize()
    # Connect to the Word application
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True
    doc = word.ActiveDocument
    print(f"\nActive document: {doc}")

    text, inline_block = (
        get_selection_text(word) or get_clipboard_text() or (None, False)
    )

    if text is None:
        print("No text to convert")
        return

    def insert_into_word(docx_file: str):
        # Copy to Download folder
        from shutil import copy

        download_folder = os.path.expanduser("~/Downloads")
        copy(docx_file, download_folder)

        insert_into_docx(word, docx_file, inline_block)

    convert_to_docx(text, insert_into_word)
    print("Done.")


keyboard.add_hotkey("ctrl+shift+t", on_triggered)

print("Press Ctrl+Shift+T to convert selected text to docx")

keyboard.wait()
