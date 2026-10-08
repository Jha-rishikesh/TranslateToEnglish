import os

from translate_docx import translate_docx
from translate_excel import translate_excel
from translate_ppt import translate_ppt


file_name = input(
    "Enter file name with extension: "
).strip()

source_language = input(
    "Enter source language "
    "(fr/de/es/pt/ja/zh): "
).strip().lower()

input_file = os.path.join(
    "input",
    file_name
)

if not os.path.exists(input_file):
    print("File not found:", input_file)
    raise SystemExit

file_base, extension = os.path.splitext(file_name)
extension = extension.lower()

output_file = os.path.join(
    "output",
    f"{file_base}_English{extension}"
)

try:
    if extension == ".docx":
        print("Calling DOCX Translator...")

        translate_docx(
            input_file,
            output_file,
            source_language
        )

    elif extension == ".xlsx":
        print("Calling Excel Translator...")

        translate_excel(
            input_file,
            output_file,
            source_language
        )

    elif extension == ".pptx":
        print("Calling PPT Translator...")

        translate_ppt(
            input_file,
            output_file,
            source_language
        )

    else:
        print("Unsupported file type:", extension)

except StopIteration:
    print(
        "Required language model is not installed:",
        source_language,
        "-> en"
    )

except PermissionError:
    print(
        "File is open or locked. Close the Word, Excel, "
        "or PowerPoint file and try again."
    )

except Exception as error:
    print("Translation failed:", error)