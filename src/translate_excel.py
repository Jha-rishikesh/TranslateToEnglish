# from openpyxl import load_workbook

# input_file = "input/DummyFrench.xlsx"
# output_file = "output/DummyFrench_English.xlsx"

# workbook = load_workbook(input_file)

# for sheet in workbook.worksheets:

#     for row in sheet.iter_rows():

#         for cell in row:

#             if isinstance(cell.value, str):

#                 print(cell.coordinate, "=", cell.value)

# workbook.save(output_file)

# print("Excel processed successfully")

import os
os.environ["ARGOS_CHUNK_TYPE"] = "MINISBD"

from openpyxl import load_workbook
import argostranslate.translate

input_file = "input/DummyFrench.xlsx"
output_file = "output/DummyFrench_English.xlsx"

source_language = "fr"
target_language = "en"

installed_languages = argostranslate.translate.get_installed_languages()

source = next(
    lang for lang in installed_languages
    if lang.code == source_language
)

target = next(
    lang for lang in installed_languages
    if lang.code == target_language
)

translator = source.get_translation(target)

workbook = load_workbook(input_file)

for sheet in workbook.worksheets:

    print(f"\nSheet: {sheet.title}")

    for row in sheet.iter_rows():

        for cell in row:

            if isinstance(cell.value, str):

                original = cell.value
                translated = translator.translate(original)

                cell.value = translated

                print(
                    f"{cell.coordinate}: "
                    f"{original} -> {translated}"
                )

workbook.save(output_file)

print("\nTranslation Completed!")
print(f"Output File: {output_file}")