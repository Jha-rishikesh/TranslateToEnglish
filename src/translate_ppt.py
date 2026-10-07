import os
os.environ["ARGOS_CHUNK_TYPE"] = "MINISBD"

from pptx import Presentation
import argostranslate.translate

def should_translate(text):

    text = text.strip()

    if not text:
        return False

    if len(text) <= 1:
        return False

    if text.isdigit():
        return False

    # Place names with Japanese markers
    japanese_markers = [
        "Ō", "ō", "chō",
        "shi", "ku",
        "aza", "miya"
    ]

    for marker in japanese_markers:
        if marker in text:
            return False

    return True


input_file = "input/DummyFrench.pptx"
output_file = "output/DummyFrench_English.pptx"

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

ppt = Presentation(input_file)

# for slide in ppt.slides:

#     for shape in slide.shapes:

        # if hasattr(shape, "text"):

        #     if shape.text.strip():

        #         original = shape.text
        #         translated = translator.translate(original)

        #         shape.text = translated

        #         print(original)
        #         print("↓")
        #         print(translated)
        #         print("-" * 40)

        # if hasattr(shape, "text_frame"):
        

        #  for paragraph in shape.text_frame.paragraphs:
        
        #     if not paragraph.text.strip():
        #         continue
        
        #     original = paragraph.text
        #     translated = translator.translate(original)
        
        #     paragraph.runs[0].text = translated
        
        #     for run in paragraph.runs[1:]:
        #         run.text = ""
        
        #     print(original)
        #     print("↓")
        #     print(translated)
        #     print("-" * 40)
        
        # # Translate tables
        # if shape.has_table:
        
        #     table = shape.table
        
        #     for row in table.rows:
        
        #         for cell in row.cells:
        
        #             if cell.text.strip():
        
        #                 translated = translator.translate(
        #                     cell.text
        #                 )
        
        #                 cell.text = translated
        
        #                 print("TABLE:")
        #                 print(cell.text)    

for slide in ppt.slides:

    for shape in slide.shapes:

        # TEXT SHAPES
        if hasattr(shape, "text_frame"):

            for paragraph in shape.text_frame.paragraphs:

                if not paragraph.text.strip():
                    continue

                original = paragraph.text
                if not should_translate(original):
                  continue
                translated = translator.translate(original)

                paragraph.runs[0].text = translated

                for run in paragraph.runs[1:]:
                    run.text = ""

                    print(original)
                    print("↓")
                    print(translated)
                    print("-" * 40)

        # TABLES
        # if shape.has_table:

        #     table = shape.table

        #     for row in table.rows:

        #         for cell in row.cells:

        #             if cell.text.strip():

        #                 translated = translator.translate(
        #                     cell.text
        #                 )

        #                 # cell.text = translated

        #                 print("TABLE:")
        #                 print(cell.text)
        # Translate tables while preserving text formatting
        if shape.has_table:
        
            table = shape.table
        
            for row in table.rows:
        
                for cell in row.cells:
        
                    for paragraph in cell.text_frame.paragraphs:
        
                        if not paragraph.text.strip():
                            continue
        
                        original = paragraph.text
                        if not should_translate(original):
                         continue
                        translated = translator.translate(original)
        
                        if paragraph.runs:
                            paragraph.runs[0].text = translated
        
                            for run in paragraph.runs[1:]:
                                run.text = ""
        
                        print("TABLE:")
                        print(original)
                        print("↓")
                        print(translated)

ppt.save(output_file)

print("Translation Completed!")
