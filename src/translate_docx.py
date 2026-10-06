import os
os.environ["ARGOS_CHUNK_TYPE"] = "MINISBD"

from docx import Document
import argostranslate.translate

input_file = "input/DummyFrench.docx"
output_file = "output/DummyFrench_English_V2.docx"

installed_languages = argostranslate.translate.get_installed_languages()

french = next(lang for lang in installed_languages if lang.code == "fr")
english = next(lang for lang in installed_languages if lang.code == "en")

translator = french.get_translation(english)

doc = Document(input_file)

# for paragraph in doc.paragraphs:
#     for run in paragraph.runs:

#         if run.text.strip():
#             original_text = run.text
#             translated_text = translator.translate(original_text)

#             run.text = translated_text

#             print("French :", original_text)
#             print("English:", translated_text)
#             print("-" * 40)

for paragraph in doc.paragraphs:

    original_text = paragraph.text

    if not original_text.strip():
        continue

    translated_text = translator.translate(original_text)

    # Translation ko first run me rakho
    paragraph.runs[0].text = translated_text

    # Baaki runs ka French text remove karo
    for run in paragraph.runs[1:]:
        run.text = ""

    print("French :", original_text)
    print("English:", translated_text)
    print("-" * 40)

    ###################
for table in doc.tables:
   
      for row in table.rows:
   
       for cell in row.cells:
   
          if cell.text.strip():
   
           translated = translator.translate(cell.text)
   
           cell.text = translated
   
           print("Table:", translated)


doc.save(output_file)

print("\nTranslation completed successfully")
print("Output:", output_file)
########################
