from docx import Document

doc = Document("input/DummyFrench.docx")

for para in doc.paragraphs:

    if para.text == "Un nouveau jour":

        para.runs[0].text = "A New Day"

doc.save("output/DummyFrench_Test.docx")

print("Done")