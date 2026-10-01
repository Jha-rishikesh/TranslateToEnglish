from docx import Document

doc = Document("input/DummyFrench.docx")

for para_no, para in enumerate(doc.paragraphs, start=1):

    if para.text.strip():

        print("\n" + "=" * 50)
        print(f"Paragraph {para_no}")
        print(f"Style: {para.style.name}")

        for run_no, run in enumerate(para.runs, start=1):

            print(f"\nRun {run_no}")
            print("Text :", repr(run.text))
            print("Bold :", run.bold)
            print("Italic :", run.italic)

            if run.font.color.rgb:
                print("Color :", run.font.color.rgb)
            else:
                print("Color : Default")