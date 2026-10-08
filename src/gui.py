import os
import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import messagebox

from translate_docx import translate_docx
from translate_excel import translate_excel
from translate_ppt import translate_ppt


selected_file = ""


def browse_file():
    global selected_file

    selected_file = filedialog.askopenfilename(
        filetypes=[
            ("Supported Files",
             "*.docx *.xlsx *.pptx")
        ]
    )

    file_label.config(text=selected_file)


def translate_file():

    global selected_file

    if not selected_file:
        messagebox.showerror(
            "Error",
            "Please select a file"
        )
        return

    source_language = language_var.get()

    filename = os.path.basename(selected_file)
    file_base, extension = os.path.splitext(filename)

    output_file = os.path.join(
        "output",
        f"{file_base}_English{extension}"
    )

    try:

        if extension.lower() == ".docx":

            translate_docx(
                selected_file,
                output_file,
                source_language
            )

        elif extension.lower() == ".xlsx":

            translate_excel(
                selected_file,
                output_file,
                source_language
            )

        elif extension.lower() == ".pptx":

            translate_ppt(
                selected_file,
                output_file,
                source_language
            )

        else:

            messagebox.showerror(
                "Error",
                "Unsupported file type"
            )
            return

        messagebox.showinfo(
            "Success",
            f"Translation completed!\n\n{output_file}"
        )

    except Exception as e:

        messagebox.showerror(
            "Error",
            str(e)
        )


root = tk.Tk()
root.title("Translate To English")
root.geometry("600x250")

title = tk.Label(
    root,
    text="Offline Document Translator",
    font=("Arial", 14, "bold")
)

title.pack(pady=10)

browse_btn = tk.Button(
    root,
    text="Browse File",
    command=browse_file
)

browse_btn.pack()

file_label = tk.Label(
    root,
    text="No file selected",
    wraplength=550
)

file_label.pack(pady=10)

language_var = tk.StringVar()
language_var.set("fr")

language_dropdown = ttk.Combobox(
    root,
    textvariable=language_var,
    values=[
        "fr",
        "de",
        "es",
        "pt",
        "ja",
        "zh"
    ]
)

language_dropdown.pack(pady=10)

translate_btn = tk.Button(
    root,
    text="Translate",
    command=translate_file
)

translate_btn.pack(pady=20)

root.mainloop()