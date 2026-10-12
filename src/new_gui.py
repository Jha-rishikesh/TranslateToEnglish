import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from PIL import Image, ImageTk, ImageEnhance

from translate_docx import translate_docx
from translate_excel import translate_excel
from translate_ppt import translate_ppt


# =========================================================
# APPLICATION SETTINGS
# =========================================================

WINDOW_WIDTH = 760
WINDOW_HEIGHT = 500

TEAL = "#16CEC1"
LIGHT_TEAL = "#56E1D7"
DARK_TEAL = "#095F5A"
HOVER_TEAL = "#69E8DF"
BLACK = "#111111"
CARD_COLOR = "#172321"
WHITE = "#FFFFFF"
SOFT_WHITE = "#E9FFFF"
GRAY = "#A9B8B6"
YELLOW = "#FFD447"

BACKGROUND_IMAGE = "colt_background.jpg"

selected_file = ""
translate_button = None


# =========================================================
# LANGUAGE DETAILS
# =========================================================

LANGUAGES = {
    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Portuguese": "pt",
    "Japanese": "ja",
    "Chinese": "zh"
}


# =========================================================
# WINDOW MOVEMENT
# =========================================================

def start_move(event):
    root.drag_start_x = event.x_root
    root.drag_start_y = event.y_root


def move_window(event):
    new_x = root.winfo_x() + event.x_root - root.drag_start_x
    new_y = root.winfo_y() + event.y_root - root.drag_start_y

    root.geometry(f"+{new_x}+{new_y}")


    root.drag_start_x = event.x_root
    root.drag_start_y = event.y_root


def minimize_window():
    root.overrideredirect(False)
    root.iconify()


def restore_custom_title_bar(event=None):
    root.overrideredirect(True)


def close_application():
    root.destroy()


# =========================================================
# FILE BROWSER
# =========================================================

def browse_file():
    global selected_file

    new_file = filedialog.askopenfilename(
        title="Select a document",
        filetypes=[
            ("Supported Documents", "*.docx *.xlsx *.pptx"),
            ("Word Documents", "*.docx"),
            ("Excel Workbooks", "*.xlsx"),
            ("PowerPoint Presentations", "*.pptx")
        ]
    )

    if new_file:
        selected_file = new_file
        filename = os.path.basename(selected_file)

        file_name_label.config(
            text=filename,
            fg=WHITE
        )

        selection_status_label.config(
            text="FILE READY",
            fg=LIGHT_TEAL
        )


# =========================================================
# TRANSLATION
# =========================================================

def translate_file():
    global selected_file

    if not selected_file:
        messagebox.showerror(
            "No File Selected",
            "Please select a DOCX, XLSX, or PPTX file."
        )
        return

    selected_language = language_var.get()

    if selected_language not in LANGUAGES:
        messagebox.showerror(
            "Language Required",
            "Please select the source language."
        )
        return

    source_language = LANGUAGES[selected_language]

    filename = os.path.basename(selected_file)
    file_base, extension = os.path.splitext(filename)

    os.makedirs("output", exist_ok=True)

    output_file = os.path.join(
        "output",
        f"{file_base}_English{extension}"
    )

    translate_button.set_text("TRANSLATING...")
    translate_button.set_enabled(False)

    selection_status_label.config(
        text="TRANSLATION IN PROGRESS",
        fg=YELLOW
    )

    root.update_idletasks()

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
                "Unsupported File",
                "Only DOCX, XLSX, and PPTX files are supported."
            )
            return

        selection_status_label.config(
            text="TRANSLATION COMPLETED",
            fg=LIGHT_TEAL
        )

        messagebox.showinfo(
            "Translation Completed",
            f"Your file has been translated successfully.\n\n"
            f"Output file:\n{os.path.basename(output_file)}"
        )

    except Exception as error:
        selection_status_label.config(
            text="TRANSLATION FAILED",
            fg="#FF7777"
        )

        messagebox.showerror(
            "Translation Error",
            str(error)
        )

    finally:
        translate_button.set_text("START TRANSLATION")
        translate_button.set_enabled(True)


# =========================================================
# ROUNDED BUTTON CLASS
# =========================================================

class RoundedButton(tk.Canvas):

    def __init__(
        self,
        parent,
        text,
        command,
        width=210,
        height=48,
        radius=22,
        bg_color=TEAL,
        hover_color=HOVER_TEAL,
        text_color=BLACK,
        font=("Segoe UI", 10, "bold")
    ):
        parent_background = parent.cget("bg")

        super().__init__(
            parent,
            width=width,
            height=height,
            bg=parent_background,
            highlightthickness=0,
            bd=0,
            cursor="hand2"
        )

        self.button_width = width
        self.button_height = height
        self.radius = radius
        self.normal_color = bg_color
        self.hover_color = hover_color
        self.text_color = text_color
        self.button_font = font
        self.command = command
        self.enabled = True
        self.button_text = text

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("<Button-1>", self.on_click)

        self.draw_button(self.normal_color)

    def rounded_rectangle(self, x1, y1, x2, y2, radius, **kwargs):
        points = [
            x1 + radius, y1,
            x2 - radius, y1,
            x2, y1,
            x2, y1 + radius,
            x2, y2 - radius,
            x2, y2,
            x2 - radius, y2,
            x1 + radius, y2,
            x1, y2,
            x1, y2 - radius,
            x1, y1 + radius,
            x1, y1
        ]

        return self.create_polygon(
            points,
            smooth=True,
            **kwargs
        )

    def draw_button(self, color):
        self.delete("all")

        self.rounded_rectangle(
            2,
            2,
            self.button_width - 2,
            self.button_height - 2,
            self.radius,
            fill=color,
            outline=""
        )

        self.create_text(
            self.button_width / 2,
            self.button_height / 2,
            text=self.button_text,
            fill=self.text_color,
            font=self.button_font
        )

    def on_enter(self, event):
        if self.enabled:
            self.draw_button(self.hover_color)

    def on_leave(self, event):
        if self.enabled:
            self.draw_button(self.normal_color)

    def on_click(self, event):
        if self.enabled and self.command:
            self.command()

    def set_text(self, new_text):
        self.button_text = new_text
        self.draw_button(self.normal_color)

    def set_enabled(self, enabled):
        self.enabled = enabled

        if enabled:
            self.config(cursor="hand2")
            self.draw_button(self.normal_color)
        else:
            self.config(cursor="arrow")
            self.draw_button("#7A9996")


# =========================================================
# ROOT WINDOW
# =========================================================

root = tk.Tk()

root.title("INFRA DOCUMENT TRANSLATOR")
root.configure(bg=BLACK)

# Remove normal Windows title bar
root.overrideredirect(True)

# Fixed application size
root.resizable(False, False)
root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
root.maxsize(WINDOW_WIDTH, WINDOW_HEIGHT)

# Open window in screen centre
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

window_x = int((screen_width - WINDOW_WIDTH) / 2)
window_y = int((screen_height - WINDOW_HEIGHT) / 2)

root.geometry(
    f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{window_x}+{window_y}"
)

root.bind("<Map>", restore_custom_title_bar)


# =========================================================
# BACKGROUND IMAGE
# =========================================================


outer_border = tk.Frame(
    root,
    bg="#39E6DB"
)

outer_border.place(
    x=1,
    y=1,
    width=758,
    height=498
)

try:
    background_image = Image.open(BACKGROUND_IMAGE)
    background_image = background_image.resize(
        (WINDOW_WIDTH, WINDOW_HEIGHT),
        Image.Resampling.LANCZOS
    )

    # Image ko thoda dark karna so text clearly visible rahe
    background_image = ImageEnhance.Brightness(
        background_image
    ).enhance(0.72)

    background_photo = ImageTk.PhotoImage(background_image)

    background_label = tk.Label(
        root,
        image=background_photo,
        bd=0
    )

    background_label.place(
        x=0,
        y=0,
        relwidth=1,
        relheight=1
    )

except Exception:
    background_label = tk.Label(
        root,
        bg=BLACK
    )

    background_label.place(
        x=0,
        y=0,
        relwidth=1,
        relheight=1
    )


# =========================================================
# CUSTOM TOP BAR
# =========================================================

top_bar = tk.Frame(
    root,
    bg="#0D6862",
    height=44
)

top_bar.place(
    x=0,
    y=0,
    width=WINDOW_WIDTH,
    height=44
)

top_bar.pack_propagate(False)

top_accent = tk.Frame(
    top_bar,
    bg="#2FF0E1",
    width=12
)

top_accent.pack(
    side="left",
    fill="y"
)

window_title = tk.Label(
    top_bar,
    text="INFRA DOCUMENT TRANSLATOR",
    bg=DARK_TEAL,
    fg=WHITE,
    font=("Segoe UI", 10, "bold")
)

window_title.pack(
    side="left",
    padx=14
)

close_button = tk.Label(
    top_bar,
    text="✕",
    bg=DARK_TEAL,
    fg=WHITE,
    font=("Segoe UI", 13),
    cursor="hand2",
    width=4
)

close_button.pack(
    side="right",
    fill="y"
)

minimize_button = tk.Label(
    top_bar,
    text="—",
    bg=DARK_TEAL,
    fg=WHITE,
    font=("Segoe UI", 13),
    cursor="hand2",
    width=4
)

minimize_button.pack(
    side="right",
    fill="y"
)

close_button.bind(
    "<Button-1>",
    lambda event: close_application()
)

minimize_button.bind(
    "<Button-1>",
    lambda event: minimize_window()
)

close_button.bind(
    "<Enter>",
    lambda event: close_button.config(bg="#C42B1C")
)

close_button.bind(
    "<Leave>",
    lambda event: close_button.config(bg=DARK_TEAL)
)

minimize_button.bind(
    "<Enter>",
    lambda event: minimize_button.config(bg="#117A73")
)

minimize_button.bind(
    "<Leave>",
    lambda event: minimize_button.config(bg=DARK_TEAL)
)

for draggable_widget in [
    top_bar,
    window_title,
    top_accent
]:
    draggable_widget.bind("<Button-1>", start_move)
    draggable_widget.bind("<B1-Motion>", move_window)


# =========================================================
# LEFT-SIDE HEADER
# =========================================================

heading_label = tk.Label(
    root,
    text="INFRA DOCUMENT\nTRANSLATOR",
    justify="left",
    anchor="w",
    bg=BLACK,
    fg=WHITE,
    font=("Segoe UI", 24, "bold")
)

heading_label.place(
    x=48,
    y=68
)

heading_accent = tk.Frame(
    root,
    bg=TEAL
)

heading_accent.place(
    x=48,
    y=150,
    width=130,
    height=5
)

subtitle_label = tk.Label(
    root,
    text="Offline document translation tool",
    justify="left",
    anchor="w",
    bg=BLACK,
    fg=LIGHT_TEAL,
    font=("Segoe UI", 9)
)

subtitle_label.place(
    x=48,
    y=155
)


# =========================================================
# COLT TEXT ON RIGHT SIDE
# =========================================================

colt_label = tk.Label(
    root,
    text="colt",
    bg=BLACK,
    fg=WHITE,
    font=("Segoe UI", 30, "bold")
)

colt_label.place(
    x=625,
    y=75
)


# =========================================================
# MAIN CONTENT CARD
# =========================================================


shadow_card = tk.Frame(
    root,
    bg="#082E2C"
)

shadow_card.place(
    x=50,
    y=207,
    width=664,
    height=235
)

main_card = tk.Frame(
    root,
    bg=CARD_COLOR,
    highlightbackground="#3BE8DC",
    highlightcolor="#3BE8DC",
    highlightthickness=2
)

main_card.place(
    x=48,
    y=205,
    width=664,
    height=235
)


# =========================================================
# FILE SELECTION SECTION
# =========================================================

select_label = tk.Label(
    main_card,
    text="SELECT DOCUMENT",
    bg=CARD_COLOR,
    fg=LIGHT_TEAL,
    font=("Segoe UI", 9, "bold")
)

select_label.place(
    x=24,
    y=18
)

browse_button = RoundedButton(
    main_card,
    text="SELECT FILE",
    command=browse_file,
    width=185,
    height=44,
    radius=25,
    bg_color=LIGHT_TEAL,
    hover_color=WHITE,
    text_color=BLACK
)

browse_button.place(
    x=22,
    y=47
)

selection_status_label = tk.Label(
    main_card,
    text="NO FILE SELECTED",
    bg=CARD_COLOR,
    fg=GRAY,
    font=("Segoe UI", 8, "bold")
)

selection_status_label.place(
    x=24,
    y=104
)

file_name_label = tk.Label(
    main_card,
    text="Choose a file to begin",
    bg=CARD_COLOR,
    fg=GRAY,
    anchor="w",
    justify="left",
    font=("Segoe UI", 9),
    wraplength=275
)

file_name_label.place(
    x=24,
    y=128,
    width=280
)


# =========================================================
# DIVIDER
# =========================================================

divider = tk.Frame(
    main_card,
    bg="#40615E"
)

divider.place(
    x=324,
    y=22,
    width=1,
    height=185
)


# =========================================================
# LANGUAGE SELECTION SECTION
# =========================================================

language_label = tk.Label(
    main_card,
    text="SOURCE LANGUAGE",
    bg=CARD_COLOR,
    fg=LIGHT_TEAL,
    font=("Segoe UI", 9, "bold")
)

language_label.place(
    x=356,
    y=18
)

language_var = tk.StringVar()
language_var.set("Select Language")


# =========================================================
# COMBOBOX STYLE
# =========================================================

style = ttk.Style()
style.theme_use("clam")

style.configure(
    "Language.TCombobox",
    fieldbackground=SOFT_WHITE,
    background=TEAL,
    foreground=BLACK,
    arrowcolor=BLACK,
    bordercolor=TEAL,
    lightcolor=TEAL,
    darkcolor=TEAL,
    selectbackground=SOFT_WHITE,
    selectforeground=BLACK,
    padding=8,
    font=("Segoe UI", 10, "bold")
)

style.map(
    "Language.TCombobox",
    fieldbackground=[
        ("readonly", SOFT_WHITE)
    ],
    foreground=[
        ("readonly", BLACK)
    ],
    background=[
        ("readonly", TEAL)
    ],
    bordercolor=[
        ("focus", LIGHT_TEAL),
        ("readonly", TEAL)
    ]
)

language_dropdown = ttk.Combobox(
    main_card,
    textvariable=language_var,
    values=list(LANGUAGES.keys()),
    state="readonly",
    width=23,
    style="Language.TCombobox",
    font=("Segoe UI", 10)
)

language_dropdown.place(
    x=354,
    y=47,
    width=270,
    height=42
)


# Dropdown list colours
root.option_add(
    "*TCombobox*Listbox.background",
    SOFT_WHITE
)

root.option_add(
    "*TCombobox*Listbox.foreground",
    BLACK
)

root.option_add(
    "*TCombobox*Listbox.selectBackground",
    TEAL
)

root.option_add(
    "*TCombobox*Listbox.selectForeground",
    BLACK
)

translate_button = RoundedButton(
    main_card,
    text="START TRANSLATION",
    command=translate_file
)

translate_button.place(
    x=354,
    y=112
)

root.mainloop()