import os
import sys
import threading
import customtkinter as ctk

from tkinter import filedialog, messagebox
from PIL import Image

from translate_docx import translate_docx
from translate_excel import translate_excel
from translate_ppt import translate_ppt


# =========================================================
# APPLICATION SETTINGS
# =========================================================

WINDOW_WIDTH = 760
WINDOW_HEIGHT = 500
WINDOW_RADIUS = 24
BORDER_SIZE = 2

TEAL = "#16CEC1"
LIGHT_TEAL = "#56E1D7"
DARK_TEAL = "#0D6862"
HOVER_TEAL = "#78E9E1"

BLACK = "#101212"
CARD_COLOR = "#172321"
CARD_BORDER = "#3BE8DC"

WHITE = "#FFFFFF"
SOFT_WHITE = "#E9FFFF"
GRAY = "#A9B8B6"
MUTED_GRAY = "#78918E"
YELLOW = "#FFD447"
ERROR_RED = "#FF7777"

TRANSPARENT_KEY = "#010203"

selected_file = ""

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


# =========================================================
# PATH HANDLING
# =========================================================

def resource_path(relative_path):
    """
    Returns the correct path when running normally
    or later as a packaged EXE.
    """

    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)


BACKGROUND_IMAGE = resource_path("colt_background.jpg")


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
# WINDOW HELPERS
# =========================================================

def apply_rounded_window():
    """
    Applies rounded corners to the entire application window
    on Windows.
    """

    if not sys.platform.startswith("win"):
        return

    try:
        import ctypes

        root.update_idletasks()

        hwnd = ctypes.windll.user32.GetParent(
            root.winfo_id()
        )

        region = ctypes.windll.gdi32.CreateRoundRectRgn(
            0,
            0,
            WINDOW_WIDTH + 1,
            WINDOW_HEIGHT + 1,
            WINDOW_RADIUS,
            WINDOW_RADIUS
        )

        ctypes.windll.user32.SetWindowRgn(
            hwnd,
            region,
            True
        )

    except Exception:
        pass


def start_move(event):
    root.drag_start_x = event.x_root
    root.drag_start_y = event.y_root


def move_window(event):
    difference_x = event.x_root - root.drag_start_x
    difference_y = event.y_root - root.drag_start_y

    new_x = root.winfo_x() + difference_x
    new_y = root.winfo_y() + difference_y

    root.geometry(f"+{new_x}+{new_y}")

    root.drag_start_x = event.x_root
    root.drag_start_y = event.y_root


def minimize_window():
    root.overrideredirect(False)
    root.iconify()


def restore_custom_window(event=None):
    if root.state() == "normal":
        root.overrideredirect(True)
        root.after(20, apply_rounded_window)


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
            (
                "Supported Documents",
                "*.docx *.xlsx *.pptx"
            ),
            (
                "Word Documents",
                "*.docx"
            ),
            (
                "Excel Workbooks",
                "*.xlsx"
            ),
            (
                "PowerPoint Presentations",
                "*.pptx"
            )
        ]
    )

    if not new_file:
        return

    selected_file = new_file
    filename = os.path.basename(selected_file)
    extension = os.path.splitext(filename)[1].upper()

    file_name_label.configure(
        text=filename,
        text_color=WHITE
    )

    selection_status_label.configure(
        text=f"{extension.replace('.', '')} FILE READY",
        text_color=LIGHT_TEAL
    )

    progress_bar.set(0)


# =========================================================
# TRANSLATION UI STATUS
# =========================================================

def show_translation_started():
    translate_button.configure(
        text="TRANSLATING...",
        state="disabled",
        fg_color="#547B78",
        hover_color="#547B78"
    )

    browse_button.configure(
        state="disabled"
    )

    language_dropdown.configure(
        state="disabled"
    )

    selection_status_label.configure(
        text="TRANSLATION IN PROGRESS",
        text_color=YELLOW
    )

    progress_bar.configure(
        mode="indeterminate"
    )

    progress_bar.start()


def restore_controls():
    progress_bar.stop()

    progress_bar.configure(
        mode="determinate"
    )

    translate_button.configure(
        text="START TRANSLATION",
        state="normal",
        fg_color=TEAL,
        hover_color=LIGHT_TEAL
    )

    browse_button.configure(
        state="normal"
    )

    language_dropdown.configure(
        state="normal"
    )


def show_translation_success(output_file):
    restore_controls()

    progress_bar.set(1)

    selection_status_label.configure(
        text="TRANSLATION COMPLETED",
        text_color=LIGHT_TEAL
    )

    messagebox.showinfo(
        "Translation Completed",
        "Your file has been translated successfully.\n\n"
        f"Output file:\n{os.path.basename(output_file)}"
    )


def show_translation_error(error_message):
    restore_controls()

    progress_bar.set(0)

    selection_status_label.configure(
        text="TRANSLATION FAILED",
        text_color=ERROR_RED
    )

    messagebox.showerror(
        "Translation Error",
        error_message
    )


# =========================================================
# TRANSLATION LOGIC
# =========================================================

def run_translation(
    input_file,
    output_file,
    source_language,
    extension
):
    try:
        if extension == ".docx":
            translate_docx(
                input_file,
                output_file,
                source_language
            )

        elif extension == ".xlsx":
            translate_excel(
                input_file,
                output_file,
                source_language
            )

        elif extension == ".pptx":
            translate_ppt(
                input_file,
                output_file,
                source_language
            )

        else:
            raise ValueError(
                "Only DOCX, XLSX, and PPTX files are supported."
            )

        root.after(
            0,
            lambda: show_translation_success(output_file)
        )

    except Exception as error:
        error_message = str(error)

        root.after(
            0,
            lambda: show_translation_error(error_message)
        )


def translate_file():
    if not selected_file:
        messagebox.showerror(
            "No File Selected",
            "Please select a DOCX, XLSX, or PPTX file."
        )
        return

    selected_language = language_dropdown.get()

    if selected_language not in LANGUAGES:
        messagebox.showerror(
            "Language Required",
            "Please select the source language."
        )
        return

    if not os.path.exists(selected_file):
        messagebox.showerror(
            "File Not Found",
            "The selected file no longer exists."
        )
        return

    source_language = LANGUAGES[selected_language]

    filename = os.path.basename(selected_file)
    file_base, extension = os.path.splitext(filename)
    extension = extension.lower()

    if extension not in [".docx", ".xlsx", ".pptx"]:
        messagebox.showerror(
            "Unsupported File",
            "Only DOCX, XLSX, and PPTX files are supported."
        )
        return

    # Project root ka output folder
    project_directory = os.path.dirname(
        os.path.dirname(os.path.abspath(__file__))
    )

    output_directory = os.path.join(
        project_directory,
        "output"
    )

    os.makedirs(
        output_directory,
        exist_ok=True
    )

    output_file = os.path.join(
        output_directory,
        f"{file_base}_English{extension}"
    )

    show_translation_started()

    translation_thread = threading.Thread(
        target=run_translation,
        args=(
            selected_file,
            output_file,
            source_language,
            extension
        ),
        daemon=True
    )

    translation_thread.start()


# =========================================================
# ROOT WINDOW
# =========================================================

root = ctk.CTk()

root.title("INFRA DOCUMENT TRANSLATOR")

root.geometry(
    f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
)

root.resizable(False, False)
root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
root.maxsize(WINDOW_WIDTH, WINDOW_HEIGHT)

root.configure(
    fg_color=TRANSPARENT_KEY
)

# Native title bar remove
root.overrideredirect(True)

# Screen centre
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

window_x = int(
    (screen_width - WINDOW_WIDTH) / 2
)

window_y = int(
    (screen_height - WINDOW_HEIGHT) / 2
)

root.geometry(
    f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
    f"+{window_x}+{window_y}"
)

root.bind(
    "<Map>",
    restore_custom_window
)


# =========================================================
# OUTER WINDOW BORDER
# =========================================================

outer_frame = ctk.CTkFrame(
    root,
    width=WINDOW_WIDTH,
    height=WINDOW_HEIGHT,
    corner_radius=WINDOW_RADIUS,
    fg_color=LIGHT_TEAL,
    border_width=0
)

outer_frame.place(
    x=0,
    y=0
)

outer_frame.pack_propagate(False)


# =========================================================
# INNER APPLICATION SURFACE
# =========================================================

app_surface = ctk.CTkFrame(
    outer_frame,
    width=WINDOW_WIDTH - (BORDER_SIZE * 2),
    height=WINDOW_HEIGHT - (BORDER_SIZE * 2),
    corner_radius=WINDOW_RADIUS - 2,
    fg_color=BLACK,
    border_width=0
)

app_surface.place(
    x=BORDER_SIZE,
    y=BORDER_SIZE
)

app_surface.pack_propagate(False)


# =========================================================
# BACKGROUND IMAGE
# =========================================================

if os.path.exists(BACKGROUND_IMAGE):
    try:
        background_pil = Image.open(
            BACKGROUND_IMAGE
        )

        background_pil = background_pil.resize(
            (
                WINDOW_WIDTH - (BORDER_SIZE * 2),
                WINDOW_HEIGHT - (BORDER_SIZE * 2)
            ),
            Image.Resampling.LANCZOS
        )

        background_image = ctk.CTkImage(
            light_image=background_pil,
            dark_image=background_pil,
            size=(
                WINDOW_WIDTH - (BORDER_SIZE * 2),
                WINDOW_HEIGHT - (BORDER_SIZE * 2)
            )
        )

        background_label = ctk.CTkLabel(
            app_surface,
            text="",
            image=background_image,
            corner_radius=WINDOW_RADIUS - 2
        )

        background_label.place(
            x=0,
            y=0
        )

    except Exception:
        pass


# =========================================================
# DARK CONTENT OVERLAY
# =========================================================

content_overlay = ctk.CTkFrame(
    app_surface,
    width=WINDOW_WIDTH - (BORDER_SIZE * 2),
    height=WINDOW_HEIGHT - (BORDER_SIZE * 2),
    corner_radius=WINDOW_RADIUS - 2,
    fg_color=("#101212", "#101212"),
    border_width=0
)

content_overlay.place(
    x=0,
    y=0
)

content_overlay.pack_propagate(False)


# =========================================================
# CUSTOM TITLE BAR
# =========================================================

top_bar = ctk.CTkFrame(
    content_overlay,
    width=WINDOW_WIDTH - (BORDER_SIZE * 2),
    height=44,
    corner_radius=0,
    fg_color=DARK_TEAL
)

top_bar.place(
    x=0,
    y=0
)

top_accent = ctk.CTkFrame(
    top_bar,
    width=10,
    height=44,
    corner_radius=0,
    fg_color=LIGHT_TEAL
)

top_accent.place(
    x=0,
    y=0
)

window_title = ctk.CTkLabel(
    top_bar,
    text="INFRA DOCUMENT TRANSLATOR",
    text_color=WHITE,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=13,
        weight="bold"
    )
)

window_title.place(
    x=27,
    rely=0.5,
    anchor="w"
)


# =========================================================
# TITLE BAR BUTTONS
# =========================================================

minimize_button = ctk.CTkButton(
    top_bar,
    text="—",
    width=44,
    height=40,
    corner_radius=8,
    fg_color="transparent",
    hover_color="#16877F",
    text_color=WHITE,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=18
    ),
    command=minimize_window
)

minimize_button.place(
    x=WINDOW_WIDTH - 102,
    y=2
)

close_button = ctk.CTkButton(
    top_bar,
    text="✕",
    width=44,
    height=40,
    corner_radius=8,
    fg_color="transparent",
    hover_color="#C42B1C",
    text_color=WHITE,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=16
    ),
    command=close_application
)

close_button.place(
    x=WINDOW_WIDTH - 56,
    y=2
)


# Drag window using title bar
for draggable_widget in [
    top_bar,
    top_accent,
    window_title
]:
    draggable_widget.bind(
        "<Button-1>",
        start_move
    )

    draggable_widget.bind(
        "<B1-Motion>",
        move_window
    )


# =========================================================
# HEADER SECTION
# =========================================================

heading_label = ctk.CTkLabel(
    content_overlay,
    text="INFRA DOCUMENT\nTRANSLATOR",
    justify="left",
    anchor="w",
    text_color=WHITE,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=27,
        weight="bold"
    )
)

heading_label.place(
    x=48,
    y=69
)

heading_accent = ctk.CTkFrame(
    content_overlay,
    width=30,
    height=6,
    corner_radius=3,
    fg_color=TEAL
)

heading_accent.place(
    x=48,
    y=158
)

subtitle_label = ctk.CTkLabel(
    content_overlay,
    text="Secure offline translation for DOCX, XLSX and PPTX",
    anchor="w",
    text_color=GRAY,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=11
    )
)

subtitle_label.place(
    x=88,
    y=150
)

colt_label = ctk.CTkLabel(
    content_overlay,
    text="colt",
    text_color=WHITE,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=39,
        weight="bold"
    )
)

colt_label.place(
    x=621,
    y=66
)


# =========================================================
# MAIN ROUNDED CARD
# =========================================================

main_card = ctk.CTkFrame(
    content_overlay,
    width=664,
    height=235,
    corner_radius=20,
    fg_color=CARD_COLOR,
    border_width=2,
    border_color=CARD_BORDER
)

main_card.place(
    x=46,
    y=203
)

main_card.pack_propagate(False)


# =========================================================
# LEFT CARD SECTION
# =========================================================

select_label = ctk.CTkLabel(
    main_card,
    text="SELECT DOCUMENT",
    text_color=LIGHT_TEAL,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=12,
        weight="bold"
    )
)

select_label.place(
    x=26,
    y=18
)

browse_button = ctk.CTkButton(
    main_card,
    text="SELECT FILE",
    command=browse_file,
    width=185,
    height=44,
    corner_radius=22,
    fg_color=LIGHT_TEAL,
    hover_color=WHITE,
    text_color=BLACK,
    border_width=0,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=13,
        weight="bold"
    )
)

browse_button.place(
    x=24,
    y=50
)

selection_status_label = ctk.CTkLabel(
    main_card,
    text="NO FILE SELECTED",
    text_color=GRAY,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=10,
        weight="bold"
    )
)

selection_status_label.place(
    x=27,
    y=108
)

file_name_label = ctk.CTkLabel(
    main_card,
    text="Choose a file to begin",
    width=270,
    height=35,
    anchor="nw",
    justify="left",
    wraplength=270,
    text_color=GRAY,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=11
    )
)

file_name_label.place(
    x=27,
    y=137
)


# =========================================================
# CARD DIVIDER
# =========================================================

divider = ctk.CTkFrame(
    main_card,
    width=1,
    height=175,
    corner_radius=0,
    fg_color="#385451"
)

divider.place(
    x=326,
    y=28
)


# =========================================================
# RIGHT CARD SECTION
# =========================================================

language_label = ctk.CTkLabel(
    main_card,
    text="SOURCE LANGUAGE",
    text_color=LIGHT_TEAL,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=12,
        weight="bold"
    )
)

language_label.place(
    x=356,
    y=18
)

language_dropdown = ctk.CTkOptionMenu(
    main_card,
    values=list(LANGUAGES.keys()),
    width=270,
    height=42,
    corner_radius=14,
    fg_color=SOFT_WHITE,
    button_color=TEAL,
    button_hover_color=LIGHT_TEAL,
    text_color=BLACK,
    dropdown_fg_color=SOFT_WHITE,
    dropdown_hover_color=LIGHT_TEAL,
    dropdown_text_color=BLACK,
    dynamic_resizing=False,
    anchor="w",
    font=ctk.CTkFont(
        family="Segoe UI",
        size=13,
        weight="bold"
    ),
    dropdown_font=ctk.CTkFont(
        family="Segoe UI",
        size=12
    )
)

language_dropdown.place(
    x=354,
    y=51
)

language_dropdown.set(
    "Select Language"
)


# =========================================================
# TRANSLATE BUTTON
# =========================================================

translate_button = ctk.CTkButton(
    main_card,
    text="START TRANSLATION",
    command=translate_file,
    width=270,
    height=48,
    corner_radius=24,
    fg_color=TEAL,
    hover_color=LIGHT_TEAL,
    text_color=BLACK,
    border_width=0,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=13,
        weight="bold"
    )
)

translate_button.place(
    x=354,
    y=112
)


# =========================================================
# PROGRESS BAR
# =========================================================

progress_bar = ctk.CTkProgressBar(
    main_card,
    width=270,
    height=6,
    corner_radius=3,
    progress_color=LIGHT_TEAL,
    fg_color="#294340",
    mode="determinate"
)

progress_bar.place(
    x=354,
    y=172
)

progress_bar.set(0)


# =========================================================
# SUPPORTED FORMATS
# =========================================================

format_label = ctk.CTkLabel(
    main_card,
    text="SUPPORTED  •  DOCX  •  XLSX  •  PPTX",
    text_color=MUTED_GRAY,
    font=ctk.CTkFont(
        family="Segoe UI",
        size=9,
        weight="bold"
    )
)

format_label.place(
    x=354,
    y=188
)


# =========================================================
# START APPLICATION
# =========================================================

root.after(
    50,
    apply_rounded_window
)

root.mainloop()