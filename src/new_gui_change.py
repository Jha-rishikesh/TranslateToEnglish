import os
import sys
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
WINDOW_RADIUS = 22
BORDER_SIZE = 3

TEAL = "#16CEC1"
LIGHT_TEAL = "#56E1D7"
DARK_TEAL = "#0D6862"
HOVER_TEAL = LIGHT_TEAL
BLACK = "#101212"
CARD_COLOR = "#172321"
CARD_BORDER = "#3BE8DC"
WHITE = "#FFFFFF"
SOFT_WHITE = "#E9FFFF"
GRAY = "#A9B8B6"
YELLOW = "#FFD447"
ERROR_RED = "#FF7777"
TRANSPARENT_KEY = "#010203"

BACKGROUND_IMAGE = "colt_background.jpg"

selected_file = ""
translate_button = None
background_photo = None

LANGUAGES = {
    "French": "fr",
    "German": "de",
    "Spanish": "es",
    "Portuguese": "pt",
    "Japanese": "ja",
    # "Chinese": "zh",
}


# =========================================================
# WINDOW HELPERS
# =========================================================
def apply_rounded_window():
    """Apply rounded corners on Windows while keeping the custom title bar."""
    if not sys.platform.startswith("win"):
        return

    try:
        import ctypes

        root.update_idletasks()
        hwnd = ctypes.windll.user32.GetParent(root.winfo_id())
        region = ctypes.windll.gdi32.CreateRoundRectRgn(
            0,
            0,
            WINDOW_WIDTH + 1,
            WINDOW_HEIGHT + 1,
            WINDOW_RADIUS,
            WINDOW_RADIUS,
        )
        ctypes.windll.user32.SetWindowRgn(hwnd, region, True)
    except Exception:
        # GUI will still work if Windows rejects the rounded-region call.
        pass


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
    if root.state() == "normal":
        root.overrideredirect(True)
        root.after(10, apply_rounded_window)


def close_application():
    root.destroy()


# =========================================================
# DRAWING HELPERS
# =========================================================
def rounded_rectangle(canvas, x1, y1, x2, y2, radius, **kwargs):
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
        x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, splinesteps=36, **kwargs)


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
        font=("Segoe UI", 10, "bold"),
    ):
        parent_background = parent.cget("bg")
        super().__init__(
            parent,
            width=width,
            height=height,
            bg=parent_background,
            highlightthickness=0,
            bd=0,
            cursor="hand2",
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

    def draw_button(self, color):
        self.delete("all")
        rounded_rectangle(
            self,
            2,
            2,
            self.button_width - 2,
            self.button_height - 2,
            self.radius,
            fill=color,
            outline=color,
        )
        self.create_text(
            self.button_width / 2,
            self.button_height / 2,
            text=self.button_text,
            fill=self.text_color,
            font=self.button_font,
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
            self.draw_button("#78928F")


# =========================================================
# FILE BROWSER AND TRANSLATION
# =========================================================
def browse_file():
    global selected_file

    new_file = filedialog.askopenfilename(
        title="Select a document",
        filetypes=[
            ("Supported Documents", "*.docx *.xlsx *.pptx"),
            ("Word Documents", "*.docx"),
            ("Excel Workbooks", "*.xlsx"),
            ("PowerPoint Presentations", "*.pptx"),
        ],
    )

    if new_file:
        selected_file = new_file
        filename = os.path.basename(selected_file)
        file_name_label.config(text=filename, fg=WHITE)
        selection_status_label.config(text="FILE READY", fg=LIGHT_TEAL)


def translate_file():
    global selected_file

    if not selected_file:
        messagebox.showerror(
            "No File Selected",
            "Please select a DOCX, XLSX, or PPTX file.",
        )
        return

    selected_language = language_var.get()
    if selected_language not in LANGUAGES:
        messagebox.showerror(
            "Language Required",
            "Please select the source language.",
        )
        return

    source_language = LANGUAGES[selected_language]
    filename = os.path.basename(selected_file)
    file_base, extension = os.path.splitext(filename)

    output_directory = os.path.join(os.path.dirname(selected_file), "output")
    os.makedirs(output_directory, exist_ok=True)
    output_file = os.path.join(
        output_directory,
        f"{file_base}_English{extension}",
    )

    translate_button.set_text("TRANSLATING...")
    translate_button.set_enabled(False)
    selection_status_label.config(text="TRANSLATION IN PROGRESS", fg=YELLOW)
    root.update_idletasks()

    try:
        if extension.lower() == ".docx":
            translate_docx(selected_file, output_file, source_language)
        elif extension.lower() == ".xlsx":
            translate_excel(selected_file, output_file, source_language)
        elif extension.lower() == ".pptx":
            translate_ppt(selected_file, output_file, source_language)
        else:
            messagebox.showerror(
                "Unsupported File",
                "Only DOCX, XLSX, and PPTX files are supported.",
            )
            return

        selection_status_label.config(text="TRANSLATION COMPLETED", fg=LIGHT_TEAL)
        messagebox.showinfo(
            "Translation Completed",
            "Your file has been translated successfully.\n\n"
            f"Output file:\n{os.path.basename(output_file)}",
        )
    except Exception as error:
        selection_status_label.config(text="TRANSLATION FAILED", fg=ERROR_RED)
        messagebox.showerror("Translation Error", str(error))
    finally:
        translate_button.set_text("START TRANSLATION")
        translate_button.set_enabled(True)


# =========================================================
# ROOT WINDOW
# =========================================================
root = tk.Tk()
root.title("INFRA DOCUMENT TRANSLATOR")
root.configure(bg=TRANSPARENT_KEY)
root.overrideredirect(True)
root.resizable(False, False)
root.minsize(WINDOW_WIDTH, WINDOW_HEIGHT)
root.maxsize(WINDOW_WIDTH, WINDOW_HEIGHT)

screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_x = int((screen_width - WINDOW_WIDTH) / 2)
window_y = int((screen_height - WINDOW_HEIGHT) / 2)
root.geometry(f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{window_x}+{window_y}")
root.bind("<Map>", restore_custom_title_bar)

# =========================================================
# OUTER TEAL BORDER AND INNER SURFACE
# =========================================================
outer_shell = tk.Canvas(
    root,
    width=WINDOW_WIDTH,
    height=WINDOW_HEIGHT,
    bg=TRANSPARENT_KEY,
    highlightthickness=0,
    bd=0,
)
outer_shell.place(x=0, y=0)

rounded_rectangle(
    outer_shell,
    0,
    0,
    WINDOW_WIDTH - 1,
    WINDOW_HEIGHT - 1,
    WINDOW_RADIUS,
    fill=LIGHT_TEAL,
    outline=LIGHT_TEAL,
)
rounded_rectangle(
    outer_shell,
    BORDER_SIZE,
    BORDER_SIZE,
    WINDOW_WIDTH - BORDER_SIZE - 1,
    WINDOW_HEIGHT - BORDER_SIZE - 1,
    WINDOW_RADIUS - 3,
    fill=BLACK,
    outline=BLACK,
)

app_surface = tk.Frame(root, bg=BLACK)
app_surface.place(
    x=BORDER_SIZE,
    y=BORDER_SIZE,
    width=WINDOW_WIDTH - (BORDER_SIZE * 2),
    height=WINDOW_HEIGHT - (BORDER_SIZE * 2),
)

# =========================================================
# OPTIONAL BACKGROUND IMAGE
# =========================================================
try:
    background_image = Image.open(BACKGROUND_IMAGE)
    background_image = background_image.resize(
        (WINDOW_WIDTH - 6, WINDOW_HEIGHT - 6),
        Image.Resampling.LANCZOS,
    )
    background_image = ImageEnhance.Brightness(background_image).enhance(0.38)
    background_photo = ImageTk.PhotoImage(background_image)
    background_label = tk.Label(app_surface, image=background_photo, bd=0)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)
except Exception:
    pass

# =========================================================
# CUSTOM TOP BAR
# =========================================================
top_bar = tk.Frame(app_surface, bg=DARK_TEAL)
top_bar.place(x=0, y=0, width=WINDOW_WIDTH - 6, height=44)

top_accent = tk.Frame(top_bar, bg=LIGHT_TEAL, width=10)
top_accent.pack(side="left", fill="y")

window_title = tk.Label(
    top_bar,
    text="INFRA DOCUMENT TRANSLATOR",
    bg=DARK_TEAL,
    fg=WHITE,
    font=("Segoe UI", 10, "bold"),
)
window_title.pack(side="left", padx=14)

close_button = tk.Label(
    top_bar,
    text="✕",
    bg=DARK_TEAL,
    fg=WHITE,
    font=("Segoe UI", 13),
    cursor="hand2",
    width=4,
)
close_button.pack(side="right", fill="y")

minimize_button = tk.Label(
    top_bar,
    text="—",
    bg=DARK_TEAL,
    fg=WHITE,
    font=("Segoe UI", 13),
    cursor="hand2",
    width=4,
)
minimize_button.pack(side="right", fill="y")

close_button.bind("<Button-1>", lambda event: close_application())
minimize_button.bind("<Button-1>", lambda event: minimize_window())
close_button.bind("<Enter>", lambda event: close_button.config(bg="#C42B1C"))
close_button.bind("<Leave>", lambda event: close_button.config(bg=DARK_TEAL))
minimize_button.bind("<Enter>", lambda event: minimize_button.config(bg="#16877F"))
minimize_button.bind("<Leave>", lambda event: minimize_button.config(bg=DARK_TEAL))

for draggable_widget in [top_bar, window_title, top_accent]:
    draggable_widget.bind("<Button-1>", start_move)
    draggable_widget.bind("<B1-Motion>", move_window)

# =========================================================
# HEADER
# =========================================================
heading_label = tk.Label(
    app_surface,
    text="INFRA DOCUMENT\nTRANSLATOR",
    justify="left",
    anchor="w",
    bg=BLACK,
    fg=WHITE,
    font=("Segoe UI", 24, "bold"),
)
heading_label.place(x=46, y=66)

# Replaced the long underline with a compact accent pill.
accent_canvas = tk.Canvas(
    app_surface,
    width=30,
    height=8,
    bg=BLACK,
    highlightthickness=0,
    bd=0,
)
accent_canvas.place(x=47, y=157)
rounded_rectangle(
    accent_canvas,
    0,
    1,
    28,
    7,
    4,
    fill=TEAL,
    outline=TEAL,
)

subtitle_label = tk.Label(
    app_surface,
    text="Secure offline translation for DOCX, XLSX and PPTX",
    justify="left",
    anchor="w",
    bg=BLACK,
    fg=GRAY,
    font=("Segoe UI", 9),
)
subtitle_label.place(x=84, y=151)

colt_label = tk.Label(
    app_surface,
    text="colt",
    bg=BLACK,
    fg=WHITE,
    font=("Segoe UI", 30, "bold"),
)
colt_label.place(x=622, y=72)

# =========================================================
# ROUNDED MAIN CARD
# =========================================================
card_canvas = tk.Canvas(
    app_surface,
    width=664,
    height=235,
    bg=BLACK,
    highlightthickness=0,
    bd=0,
)
card_canvas.place(x=45, y=202)

rounded_rectangle(
    card_canvas,
    3,
    5,
    661,
    233,
    18,
    fill="#082E2C",
    outline="#082E2C",
)
rounded_rectangle(
    card_canvas,
    1,
    1,
    658,
    228,
    18,
    fill=CARD_COLOR,
    outline=CARD_BORDER,
    width=2,
)

select_label = tk.Label(
    card_canvas,
    text="SELECT DOCUMENT",
    bg=CARD_COLOR,
    fg=LIGHT_TEAL,
    font=("Segoe UI", 9, "bold"),
)
select_label.place(x=25, y=20)

browse_button = RoundedButton(
    card_canvas,
    text="SELECT FILE",
    command=browse_file,
    width=185,
    height=44,
    radius=22,
    bg_color=TEAL,
    hover_color=LIGHT_TEAL,
    text_color=BLACK,
)
browse_button.place(x=23, y=49)

selection_status_label = tk.Label(
    card_canvas,
    text="NO FILE SELECTED",
    bg=CARD_COLOR,
    fg=GRAY,
    font=("Segoe UI", 8, "bold"),
)
selection_status_label.place(x=25, y=106)

file_name_label = tk.Label(
    card_canvas,
    text="Choose a file to begin",
    bg=CARD_COLOR,
    fg=GRAY,
    anchor="w",
    justify="left",
    font=("Segoe UI", 9),
    wraplength=270,
)
file_name_label.place(x=25, y=130, width=275)

# Soft divider rather than a bright hard line.
card_canvas.create_line(327, 28, 327, 202, fill="#385451", width=1)

language_label = tk.Label(
    card_canvas,
    text="SOURCE LANGUAGE",
    bg=CARD_COLOR,
    fg=LIGHT_TEAL,
    font=("Segoe UI", 9, "bold"),
)
language_label.place(x=356, y=20)

language_var = tk.StringVar(value="Select Language")

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
    font=("Segoe UI", 10, "bold"),
)
style.map(
    "Language.TCombobox",
    fieldbackground=[("readonly", SOFT_WHITE)],
    foreground=[("readonly", BLACK)],
    background=[("readonly", TEAL)],
    bordercolor=[("focus", LIGHT_TEAL), ("readonly", TEAL)],
)

language_dropdown = ttk.Combobox(
    card_canvas,
    textvariable=language_var,
    values=list(LANGUAGES.keys()),
    state="readonly",
    width=23,
    style="Language.TCombobox",
    font=("Segoe UI", 10),
)
language_dropdown.place(x=354, y=49, width=270, height=42)

root.option_add("*TCombobox*Listbox.background", SOFT_WHITE)
root.option_add("*TCombobox*Listbox.foreground", BLACK)
root.option_add("*TCombobox*Listbox.selectBackground", LIGHT_TEAL)
root.option_add("*TCombobox*Listbox.selectForeground", BLACK)

translate_button = RoundedButton(
    card_canvas,
    text="START TRANSLATION",
    command=translate_file,
    width=210,
    height=48,
    radius=22,
    bg_color=TEAL,
    hover_color=LIGHT_TEAL,
    text_color=BLACK,
)
translate_button.place(x=354, y=114)

format_label = tk.Label(
    card_canvas,
    text="SUPPORTED  •  DOCX  •  XLSX  •  PPTX",
    bg=CARD_COLOR,
    fg="#75908D",
    font=("Segoe UI", 7, "bold"),
)
format_label.place(x=354, y=180)

root.after(30, apply_rounded_window)
root.mainloop()
