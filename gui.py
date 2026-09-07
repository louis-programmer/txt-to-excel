import tkinter as tk
from tkinter import filedialog, messagebox
import os
import re

from converter import convert_file


# -----------------------------------
# GUI Colors
# -----------------------------------

BG_COLOR = "#F5F5F5"

PURPLE = "#7030A0"
PURPLE_DARK = "#5B2485"

YELLOW = "#FFF2CC"
YELLOW_DARK = "#FFD966"

BLUE = "#4472C4"
BLUE_DARK = "#2F5597"

WHITE = "#FFFFFF"
TEXT_COLOR = "#222222"


# -----------------------------------
# Filename safety
# -----------------------------------

def make_safe_filename(filename):
    """
    Make sure the filename is safe for Windows/Linux.
    """

    name, _ = os.path.splitext(filename)

    name = re.sub(
        r'[<>:"/\\|?*]',
        "_",
        name
    )

    name = name.rstrip(" .")

    if not name:
        name = "converted_file"

    return name + ".xlsx"


# -----------------------------------
# Prevent overwriting existing files
# -----------------------------------

def get_unique_output_path(output_folder, filename):
    """
    Return a filename that does not overwrite
    an existing file.
    """

    base_name, extension = os.path.splitext(filename)

    output_path = os.path.join(
        output_folder,
        filename
    )

    counter = 1

    while os.path.exists(output_path):

        output_path = os.path.join(
            output_folder,
            f"{base_name}_{counter}{extension}"
        )

        counter += 1

    return output_path


# -----------------------------------
# Browse TXT files
# -----------------------------------

def browse_files():

    files = filedialog.askopenfilenames(
        title="Select TXT Files",
        filetypes=[
            ("TXT files", "*.txt"),
            ("Text files", "*.log"),
            ("All files", "*.*")
        ]
    )

    if not files:
        return

    selected_files.clear()

    for file in files:
        selected_files.append(file)

    update_file_list()

    status_label.config(
        text=f"{len(selected_files)} file(s) selected."
    )


# -----------------------------------
# Clear selected files
# -----------------------------------

def clear_files():

    selected_files.clear()

    update_file_list()

    status_label.config(
        text="Ready"
    )


# -----------------------------------
# Update file list
# -----------------------------------

def update_file_list():

    file_listbox.delete(
        0,
        tk.END
    )

    for file in selected_files:

        file_listbox.insert(
            tk.END,
            os.path.basename(file)
        )


# -----------------------------------
# Browse output folder
# -----------------------------------

def browse_output_folder():

    folder = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if not folder:
        return

    output_folder.set(folder)

    status_label.config(
        text="Output folder selected."
    )


# -----------------------------------
# Convert files
# -----------------------------------

def convert():

    # -----------------------------------
    # Check files
    # -----------------------------------

    if not selected_files:

        messagebox.showwarning(
            "No files selected",
            "Please select at least one TXT file."
        )

        return


    # -----------------------------------
    # Check output folder
    # -----------------------------------

    folder = output_folder.get()

    if not folder:

        messagebox.showwarning(
            "No output folder",
            "Please select an output folder first."
        )

        return


    if not os.path.isdir(folder):

        messagebox.showerror(
            "Invalid folder",
            "The selected output folder does not exist."
        )

        return


    # -----------------------------------
    # Conversion counters
    # -----------------------------------

    successful = 0
    failed = 0

    failed_files = []


    # -----------------------------------
    # Convert each TXT file
    # -----------------------------------

    for input_file in selected_files:

        original_filename = os.path.basename(
            input_file
        )

        try:

            # -----------------------------------
            # Create Excel filename
            # -----------------------------------

            excel_filename = make_safe_filename(
                original_filename
            )


            # -----------------------------------
            # Prevent overwrite
            # -----------------------------------

            output_file = get_unique_output_path(
                folder,
                excel_filename
            )


            # -----------------------------------
            # Convert
            # -----------------------------------

            convert_file(
                input_file,
                output_file
            )

            successful += 1


        except Exception as error:

            failed += 1

            failed_files.append(
                f"{original_filename}: {error}"
            )


    # -----------------------------------
    # Update status
    # -----------------------------------

    if failed == 0:

        status_label.config(
            text=(
                f"Conversion completed. "
                f"{successful} file(s) converted."
            )
        )

    else:

        status_label.config(
            text=(
                f"Completed with errors. "
                f"{successful} succeeded, "
                f"{failed} failed."
            )
        )


    # -----------------------------------
    # Result message
    # -----------------------------------

    message = (
        "Conversion completed.\n\n"
        f"Files selected: {len(selected_files)}\n"
        f"Successful: {successful}\n"
        f"Failed: {failed}\n\n"
        f"Output folder:\n{folder}"
    )


    if failed_files:

        message += (
            "\n\nFailed files:\n\n"
            + "\n".join(failed_files)
        )


    # -----------------------------------
    # Show result
    # -----------------------------------

    if failed == 0:

        messagebox.showinfo(
            "Conversion Complete",
            message
        )

    else:

        messagebox.showwarning(
            "Conversion Completed With Errors",
            message
        )


# ===================================
# GUI
# ===================================

window = tk.Tk()

icon = tk.PhotoImage(file="logo.png")
window.iconphoto(True, icon)


window.title(
    "TXT to Excel"
)


window.geometry(
    "700x550"
)

window.resizable(
    False,
    False
)

window.configure(
    bg=BG_COLOR
)


# -----------------------------------
# Selected files
# -----------------------------------

selected_files = []


# -----------------------------------
# Variables
# -----------------------------------

output_folder = tk.StringVar()


# ===================================
# HEADER
# ===================================

header_frame = tk.Frame(
    window,
    bg=PURPLE,
    height=65
)

header_frame.pack(
    fill="x",
    pady=0
)
header_frame.pack_propagate(False)


# -----------------------------------
# Header title
# -----------------------------------

title_label = tk.Label(
    header_frame,
    text="TXT → EXCEL",
    font=(
        "Arial",
        20,
        "bold"
    ),
    bg=PURPLE,
    fg=WHITE
)

title_label.pack(
    side="left",
    padx=25
)


# -----------------------------------
# Yellow underline
# -----------------------------------

header_underline = tk.Frame(
    window,
    bg=YELLOW_DARK,
    height=5
)

header_underline.pack(
    fill="x"
)




# ===================================
# FILES LABEL
# ===================================

files_label = tk.Label(
    window,
    text="TXT Files:",
    font=(
        "Arial",
        11,
        "bold"
    ),
    bg=BG_COLOR,
    fg=TEXT_COLOR
)

files_label.pack(
    anchor="w",
    padx=35,
    pady=(20, 5)
)


# ===================================
# FILE LIST
# ===================================

file_list_frame = tk.Frame(
    window,
    bg=BG_COLOR
)

file_list_frame.pack(
    padx=35,
    fill="x"
)


file_listbox = tk.Listbox(
    file_list_frame,
    height=8,
    width=75,
    font=(
        "Arial",
        10
    ),
    bg=YELLOW,
    fg=TEXT_COLOR,
    selectbackground=PURPLE,
    selectforeground=WHITE,
    relief="solid",
    borderwidth=1
)

file_listbox.pack(
    side="left",
    fill="both",
    expand=True
)


# -----------------------------------
# Scrollbar
# -----------------------------------

scrollbar = tk.Scrollbar(
    file_list_frame,
    orient="vertical",
    command=file_listbox.yview
)

scrollbar.pack(
    side="right",
    fill="y"
)

file_listbox.config(
    yscrollcommand=scrollbar.set
)


# ===================================
# FILE BUTTONS
# ===================================

file_button_frame = tk.Frame(
    window,
    bg=BG_COLOR
)

file_button_frame.pack(
    pady=10
)


# -----------------------------------
# Add TXT Files
# -----------------------------------

add_button = tk.Button(
    file_button_frame,
    text="Add TXT Files",
    command=browse_files,
    width=18,
    bg=PURPLE,
    fg=WHITE,
    activebackground=PURPLE_DARK,
    activeforeground=WHITE,
    relief="flat",
    cursor="hand2"
)

add_button.pack(
    side="left",
    padx=5
)


# -----------------------------------
# Clear
# -----------------------------------

clear_button = tk.Button(
    file_button_frame,
    text="Clear",
    command=clear_files,
    width=12,
    bg=BLUE,
    fg=WHITE,
    activebackground=BLUE_DARK,
    activeforeground=WHITE,
    relief="flat",
    cursor="hand2"
)

clear_button.pack(
    side="left",
    padx=5
)


# ===================================
# OUTPUT FOLDER
# ===================================

output_label = tk.Label(
    window,
    text="Output Folder:",
    font=(
        "Arial",
        11,
        "bold"
    ),
    bg=BG_COLOR,
    fg=TEXT_COLOR
)

output_label.pack(
    anchor="w",
    padx=35,
    pady=(15, 5)
)


# -----------------------------------
# Output frame
# -----------------------------------

output_frame = tk.Frame(
    window,
    bg=BG_COLOR
)

output_frame.pack(
    padx=35,
    fill="x"
)


# -----------------------------------
# Output entry
# -----------------------------------

output_entry = tk.Entry(
    output_frame,
    textvariable=output_folder,
    width=60,
    bg=WHITE,
    fg=TEXT_COLOR,
    relief="solid"
)

output_entry.pack(
    side="left",
    padx=(0, 5)
)


# -----------------------------------
# Browse button
# -----------------------------------

output_button = tk.Button(
    output_frame,
    text="Browse",
    command=browse_output_folder,
    width=10,
    bg=PURPLE,
    fg=WHITE,
    activebackground=PURPLE_DARK,
    activeforeground=WHITE,
    relief="flat",
    cursor="hand2"
)

output_button.pack(
    side="left"
)


# ===================================
# CONVERT BUTTON
# ===================================

convert_button = tk.Button(
    window,
    text="Convert to Excel",
    command=convert,
    width=25,
    height=2,
    font=(
        "Arial",
        11,
        "bold"
    ),
    bg=YELLOW_DARK,
    fg=TEXT_COLOR,
    activebackground=YELLOW,
    activeforeground=TEXT_COLOR,
    relief="flat",
    cursor="hand2"
)

convert_button.pack(
    pady=25
)


# ===================================
# STATUS
# ===================================

status_label = tk.Label(
    window,
    text="Ready",
    font=(
        "Arial",
        10
    ),
    bg=BG_COLOR,
    fg=TEXT_COLOR
)

status_label.pack()

# ===================================
# BOTTOM-RIGHT LOGO
# ===================================

logo_image = tk.PhotoImage(
    file="logo.png"
)


# -----------------------------------
# Logo safe box
# -----------------------------------

logo_box = tk.Frame(
    window,
    bg="#E8E8E8",
    padx=8,
    pady=8,
    relief="flat"
)

logo_box.place(
    relx=1.0,
    rely=1.0,
    anchor="se",
    x=-15,
    y=-10
)


# -----------------------------------
# Logo
# -----------------------------------

logo_label = tk.Label(
    logo_box,
    image=logo_image,
    bg="#E8E8E8",
    borderwidth=0
)

logo_label.pack()
# ===================================
# START GUI
# ===================================

window.mainloop()

