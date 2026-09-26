import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import re
import sys
from converter import convert_file

def resource_path(filename):
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, filename)


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
# Filename validation
# -----------------------------------

def has_valid_filename(filename):
    """
    Check whether the filename follows
    the established naming format:

    serial_YYYYMMDD_number.txt
    """

    return re.match(
        r"^serial_\d{8}_\d+\.txt$",
        os.path.basename(filename)
    ) is not None


# -----------------------------------
# Detect date from filename
# -----------------------------------

def get_filename_date(filename):
    """
    Try to extract a date from a filename.

    Returns:
        MM/DD/YYYY
        or None
    """

    match = re.search(
        r"serial_(\d{8})",
        os.path.basename(filename)
    )

    if not match:
        return None

    date_text = match.group(1)

    try:

        from datetime import datetime

        return datetime.strptime(
            date_text,
            "%Y%m%d"
        ).strftime("%m/%d/%Y")

    except ValueError:

        return None

# -----------------------------------
# Ask user for date
# -----------------------------------

def ask_for_date(filename):
    """
    Ask the user to select a date for a
    filename that does not follow the
    established naming format.

    Returns:
        MM/DD/YYYY
        or None if cancelled.
    """

    from datetime import datetime
    import calendar

    dialog = tk.Toplevel(window)

    dialog.title("Select File Date")
    dialog.geometry("420x260")
    dialog.resizable(False, False)

    dialog.configure(
        bg=BG_COLOR
    )

    dialog.transient(window)
    dialog.grab_set()

    result = {
        "date": None
    }

    # -----------------------------------
    # Title
    # -----------------------------------

    title = tk.Label(
        dialog,
        text="Filename Date Required",
        font=(
            "Arial",
            15,
            "bold"
        ),
        bg=PURPLE,
        fg=WHITE
    )

    title.pack(
        fill="x",
        pady=(0, 15),
        ipady=8
    )

    # -----------------------------------
    # Filename
    # -----------------------------------

    filename_label = tk.Label(
        dialog,
        text=(
            "The filename does not follow the\n"
            "expected format:\n\n"
            "serial_YYYYMMDD_number.txt\n\n"
            f"File: {filename}"
        ),
        font=(
            "Arial",
            10
        ),
        bg=BG_COLOR,
        fg=TEXT_COLOR,
        justify="center"
    )

    filename_label.pack(
        pady=(0, 12)
    )

    # -----------------------------------
    # Date frame
    # -----------------------------------

    date_frame = tk.Frame(
        dialog,
        bg=BG_COLOR
    )

    date_frame.pack()

    # -----------------------------------
    # Current date
    # -----------------------------------

    now = datetime.now()

    years = list(
        range(
            now.year - 5,
            now.year + 6
        )
    )

    months = [
        f"{month:02d}"
        for month in range(1, 13)
    ]

    days = [
        f"{day:02d}"
        for day in range(1, 32)
    ]

    # -----------------------------------
    # Month
    # -----------------------------------

    month_var = tk.StringVar(
        value=f"{now.month:02d}"
    )

    month_box = ttk.Combobox(
        date_frame,
        textvariable=month_var,
        values=months,
        width=5,
        state="readonly"
    )

    month_box.pack(
        side="left",
        padx=3
    )

    # -----------------------------------
    # Day
    # -----------------------------------

    day_var = tk.StringVar(
        value=f"{now.day:02d}"
    )

    day_box = ttk.Combobox(
        date_frame,
        textvariable=day_var,
        values=days,
        width=5,
        state="readonly"
    )

    day_box.pack(
        side="left",
        padx=3
    )

    # -----------------------------------
    # Year
    # -----------------------------------

    year_var = tk.StringVar(
        value=str(now.year)
    )

    year_box = ttk.Combobox(
        date_frame,
        textvariable=year_var,
        values=[str(year) for year in years],
        width=7,
        state="readonly"
    )

    year_box.pack(
        side="left",
        padx=3
    )

    # -----------------------------------
    # Buttons
    # -----------------------------------

    button_frame = tk.Frame(
        dialog,
        bg=BG_COLOR
    )

    button_frame.pack(
        pady=20
    )

    def use_date():

        try:

            selected = datetime.strptime(
                f"{year_var.get()}-{month_var.get()}-{day_var.get()}",
                "%Y-%m-%d"
            )

        except ValueError:

            messagebox.showerror(
                "Invalid Date",
                "Please select a valid date.",
                parent=dialog
            )

            return

        result["date"] = selected.strftime(
            "%m/%d/%Y"
        )

        dialog.destroy()

    def cancel():

        dialog.destroy()

    cancel_button = tk.Button(
        button_frame,
        text="Cancel",
        command=cancel,
        width=12,
        bg=BLUE,
        fg=WHITE,
        activebackground=BLUE_DARK,
        activeforeground=WHITE,
        relief="flat",
        cursor="hand2"
    )

    cancel_button.pack(
        side="left",
        padx=5
    )

    use_button = tk.Button(
        button_frame,
        text="Use Date",
        command=use_date,
        width=12,
        bg=PURPLE,
        fg=WHITE,
        activebackground=PURPLE_DARK,
        activeforeground=WHITE,
        relief="flat",
        cursor="hand2"
    )

    use_button.pack(
        side="left",
        padx=5
    )

    window.wait_window(dialog)

    return result["date"]


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
            # Determine file date
            # -----------------------------------

            selected_date = None

            if not has_valid_filename(
                original_filename
            ):

                detected_date = get_filename_date(
                    original_filename
                )

                if detected_date:

                    use_detected = messagebox.askyesno(
                        "Filename Format Issue",
                        (
                            f"The filename does not follow the "
                            f"expected format.\n\n"
                            f"File:\n{original_filename}\n\n"
                            f"We found this date in the filename:\n"
                            f"{detected_date}\n\n"
                            f"Use this date for the conversion?"
                        )
                    )

                    if use_detected:

                        selected_date = detected_date

                    else:

                        selected_date = ask_for_date(
                            original_filename
                        )

                else:

                    selected_date = ask_for_date(
                        original_filename
                    )


                # -----------------------------------
                # User cancelled date selection
                # -----------------------------------

                if selected_date is None:

                    raise ValueError(
                        "Date selection was cancelled."
                    )


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
                output_file,
                selected_date
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
#branding
icon = tk.PhotoImage(file=resource_path("logo.png"))

#Client
#icon = tk.PhotoImage(file=resource_path("logo.png"))

window.iconphoto(True, icon)


window.title(
    "HPS Converter"
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
    text="HPS CONVERTER",
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

#Branding
logo_image = tk.PhotoImage(file=resource_path("logo.png"))
logo_image = logo_image.subsample(2, 2)
# Client Logo
#logo_image = tk.PhotoImage(file=resource_path("logo.png"))


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

powered_by_label = tk.Label(
    logo_box,
    text="Powered by Wilhelm-Clark",
    bg="#E8E8E8",
    font=("Arial", 8)
)

powered_by_label.pack(pady=(3, 0))


# ===================================
# START GUI
# ===================================

window.mainloop()

