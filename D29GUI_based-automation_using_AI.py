import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os


# ---------------------------------------------------------
# GLOBAL VARIABLES
# ---------------------------------------------------------

input_file = ""


# ---------------------------------------------------------
# BROWSE INPUT FILE
# ---------------------------------------------------------

def browse_input_file():
    global input_file

    file_path = filedialog.askopenfilename(
        title="Select Raw Excel File",
        filetypes=[
            ("Excel Files", "*.xlsx *.xls"),
            ("Excel Workbook", "*.xlsx"),
            ("Excel 97-2003", "*.xls")
        ]
    )

    if file_path:
        input_file = file_path
        input_entry.delete(0, tk.END)
        input_entry.insert(0, file_path)


# ---------------------------------------------------------
# BROWSE OUTPUT FOLDER
# ---------------------------------------------------------

def browse_output_folder():
    folder_path = filedialog.askdirectory(
        title="Select Output Folder"
    )

    if folder_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, folder_path)


# ---------------------------------------------------------
# RESET PATHS
# ---------------------------------------------------------

def reset_paths():
    global input_file

    input_file = ""

    input_entry.delete(0, tk.END)
    output_entry.delete(0, tk.END)

    status_label.config(
        text="Input and output paths cleared."
    )


# ---------------------------------------------------------
# RESET CHECKBOXES
# ---------------------------------------------------------

def reset_checkboxes():
    duplicate_var.set(False)
    blank_var.set(False)
    spaces_var.set(False)
    title_var.set(False)

    status_label.config(
        text="Cleaning options reset."
    )


# ---------------------------------------------------------
# CLEAN TEXT VALUES
# ---------------------------------------------------------

def clean_text_columns(df):

    # Find text/object columns
    text_columns = df.select_dtypes(
        include=["object", "string"]
    ).columns

    # Remove leading and trailing spaces
    if spaces_var.get():

        for column in text_columns:

            df[column] = df[column].apply(
                lambda x: x.strip()
                if isinstance(x, str)
                else x
            )

    # Convert text to Title Case
    if title_var.get():

        for column in text_columns:

            df[column] = df[column].apply(
                lambda x: x.title()
                if isinstance(x, str)
                else x
            )

    return df


# ---------------------------------------------------------
# REMOVE BLANK ROWS
# ---------------------------------------------------------

def remove_blank_rows(df):

    # Consider cells containing only spaces as blank
    df_check = df.copy()

    for column in df_check.columns:

        if df_check[column].dtype == "object":

            df_check[column] = df_check[column].apply(
                lambda x: x.strip()
                if isinstance(x, str)
                else x
            )

    # Remove rows where every cell is blank/NaN
    df = df.loc[
        ~df_check.isna().all(axis=1)
    ]

    return df


# ---------------------------------------------------------
# CLEAN DATA
# ---------------------------------------------------------

def clean_data():

    global input_file

    input_path = input_entry.get().strip()
    output_folder = output_entry.get().strip()

    # Check input file
    if not input_path:

        messagebox.showerror(
            "Missing Input",
            "Please select an Excel input file."
        )

        return

    if not os.path.isfile(input_path):

        messagebox.showerror(
            "Invalid Input",
            "The selected Excel file does not exist."
        )

        return

    # Check output folder
    if not output_folder:

        messagebox.showerror(
            "Missing Output",
            "Please select an output folder."
        )

        return

    if not os.path.isdir(output_folder):

        messagebox.showerror(
            "Invalid Output",
            "The selected output folder does not exist."
        )

        return

    try:

        status_label.config(
            text="Reading Excel file..."
        )

        root.update_idletasks()

        # Read all sheets
        all_sheets = pd.read_excel(
            input_path,
            sheet_name=None
        )

        cleaned_sheets = {}

        # Process every sheet
        for sheet_name, df in all_sheets.items():

            # -----------------------------------------
            # Remove duplicate rows
            # -----------------------------------------

            if duplicate_var.get():

                df = df.drop_duplicates()

            # -----------------------------------------
            # Remove blank rows
            # -----------------------------------------

            if blank_var.get():

                df = remove_blank_rows(df)

            # -----------------------------------------
            # Clean text columns
            # -----------------------------------------

            df = clean_text_columns(df)

            cleaned_sheets[sheet_name] = df

        # ---------------------------------------------
        # Create output filename
        # ---------------------------------------------

        original_name = os.path.splitext(
            os.path.basename(input_path)
        )[0]

        output_file = os.path.join(
            output_folder,
            original_name + "_cleaned.xlsx"
        )

        # ---------------------------------------------
        # Write Excel file
        # ---------------------------------------------

        status_label.config(
            text="Saving cleaned Excel file..."
        )

        root.update_idletasks()

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:

            for sheet_name, df in cleaned_sheets.items():

                # Excel sheet names have a maximum
                # length of 31 characters
                safe_sheet_name = str(sheet_name)[:31]

                df.to_excel(
                    writer,
                    sheet_name=safe_sheet_name,
                    index=False
                )

        status_label.config(
            text="Cleaning completed successfully."
        )

        messagebox.showinfo(
            "Success",
            "Data cleaning completed successfully!\n\n"
            f"Saved file:\n{output_file}"
        )

    except Exception as e:

        status_label.config(
            text="An error occurred."
        )

        messagebox.showerror(
            "Error",
            f"Something went wrong:\n\n{str(e)}"
        )


# ---------------------------------------------------------
# CREATE GUI
# ---------------------------------------------------------

root = tk.Tk()

root.title("Dynamic Data Cleaning Automation")

root.geometry("750x600")

root.resizable(False, False)


# ---------------------------------------------------------
# TITLE
# ---------------------------------------------------------

title_label = tk.Label(
    root,
    text="Dynamic Data Cleaning Automation",
    font=("Arial", 20, "bold")
)

title_label.pack(pady=20)


# ---------------------------------------------------------
# INPUT FILE SECTION
# ---------------------------------------------------------

input_frame = tk.Frame(root)

input_frame.pack(
    fill="x",
    padx=30,
    pady=10
)

input_label = tk.Label(
    input_frame,
    text="Raw Excel File:",
    font=("Arial", 11, "bold")
)

input_label.grid(
    row=0,
    column=0,
    padx=5,
    pady=5,
    sticky="w"
)

input_entry = tk.Entry(
    input_frame,
    width=65
)

input_entry.grid(
    row=0,
    column=1,
    padx=5,
    pady=5
)

input_button = tk.Button(
    input_frame,
    text="Browse",
    width=12,
    command=browse_input_file
)

input_button.grid(
    row=0,
    column=2,
    padx=5
)


# ---------------------------------------------------------
# OUTPUT FOLDER SECTION
# ---------------------------------------------------------

output_label = tk.Label(
    input_frame,
    text="Output Folder:",
    font=("Arial", 11, "bold")
)

output_label.grid(
    row=1,
    column=0,
    padx=5,
    pady=10,
    sticky="w"
)

output_entry = tk.Entry(
    input_frame,
    width=65
)

output_entry.grid(
    row=1,
    column=1,
    padx=5,
    pady=10
)

output_button = tk.Button(
    input_frame,
    text="Browse",
    width=12,
    command=browse_output_folder
)

output_button.grid(
    row=1,
    column=2,
    padx=5
)


# ---------------------------------------------------------
# CLEANING OPTIONS
# ---------------------------------------------------------

options_frame = tk.LabelFrame(
    root,
    text="Data Cleaning Options",
    font=("Arial", 11, "bold"),
    padx=20,
    pady=15
)

options_frame.pack(
    fill="x",
    padx=30,
    pady=15
)


duplicate_var = tk.BooleanVar()

blank_var = tk.BooleanVar()

spaces_var = tk.BooleanVar()

title_var = tk.BooleanVar()


duplicate_check = tk.Checkbutton(
    options_frame,
    text="Remove duplicate rows",
    variable=duplicate_var,
    font=("Arial", 11)
)

duplicate_check.pack(
    anchor="w",
    pady=5
)


blank_check = tk.Checkbutton(
    options_frame,
    text="Remove blank rows",
    variable=blank_var,
    font=("Arial", 11)
)

blank_check.pack(
    anchor="w",
    pady=5
)


spaces_check = tk.Checkbutton(
    options_frame,
    text="Remove leading and trailing spaces from text columns",
    variable=spaces_var,
    font=("Arial", 11)
)

spaces_check.pack(
    anchor="w",
    pady=5
)


title_check = tk.Checkbutton(
    options_frame,
    text="Make Title Text (First Letter Capitalized)",
    variable=title_var,
    font=("Arial", 11)
)

title_check.pack(
    anchor="w",
    pady=5
)


# ---------------------------------------------------------
# BUTTONS
# ---------------------------------------------------------

button_frame = tk.Frame(root)

button_frame.pack(
    pady=20
)


clean_button = tk.Button(
    button_frame,
    text="Clean Data",
    width=18,
    height=2,
    font=("Arial", 11, "bold"),
    command=clean_data
)

clean_button.grid(
    row=0,
    column=0,
    padx=10
)


reset_paths_button = tk.Button(
    button_frame,
    text="Reset Paths",
    width=18,
    height=2,
    font=("Arial", 11),
    command=reset_paths
)

reset_paths_button.grid(
    row=0,
    column=1,
    padx=10
)


reset_checkbox_button = tk.Button(
    button_frame,
    text="Reset Checkboxes",
    width=18,
    height=2,
    font=("Arial", 11),
    command=reset_checkboxes
)

reset_checkbox_button.grid(
    row=0,
    column=2,
    padx=10
)


# ---------------------------------------------------------
# STATUS
# ---------------------------------------------------------

status_label = tk.Label(
    root,
    text="Ready",
    font=("Arial", 10)
)

status_label.pack(
    pady=15
)


# ---------------------------------------------------------
# START GUI
# ---------------------------------------------------------

root.mainloop()