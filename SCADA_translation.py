### Import libraries

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from googletrans import Translator

### Functions

## Load excel file and create dataframe
# Handle window close event
def on_window_close(root):
    print("\nOperation canceled. Exiting...")
    root.quit()
    exit()

# Name selection
def on_confirm_main(listbox, selected_name, names, root):
    selected_index = listbox.curselection()
    if selected_index:
        selected_name[0] = names[selected_index[0]]
    else:
        messagebox.showinfo("Warning", "Please select a proper name to confirm")
    root.quit()

# Display widget for selecting a name on a list, widget is looped till selection is made
def select_name(names, message):
    # Initialize selection
    selected_name = [None]  # Use a list to store the selected name

    while selected_name[0] is None:
        root = tk.Tk()
        root.title(f"Select the {message}")

        # Handle window close event
        root.protocol("WM_DELETE_WINDOW", lambda: on_window_close(root))

        # Dialog window features
        root.geometry("400x200")  # Set a minimum size for the window, min 200 to see confirm button
        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True)  # Ensure the frame is packed
        listbox = tk.Listbox(frame, selectmode=tk.SINGLE)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # List of options
        for name in names:
            listbox.insert(tk.END, name)

        #  Confirm button
        tk.Button(root, text="Confirm", command=lambda: on_confirm_main(listbox, selected_name, names, root)).pack()
        root.mainloop()
        root.destroy()

    return selected_name[0]

# Select excel file, sheet and colum, extract dataframe
def load_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo("Close", "If the excel file you want to process is open, please close it before proceeding")
    file_name = filedialog.askopenfilename(title="Load the excel file you want to process", filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm")])
    if not file_name:
        print("No file selected. Exiting...")
        exit()

    # Load the selected excel sheet into a dataframe
    try:
        excel_file = pd.ExcelFile(file_name)
        sheets = excel_file.sheet_names
        sheet_name = select_name(sheets, "excel sheet")
        cols = excel_file.parse(sheet_name).columns.tolist()
        src_col = select_name(cols, "source column")
        trans_col = select_name(cols, "translated coulmn")

        file_df = pd.read_excel(file_name, sheet_name=sheet_name, usecols=[src_col, trans_col]).dropna(how="any")

        print(f"\n{file_df.fillna("").to_string}")  # Replace nan with empty string
    except Exception as e:
        print(f"\nError: {str(e)}. Exiting...")
        exit()

    return file_name, file_df, src_col, trans_col

## Translate col_origin to col_translated
# Apply the mapping to the origin column with fallback to Google Translate
def translate_value(value, translation_map):
    if value in translation_map and pd.notna(translation_map[value]):
        return translation_map[value]  # Use dictionary translation
    else:
        try:
            return translator.translate(value, src=src_lang, dest=dest_lang).text  # Use Google Translate
        except Exception as e:
            print(f"Translation error for '{value}': {str(e)}")
            return value  # Return the original value if translation fails

#
def translate_column(file_df, src_col, trans_col):
    # Create a mapping from origin to translated
    translation_map = dict(zip(file_df[col_origin], file_df[col_translated]))
    translator = Translator()

    # Select languages for translation
    src_lang = select_name(["en", "es", "fr", "de", "it", "pt"], "source language")
    trans_lang = select_name(["en", "es", "fr", "de", "it", "pt"], "translated language")

    file_df[col_origin] = file_df[col_origin].apply(translate_value)

    return file_df
dfd
### Main

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    ## Load excel file and create dataframes for map and list
    messagebox.showinfo("Empty", "Words present in the translated columns will be kept. Please make sure to delete unwanted cells before proceeding")
    file_df, src_col, trans_col = load_file()

    file_df = translate_column(file_df, src_col, trans_col)