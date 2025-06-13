### Import libraries

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

### Functions

## Load excel file and create dataframe
# Handle window close event
def on_window_close(root):
    print("\nOperation canceled. Exiting...")
    root.quit()
    exit()

# Sheet selection
def on_confirm_main(listbox, selected_sheet, sheet_group, root):
    selected_index = listbox.curselection()
    if selected_index:
        selected_sheet[0] = sheet_group[selected_index[0]]
    else:
        messagebox.showinfo("Warning", "Please select an excel sheet to confirm")
    root.quit()

# Display widget for selecting excel sheet, widget is looped till selection is made
def select_sheet(sheet_group):
    # Initialize selection
    selected_sheet = [None]  # Use a list to store the selected sheet name

    while selected_sheet[0] is None:
        root = tk.Tk()
        root.title("Select sheet name")

        # Handle window close event
        root.protocol("WM_DELETE_WINDOW", lambda: on_window_close(root))

        # Dialog window features
        root.geometry("300x200")  # Set a minimum size for the window, min 200 to see confirm button
        frame = tk.Frame(root)
        frame.pack(fill=tk.BOTH, expand=True)  # Ensure the frame is packed
        listbox = tk.Listbox(frame, selectmode=tk.SINGLE)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # List of options
        for sheet in sheet_group:
            listbox.insert(tk.END, sheet)

        #  Confirm button
        tk.Button(root, text="Confirm", command=lambda: on_confirm_main(listbox, selected_sheet, sheet_group, root)).pack()
        root.mainloop()
        root.destroy()

    return selected_sheet[0]

# Select excel file, extract dataframe
def load_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    messagebox.showinfo("Close", "If the excel file you want to process is open, please close it  before proceeding")
    file_name = filedialog.askopenfilename(title="Load the excel file you want to process", filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm")])
    if not file_name:
        print("No file selected. Exiting...")
        exit()

    # Load the selected excel sheet into a dataframe
    try:
        excel_file = pd.ExcelFile(file_name)
        sheets = excel_file.sheet_names
        sheet_name = select_sheet(sheets)

        file_df = pd.read_excel(file_name, sheet_name=sheet_name)
        file_df = file_df.dropna(axis=1, how='all')  # Drop completely empty columns
        print(f"\nLoaded Sheet: {sheet_name}")
        print(file_df.fillna("").to_string(max_cols=None))  # Show all columns and replace nan with empty string
    except Exception as e:
        print(f"\nError: {str(e)}. Exiting...")
        exit()

    return file_name, file_df

### Main

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    ## Load excel file and create dataframes for map and list
    file_name, file_df = load_file()