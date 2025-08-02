### Import libraries

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import asyncio
from googletrans import Translator
import nltk
nltk.download('wordnet')
from nltk.corpus import wordnet
from tqdm import tqdm
import openpyxl
import os
import shutil

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

        file_df = pd.read_excel(file_name, sheet_name=sheet_name, usecols=[src_col, trans_col])
        print(f"\n{file_df.fillna("").to_string(max_rows=20)}")  # Replace nan with empty string
    except Exception as e:
        print(f"\nError: {str(e)}. Exiting...")
        exit()

    return file_name, sheet_name, file_df, cols, src_col, trans_col

## Translate col_origin to col_translated
# Get a synonym for a word using WordNet
def get_synonym(word):
    # Try to get a synonym using WordNet
    synsets = wordnet.synsets(word)
    for syn in synsets:
        for lemma in syn.lemmas():
            synonym = lemma.name().replace('_', ' ')
            if synonym.lower() != word.lower():
                return synonym
    return word

# Abbreviate a word
def abbreviate_word(word):
    # Abbreviate words longer than 5 characters: keep first 4 letters, add a dot
    if len(word) > 5:
        return word[:4] + '.'
    return word

# Shorten the translation using synonyms or abbreviations
def shorten_translation(original, translated):
    words = translated.split()
    # 1. Try to shorten with synonyms
    new_words = []
    for w in words:
        synonym = get_synonym(w)
        if len(synonym) < len(w):
            new_words.append(synonym)
        else:
            new_words.append(w)
    # 2. If still too long, abbreviate longest words as needed
    word_forms = [(w, w) for w in new_words]
    sorted_indices = sorted(range(len(new_words)), key=lambda i: -len(new_words[i]))
    shortened = ' '.join(w for _, w in word_forms)
    for idx in sorted_indices:
        if len(shortened) <= len(original):
            break
        orig, _ = word_forms[idx]
        abbr = abbreviate_word(orig)
        word_forms[idx] = (orig, abbr)
        shortened = ' '.join(w for _, w in word_forms)
    return shortened

# Apply the translation mapping with fallback to Google Translate
async def translate_value(value, translation_map, src_lang, trans_lang):
    translator = Translator()

    # Skip translation for empty cells
    if pd.isna(value) or value == "":
        return value  # Return the original value if it's empty

    # Check if the value is in the translation map
    if value in translation_map and pd.notna(translation_map[value]):
        return translation_map[value]  # Use dictionary translation
    else:
        # If not in the map, use Google Translate
        try:
            translation = await translator.translate(value, src=src_lang, dest=trans_lang)
            translated_text = translation.text

            # Enforce length limit
            if len(translated_text) > len(value):
                translated_text = shorten_translation(value, translated_text)

            return translated_text

        except Exception as e:
            print(f"Translation error for '{value}': {str(e)}")
            return value  # Return the original value if translation fails

# Apply translation only to rows where src_col matches a shortlisted value
def apply_translation(val, restored_translation_map):
    if val in restored_translation_map:
        return restored_translation_map[val]
    return val

# Translate the specified column in the dataframe
async def translate_column_async(file_df, src_col, trans_col):
    word_pattern = re.compile(r'\b\w+\b')
    tag_pattern = re.compile(r'\b([A-Za-z]+-\d+)\b')
    # This pattern matches your tag only when it is not immediately preceded or followed by a letter or digit
    #tag_pattern = re.compile(r'(?<![A-Za-z0-9])([A-Za-z]+-\d+)(?![A-Za-z0-9])')

    # Extract unique values from the source column
    unique_values = file_df[src_col].dropna().unique()

    # Shortlist: only values that match the pattern
    shortlist = [
        val for val in unique_values
        if isinstance(val, str) and any(word.isalpha() for word in word_pattern.findall(val))
    ]

    # Map: original value -> (value_with_tag_placeholder, [tags])
    tag_map = {}
    for val in shortlist:
        tags = tag_pattern.findall(val)
        val_with_tag = tag_pattern.sub('***', val)
        tag_map[val] = (val_with_tag, tags)

    # Build a set of unique values with tags replaced
    shortlist_with_tags = list({v[0] for v in tag_map.values()})
    print(f"\nShortlisted values for translation ({len(shortlist_with_tags)}):")
    for val in shortlist_with_tags:
        print(f"{val}")

    # Translation map for values
    translation_map = {}

    # Select languages for translation
    languages = ["en", "es", "fr", "de", "it", "pt"]
    src_lang = select_name(languages, "source language")
    trans_lang = select_name(languages, "translated language")

    # Translate only shortlisted values
    for value in tqdm(shortlist_with_tags, desc="Translating shortlisted values"):
        translation_map[value] = await translate_value(value, {}, src_lang, trans_lang)

    # Restore tags in translated values
    restored_translation_map = {}
    for orig_val, (val_with_tag, tags) in tag_map.items():
        translated = translation_map[val_with_tag]
        # Replace each *TAG* with the original tag, in order
        for tag in tags:
            translated = translated.replace('***', tag, 1)
        restored_translation_map[orig_val] = translated

    print("\nRestored translation map:")
    for orig_val, translated in restored_translation_map.items():
        print(f"{orig_val} -> {translated}")

    # Apply the translation to the dataframe
    file_df[trans_col] = file_df[src_col].apply(lambda val: apply_translation(val, restored_translation_map))

    print(f"\n{file_df.fillna('').to_string(max_rows=20)}")
    return file_df

## Write to excel file
# Write translated column to excel file
def write_trans_col(file_name, sheet_name, file_df, cols, trans_col):
    try:
        # Create a copy of the original file
        base, ext = os.path.splitext(file_name)
        copied_file_name = base + "_translated" + ext
        shutil.copy(file_name, copied_file_name)

        # Use keep_vba only for .xlsm files
        if ext.lower() == ".xlsm":
            wb = openpyxl.load_workbook(copied_file_name, keep_vba=True)
        else:
            wb = openpyxl.load_workbook(copied_file_name)

        if sheet_name not in wb.sheetnames:
            print(f"\nSheet {sheet_name} not found in the workbook.Exiting...")
            exit()

        ws = wb[sheet_name]

        # Write the translated column to the sheet
        for r_idx, value in enumerate(file_df[trans_col], start=1):  # Write values from the translated column
            ws.cell(row=r_idx + 1, column=cols.index(trans_col) + 1, value=value)  # Write to the column defined by trans_col

        wb.save(copied_file_name)
        wb.close()

    except Exception as e:
        print(f"\nError writing to Excel: {str(e)}. Exiting...")
        exit()

### Main

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    # Load excel file and create dataframes for map and list
    messagebox.showinfo("Empty", "Words present in the translated columns will be kept. Please make sure to delete unwanted cells before proceeding")
    file_name, sheet_name, file_df, cols, src_col, trans_col = load_file()

    # Run the asynchronous translation
    file_df = asyncio.run(translate_column_async(file_df, src_col, trans_col))

    # Write the translated column to the excel file
    write_trans_col(file_name, sheet_name, file_df, cols, trans_col)