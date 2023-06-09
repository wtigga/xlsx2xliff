import os
import re
import sys
import tkinter as tk
import webbrowser
import xml.dom.minidom as minidom
import xml.etree.ElementTree as ET
from datetime import datetime
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

import pandas as pd

script_version = '0.1'
modification_date = '2023-05-26'
script_name_short = 'Excel2XLIFF'
script_name = str(script_name_short + ', v' + script_version + ', ' + modification_date)

# Get the current timestamp in the format YYYYMMDDhhmmss
timestamp = datetime.now().strftime("%Y%m%d%H%M%S")

# Specify the output file name
output_filename = f"output_xliff_{timestamp}.xliff"

# Get the path to the script's main folder
script_dir = os.path.dirname(os.path.abspath(__file__))

# List of source columns

headers_list = ["Column 1", "Column 2", "Column 3"]
excel_file_path = ''
xliff_file_path = os.path.join(script_dir, output_filename)
# Variable to store the selected source column
source_column = ""
target_column = ""
source_lang_code = 'ZH_CN'
target_lang_code = 'RU_RU'
additional_columns = ['TextId', 'EXTRA']


def select_excel_file():
    global excel_file_path
    excel_file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
    path_entry.delete(0, tk.END)  # Clear the current text in the entry field
    path_entry.insert(tk.END, excel_file_path)  # Insert the selected file path
    print('File selected: ' + str(excel_file_path))


def select_xliff_file():
    global xliff_file_path
    xliff_file_path = filedialog.asksaveasfilename(defaultextension=".xliff", filetypes=[("XLIFF files", "*.xliff")])
    path_entry_xliff.delete(0, tk.END)  # Clear the current text in the entry field
    path_entry_xliff.insert(tk.END, xliff_file_path)  # Insert the selected file path
    print('XLIFF output selected: ' + str(xliff_file_path))


def check():
    messagebox.showinfo("Checked", "Checked")


def update_regex(event):
    global inline_tags_regex
    inline_tags_regex = regex_entry.get("1.0", tk.END).strip()


def toggle_checkbox():
    global perform_inline_tag_replacement
    perform_inline_tag_replacement = checkbox_var.get()
    print('Changed')


def on_source_column_select(event):
    global source_column
    selected_source_column = source_column_combobox.get()
    source_column = selected_source_column
    print('Source column set as ' + selected_source_column)


def on_target_column_select(event):
    global target_column
    selected_target_column = target_column_combobox.get()
    target_column = selected_target_column
    print('Target column set as ' + selected_target_column)


def update_source_column_combobox():
    source_column_combobox['values'] = headers_list


def update_target_column_combobox():
    target_column_combobox['values'] = headers_list


def get_headers():
    global headers_list
    try:
        # Read all sheets into a dictionary of DataFrames
        all_sheets = pd.read_excel(excel_file_path, sheet_name=None, header=0)

        # Combine all sheets into a single DataFrame
        combined_df = pd.concat(all_sheets.values(), ignore_index=True)

        # Get the header as a list
        headers_list = list(combined_df.columns)
        update_source_column_combobox()
        update_target_column_combobox()
        print(headers_list)
    except Exception as e:
        messagebox.showerror("Error", str(e))


def excel_to_xliff():
    global source_column
    global target_column
    global source_lang_code
    global target_lang_code
    global excel_file_path
    global additional_columns

    source_lang_code = source_column
    target_lang_code = target_column
    all_sheets = pd.read_excel(excel_file_path, sheet_name=None, header=0)
    combined_df = pd.concat(all_sheets.values(), ignore_index=True)
    xliff = ET.Element('xliff', version="1.2", xmlns='urn:oasis:names:tc:xliff:document:1.2')
    current_time = datetime.now().strftime("%Y%m%d%H%M%S")
    file_name = os.path.basename(excel_file_path)
    file_elem = ET.SubElement(xliff, 'file', id=current_time, original=file_name, datatype='plaintext',
                              sourceLang=source_lang_code, targetLang=target_lang_code)
    for index, row in combined_df.iterrows():
        source_text = row[source_column]
        target_text = row[target_column]

        # Convert nan values to None
        if pd.isnull(source_text):
            source_text = None
        if pd.isnull(target_text):
            target_text = None

        # Perform inline tag replacement if the flag is True
        if perform_inline_tag_replacement and source_text is not None:
            source_text = re.sub(inline_tags_regex, r'<x id="\g<0>"/>', str(source_text))
        if perform_inline_tag_replacement and target_text is not None:
            target_text = re.sub(inline_tags_regex, r'<x id="\g<0>"/>', str(target_text))

        # Create the trans-unit element
        trans_unit = ET.SubElement(file_elem, 'trans-unit', id=str(index + 1))

        # Create the source element and set the source text if it is not None
        if source_text is not None:
            source_elem = ET.SubElement(trans_unit, 'source')
            source_elem.text = source_text

        # Create the target element and set the target text if it is not None
        if target_text is not None:
            target_elem = ET.SubElement(trans_unit, 'target')
            if target_text.strip().lower() != 'nan':
                target_elem.text = target_text.strip()
                target_elem.set('state', 'translated')
            else:
                # Set the target element as self-closing tag if the text is 'nan'
                target_elem.set('state', 'translated', selfClosing="yes")

        # Create the note element and set the content from additional columns
        note_content = '\n'.join(str(row[col]) for col in additional_columns if col in combined_df.columns)
        if note_content:
            note_elem = ET.SubElement(trans_unit, 'note')
            note_elem.text = note_content
        else:
            print("Warning: No additional columns found or columns not found in the DataFrame.")

        # Create the XML tree
    xml_tree = ET.ElementTree(xliff)

    # Create a string representation of the XML tree with tabs for indentation
    xml_str = ET.tostring(xliff, encoding='utf-8', method='xml').decode()
    dom = minidom.parseString(xml_str)
    xml_str_prettified = dom.toprettyxml(indent="\t")

    # Remove the default namespace prefix added by minidom
    xml_str_prettified = xml_str_prettified.replace('ns0:', '')

    # Save the formatted XML as an XLIFF file
    with open(xliff_file_path, 'w', newline='\n', encoding='utf-8') as file:
        file.write(xml_str_prettified)
    messagebox.showinfo("XLIFF Saved", ("Saved to: " + xliff_file_path) + ", shutting down...")
    sys.exit()

    # Create the main window


window = tk.Tk()
window.title(script_name_short)
window.geometry("700x355")

# Create a button to select the Excel file
select_button = tk.Button(window, text="Select Excel File", command=select_excel_file)
select_button.grid(row=0, column=0, padx=10, pady=10, sticky='w')

# Create a text field to display the selected file path
path_entry = tk.Entry(window, width=80)
path_entry.grid(row=0, column=1, padx=10, pady=10, sticky='w')

# Create a button to select the XLIFF file
select_button = tk.Button(window, text="Save XLIFF to", command=select_xliff_file)
select_button.grid(row=1, column=0, padx=10, pady=10, sticky='w')

# Create a text field to display the selected file path
path_entry_xliff = tk.Entry(window, width=80)
path_entry_xliff.insert(tk.END, xliff_file_path)  # Insert the default value from xliff_file_path
path_entry_xliff.grid(row=1, column=1, padx=10, pady=10, sticky='w')

# Create a button to get headers
check_button = tk.Button(window, text="Read XLSX headers", command=get_headers)
check_button.grid(row=2, column=0, padx=10, pady=10, sticky='w')

# Create a button to save to xliff
check_button = tk.Button(window, text="Save to XLIFF", command=excel_to_xliff)
check_button.grid(row=6, column=0, padx=10, pady=10, sticky='w')

# Create a label for the dropdown menu
source_column_label = tk.Label(window, text="Source Language Column:")
source_column_label.grid(row=3, column=0, padx=10, pady=5, sticky="e")

# Create a dropdown menu for source language column
source_column_combobox = ttk.Combobox(window, values=headers_list, state="readonly")
source_column_combobox.grid(row=3, column=1, padx=10, pady=5, sticky="w")
source_column_combobox.bind("<<ComboboxSelected>>", on_source_column_select)

# Create a label for the target dropdown menu
target_column_label = tk.Label(window, text="Target Language Column:")
target_column_label.grid(row=4, column=0, padx=10, pady=5, sticky="e")

# Create a dropdown menu for target language column
target_column_combobox = ttk.Combobox(window, values=headers_list, state="readonly")
target_column_combobox.grid(row=4, column=1, padx=10, pady=5, sticky="w")
target_column_combobox.bind("<<ComboboxSelected>>", on_target_column_select)

checkbox_var = tk.BooleanVar()
checkbox_var.set(False)  # Default value
checkbox = tk.Checkbutton(window, variable=checkbox_var, command=toggle_checkbox)
checkbox.grid(row=30, column=1, padx=10, pady=5, sticky="w")

# Create a checkbox to toggle inline tag replacement
perform_replace_label = tk.Label(window, text="(EXPERIMENTAL)\nPerform Inline Tag Replacement:")
perform_replace_label.grid(row=30, column=0, padx=10, pady=5, sticky="e")
# Set the initial value of the checkbox variable
perform_inline_tag_replacement = checkbox_var.get()

# Create a text field to input regex
regex_entry = tk.Text(window, height=6, width=60)
regex_entry.grid(row=31, column=1, padx=10, pady=10)
regex_entry.bind("<Return>", update_regex)

# Set the initial value of the regex field
inline_tags_regex = r'(<.+?>)|(%[sdmyY])|({\d})|\((\+{\d})\)|({[A-Z]})|(\[[a-zA-Z0-9_\-]+\])|(\(\+\[[^\]]+\]\)%?)|(\d+\.?\d*%)|(\\n)|(\$\[[\w]+\])|(\bhttps?://\S+)|(\${\w+})|(&lt;t class="t_lc"&gt;)|(&lt;/t&gt;)|@|(\{\w+\})|({SPRITE_PRESET#\d+})'
regex_entry.insert(tk.END, inline_tags_regex)


# Text in the bottom
def open_url(url):
    webbrowser.open(url)


about_label = tk.Label(window, text="github.com/wtigga\nVladimir Zhdanov", fg="blue", cursor="hand2", justify="left")
about_text = tk.Label(window, text=script_name)
about_text.grid(row=25, column=1, sticky='w', padx=20, pady=0)
about_label.bind("<Button-1>",
                 lambda event: open_url("https://github.com/wtigga/xlsx2xliff"))
about_label.grid(row=26, column=1, sticky='w', padx=20, pady=0)

label = tk.Label(window, text="Turns a bilingual *.xlsx into XLIFF"
                              "\nWill close after the conversion."
                              "\nUse at your own risk.", justify="left")
label.grid(row=27, column=1, padx=20, pady=0, sticky='W')

# Run the main event loop
window.mainloop()
