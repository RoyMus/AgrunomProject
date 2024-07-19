from pathlib import Path
from customtkinter import *
import re
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
import shutil
import pandas as pd


class ScrollableCheckBoxFrame(CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        self.checkbox_list = []
        for i, item in enumerate(item_list):
            self.add_item(item)

    def add_item(self, item):
        checkbox = CTkCheckBox(self, text=reverse_hebrew_sentence(item))
        if self.command is not None:
            checkbox.configure(command=self.command)
        checkbox.grid(row=len(self.checkbox_list), column=0, pady=(0, 10))
        self.checkbox_list.append(checkbox)

    def remove_item(self, item):
        for checkbox in self.checkbox_list:
            if item == checkbox.cget("text"):
                checkbox.destroy()
                self.checkbox_list.remove(checkbox)
                return

    def get_checked_items(self):
        return [reverse_hebrew_sentence(checkbox.cget("text")) for checkbox in self.checkbox_list if
                checkbox.get() == 1]


class ScrollableRadiobuttonFrame(CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, text_variable=None, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        if text_variable is not None:
            self.radiobutton_variable = text_variable
        else:
            self.radiobutton_variable = StringVar()

        self.radiobutton_list = []
        for i, item in enumerate(item_list):
            self.add_item(item)

    def add_item(self, item):
        radiobutton = CTkRadioButton(self, text=reverse_hebrew_sentence(item), value=item, variable=self.radiobutton_variable)
        if self.command is not None:
            radiobutton.configure(command=self.command)
        radiobutton.grid(row=len(self.radiobutton_list), column=0, pady=(0, 10))
        self.radiobutton_list.append(radiobutton)

    def remove_item(self, item):
        for radiobutton in self.radiobutton_list:
            if item == radiobutton.cget("text"):
                radiobutton.destroy()
                self.radiobutton_list.remove(radiobutton)
                return

    def get_checked_item(self):
        return self.radiobutton_variable.get()


def reverse_hebrew_sentence(sentence):
    # Regular expression to match Hebrew words, possibly followed by numbers or special characters
    pattern = re.compile(r'([א-ת]+(?:\d+|[^\sא-ת]*)*)')

    # Find all matches
    words_numbers = pattern.findall(sentence)

    # Reverse the list of matches
    reversed_sentence = ' '.join(words_numbers[::-1])

    return reversed_sentence


def custom_sort_key(val):
    if val == 'UTC':
        return (0, '')  # Make "UTC" come first
    return (1, int(val))  # Then sort by numeric values


def append_df_to_excel(filename, df, sheet_name='Sheet1'):
    path = Path(filename)
    output_file = path.with_stem(f"{path.stem}(app_generated)")
    if not Path(output_file).is_file():
        output_path = shutil.copy(filename, output_file)
    else:
        output_path = output_file
    # Try to open an existing workbook
    try:
        book = load_workbook(output_path)
        writer = pd.ExcelWriter(output_path, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        writer.workbook = book
        if sheet_name in writer.workbook.sheetnames:
            startrow = writer.workbook[sheet_name].max_row
        else:
            startrow = 0
    except FileNotFoundError:
        # File does not exist yet, create a new one
        writer = pd.ExcelWriter(output_path, engine='openpyxl', mode='w')
        startrow = 0
    # Write the dataframe to the Excel file
    df.to_excel(writer, sheet_name=sheet_name, startrow=startrow + 1, index=False, header=True)
    writer.close()
    # Reload the workbook and the sheet to apply formatting
    book = load_workbook(output_path)
    sheet = book[sheet_name]

    # Apply font and borders
    font = Font(name='David', size=12)
    border = Border(left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000'))
    alignment = Alignment(horizontal='center')
    dot_index = list(df.columns).index("DOT") + 1
    # Find the index of the 'DOT' column
    for row in sheet.iter_rows(min_row=startrow + 2, max_row=sheet.max_row, min_col=1, max_col=sheet.max_column):
        for cell in row:
            cell.font = font
            cell.border = border
            cell.alignment = alignment
            if cell.column == dot_index:
                cell.number_format = '0.0'
            else:
                cell.number_format = '0.00'  # 2 decimal places for other columns

    book.save(output_path)
