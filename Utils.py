from pathlib import Path
import customtkinter
from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, Alignment
import shutil
import pandas as pd
from openpyxl.workbook import Workbook


class ScrollableCheckBoxFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        self.checkbox_list = []
        for i, item in enumerate(item_list):
            self.add_item(item)

    def add_item(self, item):
        checkbox = customtkinter.CTkCheckBox(self, text=reverse_hebrew_sentence(item))
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


class ScrollableRadiobuttonFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, text_variable=None, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        if text_variable is not None:
            self.radiobutton_variable = text_variable
        else:
            self.radiobutton_variable = customtkinter.StringVar()
            if len(item_list) > 0:
                self.radiobutton_variable.set(item_list[0])

        self.radiobutton_list = []
        for i, item in enumerate(item_list):
            self.add_item(item)

    def add_item(self, item):
        radiobutton = customtkinter.CTkRadioButton(self, text=reverse_hebrew_sentence(item), value=item,
                                                   variable=self.radiobutton_variable)
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
    split = sentence.split(' ')
    return " ".join(split[::-1])


def custom_sort_key(val):
    if val == 'UTC':
        return (0, '')  # Make "UTC" come first
    try:
        return (1, int(val))  # Then sort by numeric values
    except ValueError:
        return (2, str(val))  # Fallback to string comparison for other values


def generate_new_file(filename):
    path = Path(filename)
    output_file = path.with_stem(f"{path.stem}(app_generated)")
    if not Path(output_file).is_file():
        output_path = shutil.copy(filename, output_file)
    else:
        output_path = output_file
    return output_path


def append_df_to_excel(filename, df, sheet_name='Sheet1'):
    # Try to open an existing workbook
    try:
        book = load_workbook(filename)
        writer = pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='overlay')
        writer.workbook = book
        if sheet_name in writer.workbook.sheetnames:
            startrow = writer.workbook[sheet_name].max_row
        else:
            startrow = 0
    except FileNotFoundError:
        # File does not exist yet, create a new one
        writer = pd.ExcelWriter(filename, engine='openpyxl', mode='w')
        startrow = 0
    # Write the dataframe to the Excel file
    df.to_excel(writer, sheet_name=sheet_name, startrow=startrow + 1, index=False, header=True)
    writer.close()
    # Reload the workbook and the sheet to apply formatting
    book = load_workbook(filename)
    sheet = book[sheet_name]

    # Apply font and borders
    font = Font(name='David', size=12)
    border = Border(left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000'))
    alignment = Alignment(horizontal='center')
    dot_index = 0
    if "DOT" in df.columns:
        dot_index = list(df.columns).index("DOT") + 1
    frequency = False
    for col in df.columns:
        if type(col) is str and "שכיחות" in col:
            frequency = True
    # Find the index of the 'DOT' column
    for row in sheet.iter_rows(min_row=startrow + 2, max_row=sheet.max_row, min_col=1, max_col=len(df.columns)):
        for cell in row:
            cell.font = font
            cell.border = border
            cell.alignment = alignment
            if frequency:
                cell.number_format = '0.00%'
            elif cell.column == dot_index:
                cell.number_format = '0.0'
            else:
                cell.number_format = '0.00'  # 2 decimal places for other columns

    book.save(filename)


def write_text_to_excel(filename, text, sheet_name='Sheet1'):
    # Try to open an existing workbook
    try:
        wb = load_workbook(filename)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.create_sheet(sheet_name)
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
    font = Font(name='David', size=12,bold=True)
    # Find the next empty row
    start_row = ws.max_row + 2
    # Write the text to the first cell in the next empty row
    cell = ws.cell(row=start_row, column=1, value=text)
    cell.font = font
    wb.save(filename)
# Function to find the first row with the date format
def find_first_date_row(df):
    for index, row in df.iterrows():
        for col in row:
            try:
                pd.to_datetime(col, format='%m/%d/%y')
                return index
            except (ValueError, TypeError):
                continue
    else:
        return -1