from pathlib import Path
import customtkinter
import openpyxl.drawing.text
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.text import CharacterProperties, ParagraphProperties, Paragraph
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.chart.layout import Layout, ManualLayout
import shutil
import pandas as pd
from openpyxl.chart.text import RichText
from openpyxl.utils.dataframe import dataframe_to_rows
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
        return 0, ''  # Make "UTC" come first
    try:
        return 1, int(val)  # Then sort by numeric values
    except ValueError:
        return 2, str(val)  # Fallback to string comparison for other values


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
    return startrow + 2, len(df.columns)


def append_chart_to_excel_openpy(y_axis_title,name, filename, startrow, len_df, column, sheet_name, cropped):
    # Step 1: Load the existing workbook and worksheet
    wb = load_workbook(filename)
    ws = wb[sheet_name]

    # Step 5: Create a BarChart object
    chart = BarChart()
    chart.title = name
    chart.y_axis.title = y_axis_title
    chart.legend.position = 'b'
    chart.layout = Layout(
        ManualLayout(
            x=0, y=0,
            h=0.8, w=0.8
        )
    )

    chart_title_font = openpyxl.drawing.text.Font(typeface='Calibri')
    chart_title_props = CharacterProperties(
        sz=1400,  # Size in EMU (14pt * 100)
        b=True,  # Bold
        u="sng",  # Underline (single)
        solidFill="FF0000"  # Red color in hex (FF0000)
    )

    # Apply the font settings to the title
    chart.title.tx.rich.p[0].r[0].rPr = chart_title_props
    chart.title.tx.rich.p[0].r[0].rPr.latin = chart_title_font

    if not cropped:
        startrow = startrow + 1
    # Step 6: Define data for the chart
    data_range = Reference(ws, min_col=column + 1, min_row=startrow - 1, max_col=len_df, max_row=ws.max_row)
    categories = Reference(ws, min_col=1, min_row=startrow, max_row=ws.max_row)

    # Add data and categories to the chart
    chart.add_data(data_range, titles_from_data=True)
    chart.set_categories(categories)
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    if cropped:
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 100
        chart.y_axis.title = "%"

    else:
        chart.y_axis.title = label
    # Step 7: Insert the chart into the worksheet
    AsciiOfLetter = ord('A')
    IncrementedLettersAscii = AsciiOfLetter + len_df + 2
    ws.add_chart(chart, f"{chr(IncrementedLettersAscii)}{startrow}")
    # Step 8: Save the Excel file
    wb.save(filename)


# def custom_series_factory(values, titles):
#
#     for value in enumerate(values):
#         title = u"{0}!{1}".format(values.sheetname, cell)
#         title = SeriesLabel(strRef=StrRef(title))
#     elif title is not None:
#         title = SeriesLabel(v=title)
#
#     source = NumDataSource(numRef=NumRef(f=values))
#     if xvalues is not None:
#         if not isinstance(xvalues, Reference):
#             xvalues = Reference(range_string=xvalues)
#         series = XYSeries()
#         series.yVal = source
#         series.xVal = AxDataSource(numRef=NumRef(f=xvalues))
#     else:
#         series = Series()
#         series.val = source
#
#     if title is not None:
#         series.title = title
#     return series


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
    font = Font(name='David', size=12, bold=True)
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
