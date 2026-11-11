import colorsys
import json
import os
import re
import sys
from pathlib import Path
import customtkinter
import openpyxl.drawing.text
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.drawing.colors import ColorChoice
from openpyxl.drawing.text import CharacterProperties, ParagraphProperties, Paragraph
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
import shutil
import tempfile
import zipfile
import pandas as pd
from openpyxl.chart.text import RichText
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook import Workbook


class ScrollableCheckBoxFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, rtl=False, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        self.rtl = rtl
        if rtl:
            self.grid_columnconfigure(0, weight=1)
        self.checkbox_list = []
        self.items = {}  # checkbox -> original (logical-order) value
        for i, item in enumerate(item_list):
            self.add_item(item)

    def add_item(self, item):
        checkbox = customtkinter.CTkCheckBox(self, text=to_display(item), checkbox_width=20,
                                             checkbox_height=20, border_width=2, corner_radius=5)
        if self.command is not None:
            checkbox.configure(command=self.command)
        if self.rtl:
            make_rtl(checkbox)
            checkbox.grid(row=len(self.checkbox_list), column=0, pady=(0, 10), padx=(0, 4), sticky="e")
        else:
            checkbox.grid(row=len(self.checkbox_list), column=0, pady=(0, 10), padx=(4, 0), sticky="w")
        self.checkbox_list.append(checkbox)
        self.items[checkbox] = item

    def set_all(self, checked):
        for checkbox in self.checkbox_list:
            if checked:
                checkbox.select()
            else:
                checkbox.deselect()

    def remove_item(self, item):
        for checkbox in self.checkbox_list:
            if item == self.items[checkbox]:
                checkbox.destroy()
                self.checkbox_list.remove(checkbox)
                del self.items[checkbox]
                return

    def get_checked_items(self):
        return [self.items[checkbox] for checkbox in self.checkbox_list if checkbox.get() == 1]


class ScrollableRadiobuttonFrame(customtkinter.CTkScrollableFrame):
    def __init__(self, master, item_list, command=None, text_variable=None, rtl=False, **kwargs):
        super().__init__(master, **kwargs)

        self.command = command
        self.rtl = rtl
        if rtl:
            self.grid_columnconfigure(0, weight=1)
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
        radiobutton = customtkinter.CTkRadioButton(self, text=to_display(item), value=item,
                                                   variable=self.radiobutton_variable)
        if self.command is not None:
            radiobutton.configure(command=self.command)
        if self.rtl:
            make_rtl(radiobutton)
            radiobutton.grid(row=len(self.radiobutton_list), column=0, pady=(0, 10), padx=(0, 4), sticky="e")
        else:
            radiobutton.grid(row=len(self.radiobutton_list), column=0, pady=(0, 10), padx=(4, 0), sticky="w")
        self.radiobutton_list.append(radiobutton)

    def remove_item(self, item):
        for radiobutton in self.radiobutton_list:
            if item == radiobutton.cget("value"):
                radiobutton.destroy()
                self.radiobutton_list.remove(radiobutton)
                return

    def get_checked_item(self):
        return self.radiobutton_variable.get()


class RadiobuttonGroup(customtkinter.CTkFrame):
    """Non-scrollable radio group for short option lists; same API as ScrollableRadiobuttonFrame."""

    def __init__(self, master, item_list, rtl=False, default=None, **kwargs):
        kwargs.setdefault("fg_color", "transparent")
        super().__init__(master, **kwargs)
        self.grid_columnconfigure(0, weight=1)
        if default not in item_list:
            default = item_list[0] if item_list else ""
        self.radiobutton_variable = customtkinter.StringVar(value=default)
        for i, item in enumerate(item_list):
            radiobutton = customtkinter.CTkRadioButton(self, text=to_display(item), value=item,
                                                       variable=self.radiobutton_variable, radiobutton_width=18,
                                                       radiobutton_height=18, border_width_checked=5,
                                                       border_width_unchecked=2)
            if rtl:
                make_rtl(radiobutton)
            radiobutton.grid(row=i, column=0, sticky="e" if rtl else "w", pady=(0, 8))

    def get_checked_item(self):
        return self.radiobutton_variable.get()


HEBREW_CHARS = re.compile(r"[\u0590-\u05FF]")


def has_hebrew(text):
    return isinstance(text, str) and HEBREW_CHARS.search(text) is not None


# Unicode RIGHT-TO-LEFT EMBEDDING ... POP DIRECTIONAL FORMATTING (invisible)
_RTL_START, _RTL_END = "‫", "‬"


def to_display(text):
    """Prepare text for showing in a Tk widget. Display only: always keep and compare the original value.

    - macOS: Tk lays out Hebrew right-to-left itself, so the text is kept as is and just wrapped in
      an RTL embedding. That makes mixed text like "ימים בציר X" read as one right-to-left sentence
      (otherwise the "X" ends up on the wrong side).
    - Other platforms (Windows): Tk draws each Hebrew word correctly but lays the words out
      left-to-right, so the word order is flipped (the project's original behaviour).
    Text without Hebrew (English, numbers) is never changed.
    """
    if not has_hebrew(text):
        return text
    if sys.platform == "darwin":
        return _RTL_START + text + _RTL_END
    return " ".join(text.split(" ")[::-1])



def make_rtl(widget):
    """Mirror a CTkRadioButton/CTkCheckBox: indicator on the right, text right-aligned to its left.
    (CustomTkinter has no option for this, so swap its internal grid columns; written against 5.2.2.)"""
    widget.grid_columnconfigure(0, weight=1)
    widget.grid_columnconfigure(2, weight=0)
    widget._canvas.grid(row=0, column=2)
    widget._text_label.grid(row=0, column=0, sticky="e")
    widget._text_label["anchor"] = "e"


def sheet_is_rtl(filename, sheet_name, columns=()):
    """True if the sheet is Hebrew: its column headers contain Hebrew, or it's set to
    right-to-left in Excel."""
    if any(has_hebrew(str(column)) for column in columns):
        return True
    try:
        workbook = load_workbook(filename)
        try:
            return bool(workbook[sheet_name].sheet_view.rightToLeft)
        finally:
            workbook.close()
    except Exception:
        return False


def custom_sort_key(val):
    if val == 'UTC':
        return 0, ''  # Make "UTC" come first
    try:
        return 1, int(val)  # Then sort by numeric values
    except ValueError:
        return 2, str(val)  # Fallback to string comparison for other values


def generate_new_file(filename):
    """Return the path of the "(app_generated)" copy the results are written into, creating it if needed.
    The copy is quietly made openable if needed (see make_workbook_writable)."""
    path = Path(filename)
    output_file = path.with_stem(f"{path.stem}(app_generated)")
    if not Path(output_file).is_file():
        output_path = shutil.copy(filename, output_file)
    else:
        output_path = output_file
    make_workbook_writable(output_path)
    return output_path


# Chart layout values openpyxl accepts (Excel itself allows more, e.g. a legend dragged below the plot)
_LAYOUT_LIMITS = {"x": (-1, 1), "y": (-1, 1), "w": (0, 1), "h": (0, 1)}
_MANUAL_LAYOUT = re.compile(r"<(?:\w+:)?manualLayout>.*?</(?:\w+:)?manualLayout>", re.S)
_LAYOUT_VALUE = re.compile(r'<((?:\w+:)?)([xywh]) val="([^"]*)"')


def _clamp_layout_value(match):
    prefix, key, raw = match.groups()
    try:
        value = float(raw)
    except ValueError:
        return match.group(0)
    low, high = _LAYOUT_LIMITS[key]
    clamped = min(max(value, low), high)
    return match.group(0) if clamped == value else f'<{prefix}{key} val="{clamped:g}"'


def _clamp_chart_layouts(xlsx_path):
    """Rewrite the chart parts of an .xlsx so their manual layout values are within openpyxl's limits."""

    def fix_block(block):
        return _LAYOUT_VALUE.sub(_clamp_layout_value, block.group(0))

    with zipfile.ZipFile(xlsx_path) as source:
        entries = [(info, source.read(info)) for info in source.infolist()]
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False, dir=Path(xlsx_path).parent) as tmp:
        tmp_path = tmp.name
    with zipfile.ZipFile(tmp_path, "w") as target:
        for info, data in entries:
            if re.match(r"xl/charts/chart\d+\.xml$", info.filename):
                data = _MANUAL_LAYOUT.sub(fix_block, data.decode("utf-8")).encode("utf-8")
            target.writestr(info, data)
    shutil.move(tmp_path, xlsx_path)


def _can_open(xlsx_path):
    try:
        load_workbook(xlsx_path).close()
        return True
    except Exception:
        return False


def make_workbook_writable(xlsx_path):
    """Make sure openpyxl can open the results copy. Only ever changes that copy, never the user's file."""
    if _can_open(xlsx_path):
        return
    _clamp_chart_layouts(xlsx_path)
    if _can_open(xlsx_path):
        return
    # Last resort: rebuild the copy from the cell values of every sheet
    sheets = pd.read_excel(xlsx_path, sheet_name=None, header=None)
    with pd.ExcelWriter(xlsx_path, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, header=False, index=False)


def generate_excel_formula(df, start_row):
    # Write the DataFrame to the Excel worksheet
    dictOutput = {}
    start_row += 1
    for i, col in enumerate(df.columns[1:]):
        dictOutput[col] = []

    for i, col in enumerate(df.columns[1:]):
        for row in range(len(df) - 1):
            currentRow = start_row + row + 1
            dictOutput[col].append(
                f"=100 - ({chr(ord('A') + i + 1)}{currentRow} / {chr(ord('A') + i + 1)}${start_row} * 100)")
    return dictOutput


def append_df_to_excel(filename, df, sheet_name='Sheet1', start_distance=1, fit_first_column=False):
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
    if startrow == 0:
        start_distance = 2
    df.to_excel(writer, sheet_name=sheet_name, startrow=startrow + start_distance, index=False, header=True)
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
    for row in sheet.iter_rows(min_row=startrow + start_distance + 1, max_row=sheet.max_row, min_col=1,
                               max_col=len(df.columns)):
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

    if fit_first_column:
        # Widen column A so the treatment names fit (never narrows it)
        longest = max(len(str(v)) for v in [df.columns[0], *df.iloc[:, 0]])
        width = max(10, round(longest * 0.9, 1))
        current = sheet.column_dimensions["A"].width or 0
        sheet.column_dimensions["A"].width = max(current, width)
    book.save(filename)
    return startrow + start_distance + 1, len(df.columns)


# ---------- chart colors ----------
# Each chart series (usually a treatment) keeps the same color in every chart and every run.
# New names get the next free color and are saved, so they keep it.
#  - Running from source: series_colors.json next to this file (the one in the repo).
#  - Packaged app (PyInstaller .exe): the repo's file is bundled as the starting colors and copied
#    to the user's app-data folder (e.g. %APPDATA%\AgrunomProject on Windows), where changes persist.
#    The bundle itself is unpacked to a temporary folder that's deleted on exit, so it can't be written to.
BUNDLED_SERIES_COLORS_FILE = Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "series_colors.json"


def user_data_dir():
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    elif sys.platform == "darwin":
        base = Path.home() / "Library" / "Application Support"
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "AgrunomProject"


if getattr(sys, "frozen", False):
    SERIES_COLORS_FILE = user_data_dir() / "series_colors.json"
else:
    SERIES_COLORS_FILE = Path(__file__).with_name("series_colors.json")

DEFAULT_SERIES_COLORS = {
    "UTC": "F672F9",
    "Flupymezotiaz 15": "C55A11",
    "Flupymezotiaz 22": "00B050",
    "Dicloromezotiaz 15": "00B0F0",
    "Dicloromezotiaz 22": "8CF0E9",
    "Broflanilide 18": "7C35B1",
    "Broflanilide 25": "404040",
    "Sparta 80": "ED7D31",
    "Fluxametamide 40": "A5A5A5",
    "Isoflualanam 100": "FFD966",
    "Isoflualanam 200": "2D508F",
    "Isocycloseram 15": "92D050",
    "Isocycloseram 30": "FF0000",
}

# Handed out, in order, to series names that don't have a color yet
EXTRA_SERIES_COLORS = [
    "4472C4", "FFC000", "70AD47", "C00000", "5B9BD5", "7F6000", "548235", "002060", "FF99CC", "00CCCC",
    "996633", "66FF66", "CC99FF", "808000", "FF6600", "008080", "993366", "99CCFF", "333399", "BFBFBF",
]

# Days (when they are the chart series, in "days in legend" mode) are colored by position:
# DOT, then each following date. Not saved to series_colors.json, so they never take treatment colors.
DAY_COLORS = ["4472C4", "FFC000", "70AD47", "C00000", "5B9BD5", "7F6000", "548235", "002060", "FF6600", "993366"]


def colors_for_days(count, start=0):
    return [DAY_COLORS[(start + i) % len(DAY_COLORS)] for i in range(count)]


def _series_key(name):
    return str(name).strip()


def _read_colors(path):
    try:
        with open(path, encoding="utf-8") as f:
            colors = json.load(f)
        return colors if isinstance(colors, dict) else None
    except (OSError, ValueError):
        return None


def load_series_colors():
    colors = _read_colors(SERIES_COLORS_FILE)
    # Starting colors: the bundled file (packaged app), else the built-in defaults.
    # Names the user's file doesn't have yet are added from them (e.g. after an app update).
    starting = _read_colors(BUNDLED_SERIES_COLORS_FILE) or dict(DEFAULT_SERIES_COLORS)
    if colors is None:
        colors = dict(starting)
        save_series_colors(colors)  # first run of the packaged app: create the user's copy
        return colors
    missing = {name: color for name, color in starting.items() if name not in colors}
    if missing:
        colors.update(missing)
        save_series_colors(colors)
    return colors


def save_series_colors(colors):
    try:
        Path(SERIES_COLORS_FILE).parent.mkdir(parents=True, exist_ok=True)
        with open(SERIES_COLORS_FILE, "w", encoding="utf-8") as f:
            json.dump(colors, f, ensure_ascii=False, indent=2)
    except OSError:
        pass  # not writable: colors still apply to this run


def _next_free_color(used):
    for color in EXTRA_SERIES_COLORS:
        if color.upper() not in used:
            return color
    # Palette exhausted: spread further hues around the color wheel (golden-angle steps)
    i = len(used)
    while True:
        r, g, b = colorsys.hsv_to_rgb((i * 0.618033988749895) % 1, 0.75, 0.85)
        color = "%02X%02X%02X" % (int(r * 255), int(g * 255), int(b * 255))
        if color not in used:
            return color
        i += 1


def colors_for_series(names):
    """Return a color (hex, no '#') for each series name, assigning and saving new ones as needed."""
    colors = load_series_colors()
    used = {c.upper() for c in colors.values()}
    changed = False
    for name in names:
        key = _series_key(name)
        if key and key not in colors:
            colors[key] = _next_free_color(used)
            used.add(colors[key].upper())
            changed = True
    if changed:
        save_series_colors(colors)
    return [colors.get(_series_key(name)) for name in names]


def append_chart_to_excel_openpy(y_axis_title, name, filename, startrow, len_df, column, sheet_name, cropped,
                                 flip_x_y=False):
    # Step 1: Load the existing workbook and worksheet
    wb = load_workbook(filename)
    ws = wb[sheet_name]

    # Step 5: Create a BarChart object
    chart = BarChart()
    chart.title = name
    chart.legend.position = 'b'
    # chart.graphical_properties = GraphicalProperties()
    # chart.graphical_properties.line.noFill = True
    # chart.graphical_properties.line.prstDash = None
    chart.width = 21

    if flip_x_y:
        # Plot area pinned to the chart's edges (10% from the left, 30% from the top, 80% x 40%),
        # leaving room for the title above and the legend below. Matches the hand-tuned charts.
        chart.layout = Layout(ManualLayout(xMode="edge", yMode="edge", x=0.1, y=0.3, w=0.8, h=0.4))
        chart.legend.overlay = True
    else:
        chart.layout = Layout(ManualLayout(x=0, y=0, h=0.8, w=0.8))
    # Legend in Excel's default bottom position (a manual offset pushes it out of place)
    chart.legend.layout = None

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
    # Step 6: Define data for the chart
    if not cropped:
        startrow = startrow + 1
    from_rows = False
    if flip_x_y:
        from_rows = True
        data_range = Reference(ws, min_col=1, min_row=startrow, max_col=len_df, max_row=ws.max_row)
        categories = Reference(ws, min_col=2, min_row=startrow - 1, max_row=startrow - 1, max_col=len_df)
    else:
        data_range = Reference(ws, min_col=column + 1, min_row=startrow - 1, max_col=len_df, max_row=ws.max_row)
        categories = Reference(ws, min_col=1, min_row=startrow, max_row=ws.max_row)

    # Add data and categories to the chart
    chart.add_data(data_range, from_rows=from_rows, titles_from_data=True)
    chart.set_categories(categories)
    # Fixed color per series name (same ranges add_data takes the series titles from)
    if from_rows:
        series_names = [ws.cell(row=r, column=1).value for r in range(startrow, ws.max_row + 1)]
    else:
        series_names = [ws.cell(row=startrow - 1, column=c).value for c in range(column + 1, len_df + 1)]
    if from_rows:
        series_colors = colors_for_series(series_names)  # treatments: fixed, saved colors
    else:
        # days: by position; the Decrease(%) chart (cropped) starts after DOT, so skip DOT's color
        series_colors = colors_for_days(len(series_names), start=1 if cropped else 0)
    for series, color in zip(chart.series, series_colors):
        if color:
            series.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.y_axis.title = ""
    if cropped:
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 100
    # y-title text properties
    yt_color = "ff0000"

    xml = f"""
    <txPr>
      <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <a:r>
           <a:rPr b="1" i="0" sz="1000" spc="-1" strike="noStrike">
              <a:solidFill>
                 <a:srgbClr val="{yt_color}" />
              </a:solidFill>
              <a:latin typeface="Calibri" />
           </a:rPr>
           <a:t>{y_axis_title}</a:t>
        </a:r>
      </a:p>
    </txPr>
    """

    chart.y_axis.title.tx.rich = RichText.from_tree(openpyxl.xml.functions.fromstring(xml))
    # Step 7: Insert the chart into the worksheet
    AsciiOfLetter = ord('A')
    number_of_lines_to_add = 2
    if cropped:
        number_of_lines_to_add += 1
    IncrementedLettersAscii = AsciiOfLetter + len_df + number_of_lines_to_add
    ws.add_chart(chart, f"{chr(IncrementedLettersAscii)}{startrow - 1}")
    # Step 8: Save the Excel file
    wb.save(filename)


def write_column_group_labels(filename, sheet_name, row, labels):
    """Write a label (e.g. "Student's" / "Tukey") above each column that has one; None leaves it blank."""
    wb = load_workbook(filename)
    ws = wb[sheet_name]
    font = Font(name='David', size=12)
    alignment = Alignment(horizontal='center')
    for column, text in enumerate(labels, start=1):
        if text:
            cell = ws.cell(row=row, column=column, value=text)
            cell.font = font
            cell.alignment = alignment
    wb.save(filename)


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
                row_with_datetime = pd.to_datetime(col, format='%m/%d/%y')
                if row_with_datetime == pd.NaT:
                    continue
                return index
            except (ValueError, TypeError):
                continue
    else:
        return -1
