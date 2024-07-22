import os
from tkinter import messagebox, filedialog
from tkinter import ttk
import statsmodels.api as sm
from scipy import stats
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import CalcUtils
import customtkinter
import Utils
import pandas as pd
from pathlib import Path


class AgrunomProjectApplication(customtkinter.CTkFrame):
    def __init__(self, root):
        self.filename = None
        customtkinter.set_appearance_mode("light")  # Modes: system (default), light, dark
        customtkinter.set_default_color_theme("green")
        # Themes: blue (default), dark-blue, green
        customtkinter.CTkFrame.__init__(self, root)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)
        root.rowconfigure(1, weight=1)
        self.upper_frame = customtkinter.CTkFrame(root, fg_color="transparent")
        self.upper_frame.grid(row=0, column=0, columnspan=2)
        self.scrollable_checkbox_frame = None
        self.value_inside = None
        self.input_df = None
        self.output_df = None
        self.root = root
        read_file_button = customtkinter.CTkButton(self.upper_frame, text="Open File", command=self.read_file)
        read_file_button.grid(row=0, column=0, columnspan=2, pady=10)
        # create tabview
        pd.options.display.float_format = '${:,.2f}'.format
        self.tabview = customtkinter.CTkTabview(self.upper_frame, width=250)
        self.tabview.add("X")
        self.tabview.add("Y")
        self.tabview.add("Block")
        self.tabview.tab("X").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Y").grid_columnconfigure(0, weight=1)
        self.tabview.tab("Block").grid_columnconfigure(0, weight=1)
        self.excel_frame = customtkinter.CTkScrollableFrame(self.root, label_text="Excel Data",
                                                            orientation="horizontal")
        self.tv1 = ttk.Treeview(self.excel_frame)
        self.tv1.grid(row=0, column=0, sticky="nsew")
        self.excel_frame.columnconfigure(0, weight=1)
        self.excel_frame.rowconfigure(0, weight=1)

    def calculate(self, tukey=False):
        x_value = self.x.get_checked_item()
        block = self.block_selection.get_checked_item()
        items = self.scrollable_checkbox_frame.get_checked_items()

        if len(items) == 0:
            messagebox.showerror("No Y chosen", "Please choose at least one Y value")
            return
        self.output_dict = {}
        output_columns = ["DOT", "DOT sig", "3DAT", "3DAT sig", "7DAT", "7DAT sig", "10DAT", "10DAT sig", "14DAT",
                          "14DAT sig"]
        output_path = None
        for label in items:
            char = ''
            index_of_row = 0
            for col in self.input_df.columns:
                if label != col.split('.')[0]:
                    continue
                df = pd.DataFrame(
                    {'Treatment': self.input_df[x_value], 'Value': pd.to_numeric(self.input_df[col], errors='coerce')})
                df['Treatment'] = df['Treatment'].astype(str)
                if block != "ללא":
                    df["Block"] = self.input_df[block].astype('category')
                    model = ols('Value ~ C(Treatment) + C(Block)', data=df).fit()
                else:
                    model = ols('Value ~ C(Treatment)', data=df).fit()

                # Perform ANOVA
                anova_table = sm.stats.anova_lm(model, typ=2)
                sigByKey = None
                # Calculate means for each treatment
                treatment_means = df.groupby('Treatment')['Value'].mean().to_dict()
                treatment_means = {r: treatment_means[r] for r in
                                   sorted(treatment_means, key=treatment_means.get, reverse=True)}

                if tukey:
                    # Perform Tukey HSD test
                    tukey = pairwise_tukeyhsd(endog=df['Value'],
                                              groups=df['Treatment'],
                                              alpha=0.05)
                    tukey_df = pd.DataFrame(data=tukey._results_table.data[1:], columns=tukey._results_table.data[0])
                    tukey_dictionary = dict()
                    for treatment in treatment_means:
                        tukey_dictionary[treatment] = dict()
                    for index, row in tukey_df.iterrows():
                        tukey_dictionary[row["group1"]][row["group2"]] = row["reject"] == True
                        tukey_dictionary[row["group2"]][row["group1"]] = row["reject"] == True

                    sigByKey = CalcUtils.calculate_significant_letters_tukey(tukey_dictionary)
                else:

                    # Calculate Mean Squared Error (MSE)
                    mse = anova_table['sum_sq']['Residual'] / anova_table['df']['Residual']

                    # Number of samples per group (assuming equal sample size)
                    n = df['Treatment'].value_counts().min()  # Use the minimum count in case of unequal sizes

                    # Degrees of freedom for error
                    df_error = anova_table['df']['Residual']

                    # Significance level
                    alpha = 0.05

                    # Critical t-value
                    t_critical = stats.t.ppf(1 - alpha / 2, df_error)
                    sigByKey = CalcUtils.calculate_significant_letters(treatment_means, t_critical, mse, n)
                means = {i: treatment_means[i] for i in sorted(treatment_means, key=lambda x: Utils.custom_sort_key(x))}
                if index_of_row >= len(output_columns):
                    index_of_row = 0
                    char += str(1)
                self.output_dict[output_columns[index_of_row] + char] = list(means.values())
                self.output_dict[output_columns[index_of_row + 1] + char] = ["".join(sorted(sigByKey[x])) for x in
                                                                             means.keys()]
                index_of_row += 2
            self.output_df = pd.DataFrame(self.output_dict)
            self.output_df.insert(0, label, means.keys())
            # self.output_df = self.output_df.sort_values(by=label, key=lambda col: col.map(custom_sort_key))
            output_path = Utils.append_df_to_excel(self.filename, self.output_df, "טבלאות")
        result = messagebox.askokcancel("Calculation finished", f"Do you want to open the file? ")
        if result:
            os.system(f'explorer /select,"{output_path}"')

    def init_ui(self, sheet_name):
        self.newWindow.withdraw()
        with open(self.filename, 'rb') as f:
            self.excel_frame._label.configure(text=Utils.reverse_hebrew_sentence(Path(self.filename).stem))
            self.input_df = pd.read_excel(f, sheet_name, index_col=False)
            columns = list(self.input_df.columns)
            self.tv1["column"] = columns
            self.tv1["show"] = "headings"
            for column in self.tv1["column"]:
                self.tv1.heading(column, text=column)
            df_rows = self.input_df.to_numpy().tolist()
            for row in df_rows:
                self.tv1.insert("", "end", values=row)
            self.tabview.grid(row=1, column=0, columnspan=2)
            unduplicatedcolumns = []
            for col in columns:
                if col.split('.')[0] not in unduplicatedcolumns:
                    unduplicatedcolumns.append(col)
            self.x = Utils.ScrollableRadiobuttonFrame(master=self.tabview.tab("X"), item_list=unduplicatedcolumns)
            self.x.grid(row=0, column=0)
            self.scrollable_checkbox_frame = Utils.ScrollableCheckBoxFrame(master=self.tabview.tab("Y"),
                                                                           item_list=unduplicatedcolumns)
            self.scrollable_checkbox_frame.grid(row=0, column=0)
            calc_button = customtkinter.CTkButton(self.upper_frame, text="Calculate Each Pair Student's T's",
                                                  command=self.calculate)
            blocks = unduplicatedcolumns
            blocks.insert(0, "ללא")
            self.block_selection = Utils.ScrollableRadiobuttonFrame(master=self.tabview.tab("Block"), item_list=blocks)
            self.block_selection.grid(row=0, column=0)
            calc_button.grid(row=2, column=0, pady=(15, 0), padx=10)
            self.excel_frame.grid(row=1, column=0, sticky="nsew", columnspan=2)
            tukey_calc = customtkinter.CTkButton(self.upper_frame, text="Calculate Each Pair Tukey ",
                                                 command=lambda: self.calculate(tukey=True))
            tukey_calc.grid(row=2, column=1, pady=(15, 0), padx=10)

    def read_file(self):
        self.filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                                   filetypes=[("Excel files", ".xlsx .xls")])
        if not self.filename:
            messagebox.showerror("File Not Selected", "You didn't select a file")
            return

        xls = pd.ExcelFile(self.filename)
        # Now you can list all sheets in the file
        self.newWindow = customtkinter.CTkToplevel(self.root)
        self.newWindow.title("Choose sheet name")
        self.newWindow.geometry("400x400")
        self.newWindow.geometry("+%d+%d" % (self.root.winfo_x() + 50, self.root.winfo_y() + 50))
        self.newWindow.attributes('-topmost', 'true')
        self.newWindow.columnconfigure(0, weight=1)
        self.newWindow.columnconfigure(1, weight=1)
        self.newWindow.rowconfigure(0, weight=1)
        self.newWindow.rowconfigure(1, weight=1)
        sheets = Utils.ScrollableRadiobuttonFrame(self.newWindow, item_list=xls.sheet_names)
        sheets.grid(row=0, column=0, columnspan=2)
        confirm = customtkinter.CTkButton(self.newWindow, text="Confirm",
                                          command=lambda: self.init_ui(sheets.get_checked_item()))
        confirm.grid(row=1, column=0, columnspan=2)
        self.newWindow.focus()

    def clear_data(self):
        self.tv1.delete(*self.tv1)


if __name__ == '__main__':
    root = customtkinter.CTk()
    root.geometry("500x700")
    root.title('AgronumProject')
    AgrunomProjectApplication(root)
    root.focus()
    root.mainloop()
