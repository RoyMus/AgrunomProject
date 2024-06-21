import numpy as np
from scipy.stats import ttest_rel
import pandas as pd
from customtkinter import *
from tkinter import filedialog

selected_y = []
selected_x = ""


def on_selected_x(choice):
    global selected_x
    selected_x = choice


class AgrunomProjectApplication(CTkFrame):
    def __init__(self, root):
        CTkFrame.__init__(self, root)
        self.value_inside = None
        self.df = None
        self.root = root
        self.checkboxes_vars = []
        open_file_label = CTkLabel(self.root, text="Open File:")
        open_file_label.grid(row=0, column=0, padx=5)
        my_button = CTkButton(self.root, text="Open File", command=self.read_file)
        my_button.grid(row=0, column=1)

    def print_table(self):
        for label,var in self.checkboxes_vars:
            if var.get() == 1:
                group_df = self.df.groupby(self.value_inside.get(), as_index=False)[label].mean()
                print(group_df)

    def read_file(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Select file")
        with open(filename, 'rb') as f:
            self.df = pd.read_excel(f, "תוצאות", index_col=False)
            self.value_inside = StringVar(root)
            self.value_inside.set(self.df.columns[0])
            x_dropdown_text = CTkLabel(root, text="x: ")
            x_dropdown_text.grid(row=1, column=0, padx=5, pady=5)
            x_dropdown = CTkOptionMenu(root, variable=self.value_inside, values=self.df.columns, command=on_selected_x)
            x_dropdown.grid(row=1, column=1, padx=5, pady=5)
            last_row = 2
            for index, column in enumerate(self.df.columns):
                var = IntVar()
                option = CTkCheckBox(root,variable=var, text=column)
                option.grid(row=last_row, column=index, padx=5, pady=5)
                self.checkboxes_vars.append((column,var))
            calc_button = CTkButton(self.root, text="Calculate", command=self.print_table)
            calc_button.grid(row=last_row + 1, column=0)


if __name__ == '__main__':
    root = CTk()
    root.geometry("400x400")
    root.title('Application')
    AgrunomProjectApplication(root)
    root.mainloop()
