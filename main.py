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
        self.root = root
        open_file_label = CTkLabel(self.root, text="Open File:")
        open_file_label.grid(row=0, column=0, padx=5)
        my_button = CTkButton(self.root, text="Open File", command=self.read_file)
        my_button.grid(row=0, column=1)

    def read_file(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Select file")
        with open(filename, 'rb') as f:
            df = pd.read_excel(f, "תוצאות", index_col=False)

            x_dropdown_text = CTkLabel(root, text="x: ")
            x_dropdown_text.grid(row=1, column=0, padx=5, pady=5)
            x_dropdown = CTkOptionMenu(root, values=df.columns, command=on_selected_x)
            x_dropdown.grid(row=1, column=1, padx=5, pady=5)
            last_row = 2
            for index, column in enumerate(df.columns):
                option = CTkCheckBox(root, text=column)
                option.grid(row=last_row + index, column=0, padx=5, pady=5)
            # group_df = df.groupby(x_dropdown, as_index=False)['עש גב היהלום.1'].mean()
            # Example paired data: before and after measurements
            before = [1.83, 1.83, 1.67, 1.75, 1.82, 1.68, 1.81, 1.84, 1.80, 1.81]
            after = [1.77, 1.79, 1.70, 1.75, 1.80, 1.72, 1.79, 1.83, 1.81, 1.78]

            # Perform paired t-test
            t_statistic, p_value = ttest_rel(before, after)

            print(f"T-statistic: {t_statistic}")
            print(f"P-value: {p_value}")

if __name__ == '__main__':
    root = CTk()
    root.geometry("400x400")
    root.title('Application')
    AgrunomProjectApplication(root)
    root.mainloop()
