import numpy as np
from scipy.stats import ttest_rel
import pandas as pd
from customtkinter import *
from tkinter import filedialog

selected_y = ["בוגר"]
selected_x = ["טיפול"]
root = CTk()
root.title('Application')


def read_file():
    filename = filedialog.askopenfilename(initialdir="/", title="Select file")
    with open(filename, 'rb') as f:
        df = pd.read_excel(f, "תוצאות",index_col=False)
        group_df = df.groupby('טיפול', as_index=False)['עש גב היהלום.1'].mean()
        print(df)

        # Example paired data: before and after measurements
        before = [1.83, 1.83, 1.67, 1.75, 1.82, 1.68, 1.81, 1.84, 1.80, 1.81]
        after = [1.77, 1.79, 1.70, 1.75, 1.80, 1.72, 1.79, 1.83, 1.81, 1.78]

        # Perform paired t-test
        t_statistic, p_value = ttest_rel(before, after)

        print(f"T-statistic: {t_statistic}")
        print(f"P-value: {p_value}")


my_button = CTkButton(root, text="Open File", command=read_file)
my_button.pack()
root.mainloop()


