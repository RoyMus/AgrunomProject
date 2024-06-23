import numpy as np
from scipy.stats import ttest_rel
import pandas as pd
from customtkinter import *
from tkinter import filedialog
from CalcUtils import *
from itertools import combinations
import statsmodels.api as sm
from statsmodels.formula.api import ols
import numpy as np
from scipy import stats

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
        x_value = self.value_inside.get()
        for label, var in self.checkboxes_vars:
            if var.get() == 1:
                df = pd.DataFrame({'Treatment': self.df[x_value], 'Value': self.df[label]})
                print("DataFrame:")
                print(df)
                df['Value'] = pd.to_numeric(df['Value'])
                df['Treatment'] = df['Treatment'].astype(str)
                # Perform ANOVA
                model = ols('Value ~ C(Treatment)', data=df).fit()
                anova_table = sm.stats.anova_lm(model, typ=2)
                print("\nANOVA Table:")
                print(anova_table)

                # Calculate Mean Squared Error (MSE)
                mse = anova_table['sum_sq']['Residual'] / anova_table['df']['Residual']
                print(f'\nMean Squared Error (MSE): {mse}')

                # Number of samples per group (assuming equal sample size)
                n = df['Treatment'].value_counts().min()  # Use the minimum count in case of unequal sizes

                # Degrees of freedom for error
                df_error = anova_table['df']['Residual']

                # Significance level
                alpha = 0.05

                # Critical t-value
                t_critical = stats.t.ppf(1 - alpha / 2, df_error)

                # Calculate LSD
                lsd = t_critical * np.sqrt(2 * mse / n)
                se_diff = lsd / t_critical
                print(f'\nLeast Significant Difference (LSD): {lsd}')
                # Dictionary to store the results
                treatment_significant_dict = {treatment: [] for treatment in df['Treatment'].unique()}

                # Calculate means for each treatment
                treatment_means = df.groupby('Treatment')['Value'].mean().to_dict()

                def get_standard_error_diff():
                    return np.sqrt(2 * mse / n)

                # Function to get the t-statistic and check if it's a critical difference
                def get_t_statistic(treatment1, treatment2):
                    mean1 = df[df['Treatment'] == treatment1]['Value'].mean()
                    mean2 = df[df['Treatment'] == treatment2]['Value'].mean()
                    sed = get_standard_error_diff()
                    t_stat = abs(mean1 - mean2) / sed
                    is_critical = t_stat > t_critical
                    return t_stat, is_critical

                # Populate the dictionary with comparisons
                for treatment1, treatment2 in combinations(treatment_means.keys(), 2):
                    _, is_significant = get_t_statistic(treatment1, treatment2)
                    treatment_significant_dict[treatment1].append((treatment2, is_significant))
                    treatment_significant_dict[treatment2].append((treatment1, is_significant))


    def read_file(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Select file")
        with open(filename, 'rb') as f:
            self.df = pd.read_excel(f, "תוצאות", index_col=False)
            self.df = self.df.drop(self.df.index[-1])
            self.value_inside = StringVar(root)
            self.value_inside.set(self.df.columns[0])
            x_dropdown_text = CTkLabel(root, text="x: ")
            x_dropdown_text.grid(row=1, column=0, padx=5, pady=5)
            x_dropdown = CTkOptionMenu(root, variable=self.value_inside, values=self.df.columns, command=on_selected_x)
            x_dropdown.grid(row=1, column=1, padx=5, pady=5)
            last_row = 2
            line_break = 0
            current_col = 0
            break_distance = 0
            n_cols = len(self.df.columns)
            for i in range(n_cols - 1, 2, -1):
                if n_cols % i == 0:
                    break_distance = i
                    break
            for index, column in enumerate(self.df.columns):
                var = IntVar()
                if index % break_distance == 0 and index != 0:
                    line_break += 1
                    current_col = 0
                option = CTkCheckBox(root, variable=var, text=column)
                option.grid(row=last_row + line_break, column=current_col, padx=5, pady=5)
                self.checkboxes_vars.append((column, var))
                current_col += 1
            calc_button = CTkButton(self.root, text="Calculate", command=self.print_table)
            calc_button.grid(row=last_row + line_break + 1, column=0)


if __name__ == '__main__':
    root = CTk()
    root.geometry("400x400")
    root.title('Application')
    AgrunomProjectApplication(root)
    root.mainloop()
