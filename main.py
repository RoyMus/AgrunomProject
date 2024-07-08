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
                print(f'\nLeast Significant Difference (LSD): {lsd}')
                # Dictionary to store the results
                treatment_significant_dict = {treatment: [] for treatment in df['Treatment'].unique()}

                # Calculate means for each treatment
                treatment_means = df.groupby('Treatment')['Value'].mean().to_dict()
                treatment_means = {r: treatment_means[r] for r in
                                   sorted(treatment_means, key=treatment_means.get, reverse=True)}
                RecurseFunc(treatment_means, t_critical, mse, n)

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


def get_standard_error_diff(setsStandardError, minimumSampleSize):
    return np.sqrt(2 * setsStandardError / minimumSampleSize)


# Function to get the t-statistic and check if it's a critical difference
def get_t_statistic(t_critical_value, setsStandardError, minimumSampleSize, firstTreatmentMean, secondTreatmentMean):
    sed = get_standard_error_diff(setsStandardError, minimumSampleSize)
    t_stat = abs(firstTreatmentMean - secondTreatmentMean) / sed
    is_critical = t_stat > t_critical_value
    return t_stat, is_critical


def RecurseFunc(SortedTreatementDictionary, t_critical_value, setsStandardError, minimumSampleSize):
    keys = list(SortedTreatementDictionary.keys())
    myDict = {key: set() for key in keys}
    myDict[keys[0]] = 'A'
    LetterCounter = 'A'
    Func(SortedTreatementDictionary, t_critical_value, setsStandardError, minimumSampleSize, myDict, keys,
         LetterCounter)
    print(myDict)


def Func(SortedTreatementDictionary, t_critical_value, setsStandardError, minimumSampleSize, myDict, keys,
         LetterCounter):
    for index, key in enumerate(keys[1:]):
        _, is_critical_dif = get_t_statistic(t_critical_value, setsStandardError, minimumSampleSize,
                                             SortedTreatementDictionary[key], SortedTreatementDictionary[keys[0]])

        if is_critical_dif:
            LetterCounter = IncrementLetterCounter(LetterCounter)
            myDict[key].add(LetterCounter)
            for secondKey in keys[1:]:
                _, is_critical_dif_second = get_t_statistic(t_critical_value, setsStandardError, minimumSampleSize,
                                                            SortedTreatementDictionary[key],
                                                            SortedTreatementDictionary[secondKey])
                if not is_critical_dif_second:
                    myDict[secondKey].add(LetterCounter)

            Func(SortedTreatementDictionary, t_critical_value, setsStandardError, minimumSampleSize, myDict,
                 keys[index + 1:], LetterCounter)
            return

        myDict[key].add(LetterCounter)



def RecursiveFunc(SortedTreatementDictionary, t_critical_value, setsStandardError, minimumSampleSize, dict):
    LetterCounter = 'A'
    keys = list(SortedTreatementDictionary.keys())


def IncrementLetterCounter(LetterCounter):
    AsciiOfLetter = ord(LetterCounter)
    IncrementedLettersAscii = AsciiOfLetter + 1
    return chr(IncrementedLettersAscii)


if __name__ == '__main__':
    root = CTk()
    root.geometry("400x400")
    root.title('Application')
    AgrunomProjectApplication(root)
    root.mainloop()
