import os
import tkinter
import tkinter.font
from tkinter import messagebox, filedialog
from tkinter import ttk
import statsmodels.api as sm
from scipy import stats
from scipy.stats import studentized_range
from statsmodels.formula.api import ols
from statsmodels.stats.multicomp import pairwise_tukeyhsd
import CalcUtils
import customtkinter
import Utils
import pandas as pd
from pathlib import Path
import numpy as np
from functools import lru_cache


@lru_cache(maxsize=None)
def _tukey_q_critical(k, df_error, alpha=0.05):
    return studentized_range.ppf(1 - alpha, k, df_error)


# Selectable tests: key -> dropdown label and output sheet name (Excel sheet names are max 31 chars)
TEST_MODES = {
    "t": {"label": "Each Pair Student's T", "sheet": "Each Pair Student's T"},
    "tukey": {"label": "Pairwise Tukey", "sheet": "Tukey Pairwise Test"},
    "combined": {"label": "Combined (T + Tukey)", "sheet": "Student's T & Tukey"},
}


# UI palette (light, neutral surfaces with the theme's green as the accent)
COLORS = {
    "bg": "#F3F5F7",
    "card": "#FFFFFF",
    "border": "#E1E5EA",
    "text": "#1F2933",
    "muted": "#6B7785",
    "hover": "#F2F4F6",
    "heading": "#F0F3F5",
    "stripe": "#F8FAFB",
    "selected": "#DCF2E7",
}


class AgrunomProjectApplication(customtkinter.CTkFrame):
    def __init__(self, root):
        customtkinter.set_appearance_mode("light")  # Modes: system (default), light, dark
        customtkinter.set_default_color_theme("green")  # Themes: blue (default), dark-blue, green
        customtkinter.CTkFrame.__init__(self, root)
        self.root = root
        self.filename = None
        self.input_df = None
        self.output_df = None
        self.scrollable_checkbox_frame = None
        self.block_selection = None
        self.chart_selection = None
        self.tests_popup = None
        self.x = None
        pd.options.display.float_format = '${:,.2f}'.format

        self.fonts = {
            "title": customtkinter.CTkFont(size=20, weight="bold"),
            "card_title": customtkinter.CTkFont(size=15, weight="bold"),
            "section": customtkinter.CTkFont(size=11, weight="bold"),
            "body": customtkinter.CTkFont(size=13),
            "small": customtkinter.CTkFont(size=12),
            "button": customtkinter.CTkFont(size=13, weight="bold"),
        }
        # Test selection (Student's T checked by default), shown as a dropdown of checkboxes
        self.test_vars = {mode: tkinter.BooleanVar(value=(mode == "t")) for mode in TEST_MODES}

        root.configure(fg_color=COLORS["bg"])
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)
        self._build_header()
        self._build_body()
        self.status_label = customtkinter.CTkLabel(root, text="", font=self.fonts["small"],
                                                   text_color=COLORS["muted"], anchor="w")
        self.status_label.grid(row=3, column=0, sticky="ew", padx=24, pady=(0, 10))
        self.ready_status = "Open an Excel file to get started."
        self.set_status(self.ready_status)
        self._set_controls_enabled(False)

        # Close the tests dropdown when clicking elsewhere or moving/resizing the window
        root.bind_all("<Button-1>", self._on_global_click, add="+")
        root.bind("<Configure>", lambda e: self.close_tests_popup() if e.widget is root else None, add="+")
        root.bind_all("<Escape>", lambda e: self.close_tests_popup(), add="+")

    # ---------- layout ----------
    def _build_header(self):
        header = customtkinter.CTkFrame(self.root, fg_color=COLORS["card"], corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(1, weight=1)
        self.file_label = customtkinter.CTkLabel(header, text="No file loaded", font=self.fonts["small"],
                                                 text_color=COLORS["muted"])
        self.file_label.grid(row=0, column=2, rowspan=2, padx=16)
        customtkinter.CTkButton(header, text="Open Excel File…", font=self.fonts["button"], height=36,
                                corner_radius=8, command=self.read_file
                                ).grid(row=0, column=3, rowspan=2, padx=(0, 24))
        # hairline under the header
        customtkinter.CTkFrame(self.root, height=1, corner_radius=0, fg_color=COLORS["border"]
                               ).grid(row=1, column=0, sticky="ew")

    def _card(self, parent, **kwargs):
        return customtkinter.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=12, border_width=1,
                                      border_color=COLORS["border"], **kwargs)

    def _section_label(self, parent, text, row, pady=(18, 6)):
        label = customtkinter.CTkLabel(parent, text=text.upper(), font=self.fonts["section"],
                                       text_color=COLORS["muted"], anchor="w")
        label.grid(row=row, column=0, sticky="w", padx=18, pady=pady)
        self.section_labels.append(label)
        return label

    def _build_body(self):
        self.section_labels = []
        body = self.body = customtkinter.CTkFrame(self.root, fg_color="transparent")
        body.grid(row=2, column=0, sticky="nsew", padx=20, pady=16)
        body.rowconfigure(0, weight=1)

        # --- left (right for Hebrew sheets): analysis settings ---
        sidebar = self.sidebar = self._card(body, width=300)
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0, weight=1)
        sidebar.rowconfigure(4, weight=1)
        self.sidebar_title = customtkinter.CTkLabel(sidebar, text="Analysis", font=self.fonts["card_title"],
                                                    text_color=COLORS["text"])
        self.sidebar_title.grid(row=0, column=0, sticky="w", padx=18, pady=(16, 0))

        self._section_label(sidebar, "X", row=1, pady=(14, 6))
        self.x_container = customtkinter.CTkFrame(sidebar, fg_color="transparent")
        self.x_container.grid(row=2, column=0, sticky="ew", padx=18)
        self.x_container.columnconfigure(0, weight=1)
        self.x_placeholder = customtkinter.CTkLabel(self.x_container, text="Load a file to choose X",
                                                    font=self.fonts["small"], text_color=COLORS["muted"])
        self.x_placeholder.grid(row=0, column=0, sticky="w")

        y_header = self.y_header = customtkinter.CTkFrame(sidebar, fg_color="transparent")
        y_header.grid(row=3, column=0, sticky="ew", padx=(0, 12))
        self.y_label = self._section_label(y_header, "Y", row=0)
        link_style = dict(width=0, height=24, fg_color="transparent", hover_color=COLORS["hover"],
                          font=self.fonts["small"], corner_radius=6)
        self.select_all_button = customtkinter.CTkButton(y_header, text="All", command=lambda: self._set_all_y(True),
                                                         **link_style)
        self.select_all_button.grid(row=0, column=1, pady=(18, 6))
        self.clear_button = customtkinter.CTkButton(y_header, text="None", command=lambda: self._set_all_y(False),
                                                    **link_style)
        self.clear_button.grid(row=0, column=2, pady=(18, 6))

        self.y_container = customtkinter.CTkFrame(sidebar, fg_color="transparent")
        self.y_container.grid(row=4, column=0, sticky="nsew", padx=12)
        self.y_container.columnconfigure(0, weight=1)
        self.y_container.rowconfigure(0, weight=1)
        self.y_placeholder = customtkinter.CTkLabel(self.y_container, text="Load a file to choose variables",
                                                    font=self.fonts["small"], text_color=COLORS["muted"])
        self.y_placeholder.grid(row=0, column=0)

        self._section_label(sidebar, "Block", row=5)
        self.block_container = customtkinter.CTkFrame(sidebar, fg_color="transparent")
        self.block_container.grid(row=6, column=0, sticky="ew", padx=18)
        self.block_container.columnconfigure(0, weight=1)

        self._section_label(sidebar, "Graphs", row=7)
        self.chart_container = customtkinter.CTkFrame(sidebar, fg_color="transparent")
        self.chart_container.grid(row=8, column=0, sticky="ew", padx=18)
        self.chart_container.columnconfigure(0, weight=1)

        self._section_label(sidebar, "Calculate options", row=9)
        self.tests_button = customtkinter.CTkButton(
            sidebar, command=self.toggle_tests_popup, height=36, corner_radius=8, anchor="w",
            fg_color=COLORS["card"], hover_color=COLORS["hover"], text_color=COLORS["text"],
            text_color_disabled=COLORS["muted"], border_width=1, border_color=COLORS["border"],
            font=self.fonts["body"])
        self.tests_button.grid(row=10, column=0, sticky="ew", padx=18)
        self.update_tests_button()

        self.calc_button = customtkinter.CTkButton(sidebar, text="Calculate", height=42, corner_radius=8,
                                                   font=customtkinter.CTkFont(size=14, weight="bold"),
                                                   command=self.on_calculate)
        self.calc_button.grid(row=11, column=0, sticky="ew", padx=18, pady=(20, 18))

        # --- right (left for Hebrew sheets): data preview ---
        preview = self.preview = self._card(body)
        preview.columnconfigure(0, weight=1)
        preview.rowconfigure(2, weight=1)
        self.preview_title = customtkinter.CTkLabel(preview, text="Data preview", font=self.fonts["card_title"],
                                                    text_color=COLORS["text"])
        self.preview_title.grid(row=0, column=0, sticky="w", padx=18, pady=(16, 0))
        self.preview_subtitle = customtkinter.CTkLabel(preview, text="No data loaded", font=self.fonts["small"],
                                                       text_color=COLORS["muted"])
        self.preview_subtitle.grid(row=1, column=0, sticky="w", padx=18, pady=(0, 10))

        self.table = customtkinter.CTkFrame(preview, fg_color=COLORS["card"], corner_radius=0)
        self.table.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 10))
        self.table.rowconfigure(0, weight=1)
        self._style_treeview()
        self.tv1 = ttk.Treeview(self.table, style="Data.Treeview", show="headings")
        self.tv1.tag_configure("odd", background=COLORS["stripe"])
        self.vsb = customtkinter.CTkScrollbar(self.table, orientation="vertical", command=self.tv1.yview)
        self.hsb = customtkinter.CTkScrollbar(self.table, orientation="horizontal", command=self.tv1.xview)
        self.tv1.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        self.empty_label = customtkinter.CTkLabel(self.table, text="Open an Excel file to preview its data here",
                                                  font=self.fonts["body"], text_color=COLORS["muted"],
                                                  fg_color=COLORS["card"])
        self._set_table_direction(rtl=False)
        self._apply_direction(rtl=False)

    def _apply_direction(self, rtl):
        """Mirror the main screen for Hebrew sheets: settings panel on the right, headings and
        options aligned right, indicators to the right of their labels."""
        side = "e" if rtl else "w"
        sidebar_col, preview_col = (1, 0) if rtl else (0, 1)
        self.body.columnconfigure(sidebar_col, weight=0)
        self.body.columnconfigure(preview_col, weight=1)
        self.sidebar.grid(row=0, column=sidebar_col, sticky="ns", padx=(16, 0) if rtl else (0, 16))
        self.preview.grid(row=0, column=preview_col, sticky="nsew")
        for label in self.section_labels + [self.sidebar_title, self.preview_title, self.preview_subtitle]:
            label.grid_configure(sticky=side)
            label.configure(anchor=side)
        # Y header: the "Y" heading on the outer side, All/None buttons on the inner side
        label_col, all_col, none_col = (2, 1, 0) if rtl else (0, 1, 2)
        for col in range(3):
            self.y_header.columnconfigure(col, weight=1 if col == label_col else 0)
        self.y_header.grid_configure(padx=(12, 0) if rtl else (0, 12))
        self.y_label.grid_configure(column=label_col)
        self.select_all_button.grid_configure(column=all_col)
        self.clear_button.grid_configure(column=none_col)
        self.x_placeholder.grid_configure(sticky=side)

    def _set_table_direction(self, rtl):
        """Lay the preview out like Excel: in a right-to-left sheet the first column is on the right
        and the vertical scrollbar is on the left."""
        table_col, bar_col = (1, 0) if rtl else (0, 1)
        self.table.columnconfigure(table_col, weight=1)
        self.table.columnconfigure(bar_col, weight=0)
        self.tv1.grid(row=0, column=table_col, sticky="nsew")
        self.vsb.grid(row=0, column=bar_col, sticky="ns")
        self.hsb.grid(row=1, column=table_col, sticky="ew")
        if self.input_df is None:
            self.empty_label.grid(row=0, column=table_col)

    def _style_treeview(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        family = tkinter.font.nametofont("TkDefaultFont").actual()["family"]
        self._tree_font = tkinter.font.Font(family=family, size=12)
        self._tree_heading_font = tkinter.font.Font(family=family, size=12, weight="bold")
        style.configure("Data.Treeview", background=COLORS["card"], fieldbackground=COLORS["card"],
                        foreground=COLORS["text"], rowheight=28, borderwidth=0, font=self._tree_font)
        style.layout("Data.Treeview", [("Data.Treeview.treearea", {"sticky": "nswe"})])
        style.map("Data.Treeview", background=[("selected", COLORS["selected"])],
                  foreground=[("selected", COLORS["text"])])
        style.configure("Data.Treeview.Heading", background=COLORS["heading"], foreground=COLORS["text"],
                        relief="flat", borderwidth=0, font=self._tree_heading_font, padding=(8, 6))
        style.map("Data.Treeview.Heading", background=[("active", COLORS["border"])])

    # ---------- small UI helpers ----------
    def set_status(self, text):
        self.status_label.configure(text=text)

    def _set_controls_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        for widget in (self.tests_button, self.calc_button, self.select_all_button, self.clear_button):
            widget.configure(state=state)

    def _set_all_y(self, checked):
        if self.scrollable_checkbox_frame is not None:
            self.scrollable_checkbox_frame.set_all(checked)

    def toggle_tests_popup(self):
        if self.tests_popup is not None:
            self.close_tests_popup()
            return
        # A white dropdown panel built from CTk widgets (a native tk.Menu ignores colors on macOS).
        # It's an overlay inside the main window rather than a separate borderless window: those
        # can't take focus on macOS and get stuck on screen.
        popup = tkinter.Frame(self.root, bg=COLORS["border"])
        inner = customtkinter.CTkFrame(popup, fg_color=COLORS["card"], corner_radius=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        modes = list(TEST_MODES.items())
        for i, (mode, info) in enumerate(modes):
            customtkinter.CTkCheckBox(inner, text=info["label"], variable=self.test_vars[mode], onvalue=True,
                                      offvalue=False, command=self.update_tests_button, font=self.fonts["body"],
                                      text_color=COLORS["text"], checkbox_width=20, checkbox_height=20,
                                      border_width=2, corner_radius=5
                                      ).pack(anchor="w", padx=14,
                                             pady=(12 if i == 0 else 5, 12 if i == len(modes) - 1 else 5))
        button = self.tests_button
        # Plain tkinter place() (pixel coordinates, no CTk scaling) relative to the main window
        popup.place(x=button.winfo_rootx() - self.root.winfo_rootx(),
                    y=button.winfo_rooty() - self.root.winfo_rooty() + button.winfo_height() + 4,
                    width=button.winfo_width())
        popup.lift()
        self.tests_popup = popup

    def close_tests_popup(self):
        if self.tests_popup is not None:
            self.tests_popup.destroy()
            self.tests_popup = None

    def _on_global_click(self, event):
        if self.tests_popup is None or not isinstance(event.widget, tkinter.Misc):
            return
        # Clicks inside the panel (checkboxes) or on the toggle button itself are handled there
        if self._is_inside(event.widget, self.tests_popup) or self._is_inside(event.widget, self.tests_button):
            return
        self.close_tests_popup()

    @staticmethod
    def _is_inside(widget, container):
        while widget is not None:
            if widget is container:
                return True
            widget = widget.master
        return False

    def update_tests_button(self):
        chosen = [TEST_MODES[mode]["label"] for mode in TEST_MODES if self.test_vars[mode].get()]
        text = ", ".join(chosen) if chosen else "Choose options"
        if len(text) > 30:
            text = f"{len(chosen)} options selected"
        self.tests_button.configure(text=f"{text}  ▾")

    def on_calculate(self):
        self.close_tests_popup()
        self.calc_button.configure(state="disabled", text="Calculating…")
        self.set_status("Calculating… this can take a moment for large files.")
        self.root.update_idletasks()
        output_path = None
        try:
            output_path = self.calculate()
        finally:
            self.calc_button.configure(state="normal", text="Calculate")
            if output_path is None:
                self.set_status(self.ready_status)

    @staticmethod
    def _tukey_letters(df, model, block, treatment_means, treatment_means_sorted):
        if block != "אקראיות גמורה":
            mse_error = model.mse_resid
            df_error = model.df_resid

            treatment_stats = df.groupby('Treatment')['Value'].agg(['mean', 'count'])
            treatments = treatment_stats.index.tolist()
            means = treatment_stats['mean'].values
            n_per_group = treatment_stats['count'].iloc[0]
            k = len(treatments)

            # Same for every pair, and studentized_range.ppf is slow (numerical integration),
            # so compute it once per (k, df) and reuse it across pairs, columns and labels.
            q_critical = _tukey_q_critical(k, float(df_error))
            # q statistic denominator: se_diff / sqrt(2) == sqrt(MSE / n)
            q_denominator = np.sqrt(mse_error / n_per_group)

            tukey_results = []
            for i in range(k):
                for j in range(i + 1, k):
                    meandiff = means[j] - means[i]
                    q_stat = abs(meandiff) / q_denominator
                    # p_adj < alpha  <=>  q_stat > q_critical (the CDF is monotonic),
                    # so there's no need to compute a p-value per pair.
                    reject = q_stat > q_critical
                    tukey_results.append([treatments[i], treatments[j], reject])
            tukey_df = pd.DataFrame(tukey_results, columns=['group1', 'group2', 'reject'])
        else:
            # Perform Tukey HSD test
            tukey_hsd = pairwise_tukeyhsd(endog=df['Value'],
                                          groups=df['Treatment'],
                                          alpha=0.05)
            tukey_df = pd.DataFrame(data=tukey_hsd._results_table.data[1:],
                                    columns=tukey_hsd._results_table.data[0])
        tukey_dictionary = dict()
        for treatment in treatment_means:
            tukey_dictionary[treatment] = dict()
        for index, row in tukey_df.iterrows():
            tukey_dictionary[row["group1"]][row["group2"]] = row["reject"] == True
            tukey_dictionary[row["group2"]][row["group1"]] = row["reject"] == True

        return CalcUtils.calculate_significant_letters_tukey(tukey_dictionary, list(treatment_means_sorted.keys()))

    @staticmethod
    def _student_t_letters(df, anova_table, treatment_means_sorted):
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
        return CalcUtils.calculate_significant_letters(treatment_means_sorted, t_critical, mse, n)

    def calculate(self):
        modes = [mode for mode in TEST_MODES if self.test_vars[mode].get()]
        x_value = self.x.get_checked_item()
        block = self.block_selection.get_checked_item()
        items = self.scrollable_checkbox_frame.get_checked_items()
        flip_xy = self.chart_selection.get_checked_item()
        flip_xy = flip_xy != "ימים במקרא"
        if len(modes) == 0:
            messagebox.showerror("No calculate option chosen", "Please choose at least one calculate option")
            return
        if len(items) == 0:
            messagebox.showerror("No Y chosen", "Please choose at least one Y value")
            return
        if len(self.input_df[x_value]) > 0 and self.input_df[x_value][0] != "UTC":
            result = messagebox.askokcancel("UTC is not in place","The UTC Column Should Be First, Would You Like To Procced?")
            if not result:
                return
        output_path, copy_note = Utils.generate_new_file(self.filename)
        if copy_note:
            messagebox.showwarning("Workbook adjusted", copy_note)
        # Each selected test gets its own sheet, with its name as the title on top
        for mode in modes:
            text = TEST_MODES[mode]["sheet"]
            Utils.write_text_to_excel(output_path, text, text)
        need_t = any(mode in ("t", "combined") for mode in modes)
        need_tukey = any(mode in ("tukey", "combined") for mode in modes)
        for label in items:
            means_dict = {}
            output_dicts = {mode: {} for mode in modes}
            # Combined table: both letter columns of a day share the header "<day> sig"; a row above
            # the headers says which is Student's and which is Tukey. Columns are keyed uniquely here
            # and renamed to their shown header when written.
            combined_headers = {}
            combined_groups = {}
            char = ''
            index_of_row = 0
            for col in self.input_df.columns:
                if label != col.split('.')[0]:
                    continue
                df = pd.DataFrame(
                    {'Treatment': self.input_df[x_value], 'Value': pd.to_numeric(self.input_df[col], errors='coerce')})
                df['Treatment'] = df['Treatment'].astype(str)
                if block != "אקראיות גמורה":
                    df["Block"] = self.input_df["חזרה מספר"].astype('category')
                    model = ols('Value ~ C(Treatment) + C(Block)', data=df).fit()
                else:
                    model = ols('Value ~ C(Treatment)', data=df).fit()

                # Perform ANOVA
                anova_table = sm.stats.anova_lm(model, typ=2)
                # Calculate means for each treatment
                treatment_means = df.groupby('Treatment', sort=False)['Value'].mean().to_dict()
                treatment_means_sorted = {r: treatment_means[r] for r in
                                          sorted(treatment_means, key=treatment_means.get, reverse=True)}

                if index_of_row >= len(self.columns):
                    index_of_row = 0
                    char += str(1)
                day = self.columns[index_of_row]
                means = list(treatment_means.values())
                means_dict[day + char] = means

                def letters(sig_by_key):
                    return ["".join(sorted(sig_by_key[x])) for x in treatment_means.keys()]

                # Each test runs once per column, even if several selected modes use it
                t_letters = None
                tukey_letters = None
                if need_t:
                    t_letters = letters(self._student_t_letters(df, anova_table, treatment_means_sorted))
                if need_tukey:
                    tukey_letters = letters(self._tukey_letters(df, model, block, treatment_means,
                                                                treatment_means_sorted))

                for mode, out in output_dicts.items():
                    out[day + char] = means
                    if mode == "combined":
                        sig_header = self.columns[index_of_row + 1] + char
                        if index_of_row == 0:
                            # DOT is the pre-treatment baseline: a single letters column
                            out[sig_header] = t_letters
                        else:
                            out[sig_header] = t_letters
                            combined_groups[sig_header] = "Student's"
                            tukey_key = sig_header + " (Tukey)"
                            out[tukey_key] = tukey_letters
                            combined_headers[tukey_key] = sig_header
                            combined_groups[tukey_key] = "Tukey"
                    else:
                        out[self.columns[index_of_row + 1] + char] = t_letters if mode == "t" else tukey_letters
                index_of_row += 2
            for mode, out in output_dicts.items():
                mode_df = pd.DataFrame(out)
                mode_df.insert(0, label, treatment_means.keys())
                sheet = TEST_MODES[mode]["sheet"]
                if mode == "combined":
                    groups = [combined_groups.get(key) for key in mode_df.columns]
                    mode_df.columns = [combined_headers.get(key, key) for key in mode_df.columns]
                    header_row, _ = Utils.append_df_to_excel(output_path, mode_df, sheet, start_distance=3)
                    Utils.write_column_group_labels(output_path, sheet, header_row - 1, groups)
                else:
                    Utils.append_df_to_excel(output_path, mode_df, sheet)
            # Graphs only use the means, so they're generated once no matter how many tests were chosen
            self.output_df = pd.DataFrame(means_dict)
            self.output_df.insert(0, label, treatment_means.keys())
            row, df_len = Utils.append_df_to_excel(output_path, self.output_df, "1גרפים", start_distance=4,
                                                  fit_first_column=True)

            if "שכיחות" in label:
                Utils.append_chart_to_excel_openpy("%", f"Frequncy of {label} per leaf in average", output_path, row,
                                                   df_len, 1, "1גרפים", cropped=False, flip_x_y=flip_xy)
                continue
            Utils.append_chart_to_excel_openpy(label, f"Number of {label} per leaf in average", output_path, row,
                                               df_len, 1, "1גרפים", cropped=False, flip_x_y=flip_xy)
            output_dict = Utils.generate_excel_formula(self.output_df,row)
            formulated_df = pd.DataFrame(output_dict)
            formulated_df.iloc[:,0] = self.output_df.iloc[1:,0]
            formulated_df = formulated_df.rename(columns={"DOT":self.output_df.columns[0]})
            row, df_len = Utils.append_df_to_excel(output_path, formulated_df, "1גרפים", start_distance=4)
            Utils.append_chart_to_excel_openpy("%", f"Decrease(%) of {label} as correlation to control", output_path,
                                               row + 1, df_len, 1, "1גרפים", cropped=True, flip_x_y=flip_xy)
        self.set_status(f"Done. Results saved to {Path(output_path).name}")
        result = messagebox.askokcancel("Calculation finished", "Would you like to open the file directory?")
        if result:
            os.system(f'explorer /select,"{output_path}"')
            os.startfile(output_path)
        return output_path

    def init_ui(self, sheet_name):
        self.newWindow.destroy()
        with open(self.filename, 'rb') as f:
            self.input_df = pd.read_excel(f, sheet_name, index_col=False)
            self.input_df.dropna(how='all', axis=1, inplace=True)
            # Identify the index of the first row
            first_date_row_index = Utils.find_first_date_row(self.input_df)
            if first_date_row_index != -1:
                # Keep only the rows before the first row with NaN
                columns = self.input_df.iloc[first_date_row_index].values
                try:
                    clean_dates = [pd.to_datetime(x) for i, x in enumerate(columns) if
                                   pd.to_datetime(x) is not pd.NaT and x not in columns[:i]]
                except:
                    messagebox.showerror("Dates Error", "Please Check The Dates")
                    return
                if len(clean_dates) > 0:
                    self.columns = ["DOT", "DOT sig"]
                    last_date = clean_dates[0]
                    for date in clean_dates[1:]:
                        self.columns.append(str((date - last_date).days) + "DAT")
                        self.columns.append(str((date - last_date).days) + "DAT sig")
                else:
                    self.columns = ["DOT", "DOT sig", "3DAT", "3DAT sig", "7DAT", "7DAT sig", "10DAT", "10DAT sig",
                                    "14DAT",
                                    "14DAT sig"]

                self.input_df = self.input_df.iloc[:first_date_row_index]
            else:
                messagebox.showerror("Dates Error","Please Check The Dates")
                return
            columns = list(self.input_df.columns)
            # Data preview: Hebrew text fixed for display, and mirrored like Excel when the sheet is RTL
            rtl = Utils.sheet_is_rtl(self.filename, sheet_name, columns)
            shown_columns = columns[::-1] if rtl else columns
            self.tv1.delete(*self.tv1.get_children())
            self.tv1["columns"] = shown_columns
            for column in shown_columns:
                self.tv1.heading(column, text=Utils.to_display(str(column)))
                self.tv1.column(column, width=110, minwidth=80, anchor="center", stretch=False)
            for i, row in enumerate(self.input_df.to_numpy().tolist()):
                values = ["" if pd.isna(v) else Utils.to_display(v) for v in row]
                if rtl:
                    values = values[::-1]
                self.tv1.insert("", "end", values=values, tags=("odd",) if i % 2 else ())
            self.empty_label.grid_remove()
            self._set_table_direction(rtl)
            self._apply_direction(rtl)
            # In RTL, start scrolled to the right so the first columns are visible
            self.tv1.update_idletasks()
            self.tv1.xview_moveto(1.0 if rtl else 0.0)
            unduplicatedcolumns = []
            for col in columns:
                if col.split('.')[0] not in unduplicatedcolumns:
                    unduplicatedcolumns.append(col)
            x_values = unduplicatedcolumns
            y_values = unduplicatedcolumns
            blocks = ["אקראיות גמורה"]
            if "טיפול" in unduplicatedcolumns:
                x_values = ["טיפול"]
                y_values.remove("טיפול")
            else:
                messagebox.showerror("Invalid Columns","No Column Named טיפול")
                return
            if "חזרה מספר" in unduplicatedcolumns:
                blocks.append("בלוקים באקראי")
                y_values.remove("חזרה מספר")
            else:
                messagebox.showerror("Invalid Columns","No Column Named חזרה מספר")
                return
            # (Re)build the selectors, replacing any from a previously loaded file
            for container in (self.x_container, self.y_container, self.block_container, self.chart_container):
                for child in container.winfo_children():
                    if child not in (self.x_placeholder, self.y_placeholder):
                        child.destroy()
            self.x_placeholder.grid_remove()
            self.y_placeholder.grid_remove()
            side = "e" if rtl else "w"
            self.x = Utils.RadiobuttonGroup(self.x_container, item_list=x_values, rtl=rtl)
            self.x.grid(row=0, column=0, sticky=side)
            self.scrollable_checkbox_frame = Utils.ScrollableCheckBoxFrame(master=self.y_container,
                                                                           item_list=y_values,
                                                                           fg_color="transparent", rtl=rtl)
            self.scrollable_checkbox_frame.grid(row=0, column=0, sticky="nsew")
            self.block_selection = Utils.RadiobuttonGroup(self.block_container, item_list=blocks, rtl=rtl,
                                                          default="בלוקים באקראי")
            self.block_selection.grid(row=0, column=0, sticky=side)
            self.chart_selection = Utils.RadiobuttonGroup(self.chart_container,
                                                          item_list=["ימים במקרא", "ימים בציר X"], rtl=rtl,
                                                          default="ימים בציר X")
            self.chart_selection.grid(row=0, column=0, sticky=side)
            self._set_controls_enabled(True)

            path = Path(self.filename)
            file_name = Utils.to_display(path.stem) + path.suffix
            self.file_label.configure(text=file_name, text_color=COLORS["text"])
            self.preview_subtitle.configure(
                text=f"{file_name}  ·  sheet {Utils.to_display(sheet_name)}  ·  {len(self.input_df)} rows")
            self.ready_status = f"{len(y_values)} variables found. Choose variables and tests, then click Calculate."
            self.set_status(self.ready_status)

    def read_file(self):
        filename = filedialog.askopenfilename(initialdir="/", title="Select file",
                                              filetypes=[("Excel files", ".xlsx .xls")])
        if not filename:
            return
        self.filename = filename
        xls = pd.ExcelFile(self.filename)

        # Sheet chooser
        self.newWindow = customtkinter.CTkToplevel(self.root, fg_color=COLORS["bg"])
        self.newWindow.title("Choose sheet")
        self.newWindow.geometry("380x420+%d+%d" % (self.root.winfo_x() + 80, self.root.winfo_y() + 80))
        self.newWindow.resizable(False, False)
        self.newWindow.attributes('-topmost', 'true')
        self.newWindow.columnconfigure(0, weight=1)
        self.newWindow.rowconfigure(1, weight=1)
        customtkinter.CTkLabel(self.newWindow, text="Choose a sheet", font=self.fonts["card_title"],
                               text_color=COLORS["text"]).grid(row=0, column=0, sticky="w", padx=20, pady=(18, 2))
        card = self._card(self.newWindow)
        card.grid(row=1, column=0, sticky="nsew", padx=20, pady=10)
        card.columnconfigure(0, weight=1)
        card.rowconfigure(0, weight=1)
        rtl = any(Utils.has_hebrew(name) for name in xls.sheet_names)
        sheets = Utils.ScrollableRadiobuttonFrame(card, item_list=xls.sheet_names, fg_color="transparent", rtl=rtl)
        sheets.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        buttons = customtkinter.CTkFrame(self.newWindow, fg_color="transparent")
        buttons.grid(row=2, column=0, sticky="e", padx=20, pady=(0, 18))
        customtkinter.CTkButton(buttons, text="Cancel", width=90, height=34, corner_radius=8,
                                fg_color=COLORS["card"], hover_color=COLORS["hover"], text_color=COLORS["text"],
                                border_width=1, border_color=COLORS["border"],
                                command=self.newWindow.destroy).grid(row=0, column=0, padx=(0, 8))
        customtkinter.CTkButton(buttons, text="Open", width=90, height=34, corner_radius=8, font=self.fonts["button"],
                                command=lambda: self.init_ui(sheets.get_checked_item())).grid(row=0, column=1)
        self.newWindow.focus()

    def clear_data(self):
        self.tv1.delete(*self.tv1.get_children())


if __name__ == '__main__':
    root = customtkinter.CTk()
    root.geometry("1100x760")
    root.minsize(900, 620)
    root.title('AgrunomProject')
    AgrunomProjectApplication(root)
    root.focus()
    root.mainloop()
