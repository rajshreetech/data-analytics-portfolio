import os
import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import matplotlib.pyplot as plt

from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class DataAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Analysis & Report Builder")
        self.root.geometry("1200x750")
        self.root.minsize(1000, 650)

        # -----------------------------
        # Application Variables
        # -----------------------------
        self.file_path = None
        self.df = None
        self.report_df = None
        self.current_figure = None
        self.chart_canvas = None

        self.text_columns = []
        self.numeric_columns = []

        self.group_by_var = tk.StringVar()
        self.aggregation_var = tk.StringVar()
        self.value_column_var = tk.StringVar()
        self.chart_type_var = tk.StringVar()
        self.export_format_var = tk.StringVar(value="Excel (.xlsx)")

        # -----------------------------
        # Create GUI
        # -----------------------------
        self.create_widgets()

    # ============================================================
    # GUI DESIGN
    # ============================================================

    def create_widgets(self):

        # Main title
        title_label = ttk.Label(
            self.root,
            text="Data Analysis & Report Builder",
            font=("Arial", 18, "bold")
        )
        title_label.pack(pady=10)

        # ========================================================
        # FILE SECTION
        # ========================================================

        file_frame = ttk.LabelFrame(
            self.root,
            text="1. File Selection",
            padding=10
        )
        file_frame.pack(fill="x", padx=15, pady=5)

        self.file_label = ttk.Label(
            file_frame,
            text="No file selected",
            width=80
        )
        self.file_label.grid(row=0, column=0, padx=5, pady=5)

        browse_button = ttk.Button(
            file_frame,
            text="Browse File",
            command=self.browse_file
        )
        browse_button.grid(row=0, column=1, padx=5, pady=5)

        read_button = ttk.Button(
            file_frame,
            text="Read Data",
            command=self.read_data
        )
        read_button.grid(row=0, column=2, padx=5, pady=5)

        # ========================================================
        # DATA INFORMATION SECTION
        # ========================================================

        info_frame = ttk.LabelFrame(
            self.root,
            text="2. Dataset Information",
            padding=10
        )
        info_frame.pack(fill="x", padx=15, pady=5)

        self.rows_label = ttk.Label(
            info_frame,
            text="Total Rows: -"
        )
        self.rows_label.grid(row=0, column=0, padx=20)

        self.columns_count_label = ttk.Label(
            info_frame,
            text="Total Columns: -"
        )
        self.columns_count_label.grid(row=0, column=1, padx=20)

        self.headings_label = ttk.Label(
            info_frame,
            text="Column Headings: -",
            wraplength=900
        )
        self.headings_label.grid(
            row=1,
            column=0,
            columnspan=2,
            pady=10
        )

        # ========================================================
        # REPORT BUILDER SECTION
        # ========================================================

        report_frame = ttk.LabelFrame(
            self.root,
            text="3. Report Builder",
            padding=10
        )
        report_frame.pack(fill="x", padx=15, pady=5)

        # Group By
        ttk.Label(
            report_frame,
            text="Group By Column:"
        ).grid(row=0, column=0, padx=10, pady=5)

        self.group_by_combo = ttk.Combobox(
            report_frame,
            textvariable=self.group_by_var,
            state="readonly",
            width=20
        )
        self.group_by_combo.grid(
            row=0,
            column=1,
            padx=10,
            pady=5
        )

        # Aggregation
        ttk.Label(
            report_frame,
            text="Aggregation:"
        ).grid(row=0, column=2, padx=10, pady=5)

        self.aggregation_combo = ttk.Combobox(
            report_frame,
            textvariable=self.aggregation_var,
            state="readonly",
            width=15,
            values=[
                "sum",
                "mean",
                "average",
                "max",
                "min",
                "count",
                "median"
            ]
        )
        self.aggregation_combo.grid(
            row=0,
            column=3,
            padx=10,
            pady=5
        )

        # Value Column
        ttk.Label(
            report_frame,
            text="Value Column:"
        ).grid(row=0, column=4, padx=10, pady=5)

        self.value_column_combo = ttk.Combobox(
            report_frame,
            textvariable=self.value_column_var,
            state="readonly",
            width=20
        )
        self.value_column_combo.grid(
            row=0,
            column=5,
            padx=10,
            pady=5
        )

        preview_report_button = ttk.Button(
            report_frame,
            text="Preview Report",
            command=self.preview_report
        )
        preview_report_button.grid(
            row=0,
            column=6,
            padx=15,
            pady=5
        )

        # ========================================================
        # REPORT TABLE SECTION
        # ========================================================

        table_frame = ttk.LabelFrame(
            self.root,
            text="4. Report Preview",
            padding=5
        )
        table_frame.pack(
            fill="both",
            expand=True,
            padx=15,
            pady=5
        )

        self.tree = ttk.Treeview(
            table_frame,
            show="headings"
        )

        vertical_scroll = ttk.Scrollbar(
            table_frame,
            orient="vertical",
            command=self.tree.yview
        )

        horizontal_scroll = ttk.Scrollbar(
            table_frame,
            orient="horizontal",
            command=self.tree.xview
        )

        self.tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set
        )

        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew"
        )

        vertical_scroll.grid(
            row=0,
            column=1,
            sticky="ns"
        )

        horizontal_scroll.grid(
            row=1,
            column=0,
            sticky="ew"
        )

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # ========================================================
        # BOTTOM CONTROL SECTION
        # ========================================================

        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill="x", padx=15, pady=10)

        # Export Report
        export_report_frame = ttk.LabelFrame(
            bottom_frame,
            text="Export Report",
            padding=8
        )
        export_report_frame.pack(
            side="left",
            padx=5
        )

        self.export_format_combo = ttk.Combobox(
            export_report_frame,
            textvariable=self.export_format_var,
            state="readonly",
            width=15,
            values=[
                "Excel (.xlsx)",
                "CSV (.csv)"
            ]
        )
        self.export_format_combo.grid(
            row=0,
            column=0,
            padx=5
        )

        export_button = ttk.Button(
            export_report_frame,
            text="Export Report",
            command=self.export_report
        )
        export_button.grid(
            row=0,
            column=1,
            padx=5
        )

        # Chart Builder
        chart_frame = ttk.LabelFrame(
            bottom_frame,
            text="Chart Builder",
            padding=8
        )
        chart_frame.pack(
            side="left",
            padx=20
        )

        ttk.Label(
            chart_frame,
            text="Chart Type:"
        ).grid(row=0, column=0, padx=5)

        self.chart_type_combo = ttk.Combobox(
            chart_frame,
            textvariable=self.chart_type_var,
            state="readonly",
            width=15,
            values=[
                "Bar Chart",
                "Column Chart",
                "Line Chart",
                "Pie Chart"
            ]
        )
        self.chart_type_combo.grid(
            row=0,
            column=1,
            padx=5
        )

        preview_chart_button = ttk.Button(
            chart_frame,
            text="Preview Chart",
            command=self.preview_chart
        )
        preview_chart_button.grid(
            row=0,
            column=2,
            padx=5
        )

        export_chart_button = ttk.Button(
            chart_frame,
            text="Export Chart",
            command=self.export_chart
        )
        export_chart_button.grid(
            row=0,
            column=3,
            padx=5
        )

        # ========================================================
        # STATUS BAR
        # ========================================================

        self.status_label = ttk.Label(
            self.root,
            text="Ready",
            relief="sunken",
            anchor="w"
        )
        self.status_label.pack(
            fill="x",
            side="bottom"
        )

    # ============================================================
    # FILE SELECTION
    # ============================================================

    def browse_file(self):

        file_path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                ("Data Files", "*.csv *.xlsx *.xls"),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls")
            ]
        )

        if file_path:
            self.file_path = file_path
            self.file_label.config(
                text=os.path.basename(file_path)
            )

            # Reset previous data
            self.df = None
            self.report_df = None

            self.clear_treeview()
            self.clear_chart()

            self.rows_label.config(
                text="Total Rows: -"
            )

            self.columns_count_label.config(
                text="Total Columns: -"
            )

            self.headings_label.config(
                text="Column Headings: -"
            )

            self.group_by_combo["values"] = []
            self.value_column_combo["values"] = []

            self.group_by_var.set("")
            self.value_column_var.set("")
            self.aggregation_var.set("")
            self.chart_type_var.set("")

            self.status_label.config(
                text="File selected. Click 'Read Data'."
            )

    # ============================================================
    # READ DATA
    # ============================================================

    def read_data(self):

        if not self.file_path:
            messagebox.showerror(
                "Error",
                "Please select a CSV or Excel file first."
            )
            return

        try:

            extension = os.path.splitext(
                self.file_path
            )[1].lower()

            # Read CSV
            if extension == ".csv":
                self.df = pd.read_csv(
                    self.file_path
                )

            # Read Excel
            elif extension in [".xlsx", ".xls"]:
                self.df = pd.read_excel(
                    self.file_path
                )

            else:
                messagebox.showerror(
                    "Error",
                    "Unsupported file format."
                )
                return

            # Dataset information
            total_rows = self.df.shape[0]
            total_columns = self.df.shape[1]

            self.rows_label.config(
                text=f"Total Rows: {total_rows}"
            )

            self.columns_count_label.config(
                text=f"Total Columns: {total_columns}"
            )

            self.headings_label.config(
                text="Column Headings: "
                     + ", ".join(
                        map(str, self.df.columns)
                     )
            )

            # Detect column types
            self.detect_columns()

            self.status_label.config(
                text="Data successfully loaded."
            )

            messagebox.showinfo(
                "Success",
                "Dataset loaded successfully."
            )

        except Exception as error:

            messagebox.showerror(
                "Error",
                f"Unable to read the file.\n\n{error}"
            )

    # ============================================================
    # DYNAMIC COLUMN DETECTION
    # ============================================================

    def detect_columns(self):

        self.text_columns = []
        self.numeric_columns = []

        for column in self.df.columns:

            series = self.df[column]

            # Already numeric
            if pd.api.types.is_numeric_dtype(series):

                self.numeric_columns.append(column)

            # Object/String columns
            elif (
                pd.api.types.is_object_dtype(series)
                or pd.api.types.is_string_dtype(series)
            ):

                # Try converting text to numeric
                converted = pd.to_numeric(
                    series,
                    errors="coerce"
                )

                # Calculate percentage of numeric values
                non_null_count = series.notna().sum()

                if non_null_count > 0:

                    numeric_count = converted.notna().sum()

                    numeric_percentage = (
                        numeric_count / non_null_count
                    )

                    # If 80% values are numeric,
                    # treat as numeric column
                    if numeric_percentage >= 0.80:

                        self.df[column] = converted
                        self.numeric_columns.append(column)

                    else:
                        self.text_columns.append(column)

                else:
                    self.text_columns.append(column)

            else:
                self.text_columns.append(column)

        # Update dropdowns
        self.group_by_combo["values"] = self.text_columns
        self.value_column_combo["values"] = self.numeric_columns

        if self.text_columns:
            self.group_by_var.set(
                self.text_columns[0]
            )

        if self.numeric_columns:
            self.value_column_var.set(
                self.numeric_columns[0]
            )

    # ============================================================
    # CREATE REPORT
    # ============================================================

    def preview_report(self):

        # Check file/data
        if self.df is None:

            messagebox.showerror(
                "Error",
                "Please select and read a file first."
            )
            return

        group_column = self.group_by_var.get()
        aggregation = self.aggregation_var.get()
        value_column = self.value_column_var.get()

        # Validation
        if not group_column:

            messagebox.showerror(
                "Error",
                "Please select a Group By column."
            )
            return

        if not aggregation:

            messagebox.showerror(
                "Error",
                "Please select an aggregation method."
            )
            return

        if not value_column:

            messagebox.showerror(
                "Error",
                "Please select a Value column."
            )
            return

        try:

            # Convert average to mean
            if aggregation == "average":
                aggregation_function = "mean"
            else:
                aggregation_function = aggregation

            # COUNT aggregation
            if aggregation_function == "count":

                report = (
                    self.df
                    .groupby(
                        group_column,
                        dropna=False
                    )[value_column]
                    .count()
                    .reset_index()
                )

            else:

                report = (
                    self.df
                    .groupby(
                        group_column,
                        dropna=False
                    )[value_column]
                    .agg(
                        aggregation_function
                    )
                    .reset_index()
                )

            # Rename aggregation column
            report.columns = [
                group_column,
                f"{aggregation.title()} of {value_column}"
            ]

            # Sort descending
            report = report.sort_values(
                by=report.columns[1],
                ascending=False
            )

            # Reset index
            report = report.reset_index(
                drop=True
            )

            self.report_df = report

            # Clear previous chart
            self.clear_chart()

            # Display table
            self.display_report()

            self.status_label.config(
                text="Report generated successfully."
            )

        except Exception as error:

            messagebox.showerror(
                "Report Error",
                f"Unable to generate report.\n\n{error}"
            )

    # ============================================================
    # DISPLAY REPORT IN TREEVIEW
    # ============================================================

    def display_report(self):

        self.clear_treeview()

        # Configure columns
        self.tree["columns"] = list(
            self.report_df.columns
        )

        for column in self.report_df.columns:

            self.tree.heading(
                column,
                text=column
            )

            self.tree.column(
                column,
                width=250,
                anchor="center"
            )

        # Insert data
        for _, row in self.report_df.iterrows():

            values = []

            for value in row:

                if isinstance(value, float):
                    values.append(
                        round(value, 2)
                    )
                else:
                    values.append(value)

            self.tree.insert(
                "",
                "end",
                values=values
            )

    # ============================================================
    # CLEAR TREEVIEW
    # ============================================================

    def clear_treeview(self):

        for item in self.tree.get_children():
            self.tree.delete(item)

        self.tree["columns"] = ()

    # ============================================================
    # EXPORT REPORT
    # ============================================================

    def export_report(self):

        if self.report_df is None:

            messagebox.showerror(
                "Error",
                "Please generate a report first."
            )
            return

        if not self.file_path:

            messagebox.showerror(
                "Error",
                "Input file path is not available."
            )
            return

        try:

            folder = os.path.dirname(
                self.file_path
            )

            base_name = os.path.splitext(
                os.path.basename(
                    self.file_path
                )
            )[0]

            export_format = self.export_format_var.get()

            # Export Excel
            if export_format == "Excel (.xlsx)":

                output_path = os.path.join(
                    folder,
                    f"{base_name}_report.xlsx"
                )

                self.report_df.to_excel(
                    output_path,
                    index=False
                )

            # Export CSV
            elif export_format == "CSV (.csv)":

                output_path = os.path.join(
                    folder,
                    f"{base_name}_report.csv"
                )

                self.report_df.to_csv(
                    output_path,
                    index=False
                )

            messagebox.showinfo(
                "Export Successful",
                f"Report saved successfully.\n\n"
                f"{output_path}"
            )

            self.status_label.config(
                text=f"Report exported: {output_path}"
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Unable to export report.\n\n{error}"
            )

    # ============================================================
    # PREVIEW CHART
    # ============================================================

    def preview_chart(self):

        if self.report_df is None:

            messagebox.showerror(
                "Error",
                "Please generate a report first."
            )
            return

        chart_type = self.chart_type_var.get()

        if not chart_type:

            messagebox.showerror(
                "Error",
                "Please select a chart type."
            )
            return

        try:

            # Clear previous chart
            self.clear_chart()

            # Create new window
            chart_window = tk.Toplevel(
                self.root
            )

            chart_window.title(
                "Chart Preview"
            )

            chart_window.geometry(
                "900x650"
            )

            chart_window.minsize(
                700,
                500
            )

            x_column = self.report_df.columns[0]
            y_column = self.report_df.columns[1]

            # Create figure
            figure, axis = plt.subplots(
                figsize=(9, 6),
                constrained_layout=True
            )

            # -----------------------------
            # BAR CHART
            # -----------------------------
            if chart_type == "Bar Chart":

                axis.barh(
                    self.report_df[x_column].astype(str),
                    self.report_df[y_column]
                )

                axis.set_xlabel(y_column)
                axis.set_ylabel(x_column)
                axis.set_title("Bar Chart")

                axis.invert_yaxis()

            # -----------------------------
            # COLUMN CHART
            # -----------------------------
            elif chart_type == "Column Chart":

                axis.bar(
                    self.report_df[x_column].astype(str),
                    self.report_df[y_column]
                )

                axis.set_xlabel(x_column)
                axis.set_ylabel(y_column)
                axis.set_title("Column Chart")

                plt.setp(
                    axis.get_xticklabels(),
                    rotation=45,
                    ha="right"
                )

            # -----------------------------
            # LINE CHART
            # -----------------------------
            elif chart_type == "Line Chart":

                axis.plot(
                    self.report_df[x_column].astype(str),
                    self.report_df[y_column],
                    marker="o"
                )

                axis.set_xlabel(x_column)
                axis.set_ylabel(y_column)
                axis.set_title("Line Chart")

                plt.setp(
                    axis.get_xticklabels(),
                    rotation=45,
                    ha="right"
                )

            # -----------------------------
            # PIE CHART
            # -----------------------------
            elif chart_type == "Pie Chart":

                axis.pie(
                    self.report_df[y_column],
                    labels=self.report_df[
                        x_column
                    ].astype(str),
                    autopct="%1.1f%%",
                    startangle=90
                )

                axis.set_title("Pie Chart")

            # Save reference
            self.current_figure = figure

            # Embed chart into Tkinter
            self.chart_canvas = FigureCanvasTkAgg(
                figure,
                master=chart_window
            )

            self.chart_canvas.draw()

            self.chart_canvas.get_tk_widget().pack(
                fill="both",
                expand=True
            )

            self.status_label.config(
                text="Chart generated successfully."
            )

        except Exception as error:

            messagebox.showerror(
                "Chart Error",
                f"Unable to generate chart.\n\n{error}"
            )

    # ============================================================
    # EXPORT CHART
    # ============================================================

    def export_chart(self):

        if self.current_figure is None:

            messagebox.showerror(
                "Error",
                "Please preview a chart first."
            )
            return

        if not self.file_path:

            messagebox.showerror(
                "Error",
                "Input file path is not available."
            )
            return

        try:

            folder = os.path.dirname(
                self.file_path
            )

            base_name = os.path.splitext(
                os.path.basename(
                    self.file_path
                )
            )[0]

            output_path = os.path.join(
                folder,
                f"{base_name}_chart.png"
            )

            self.current_figure.savefig(
                output_path,
                dpi=300,
                bbox_inches="tight"
            )

            messagebox.showinfo(
                "Export Successful",
                f"Chart saved successfully.\n\n"
                f"{output_path}"
            )

            self.status_label.config(
                text=f"Chart exported: {output_path}"
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Unable to export chart.\n\n{error}"
            )

    # ============================================================
    # CLEAR CHART
    # ============================================================

    def clear_chart(self):

        if self.current_figure is not None:

            plt.close(
                self.current_figure
            )

            self.current_figure = None
            self.chart_canvas = None


# ================================================================
# MAIN PROGRAM
# ================================================================

if __name__ == "__main__":

    root = tk.Tk()

    app = DataAnalysisApp(root)

    root.mainloop()
