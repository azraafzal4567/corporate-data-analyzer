"""
Corporate Data Analysis Tool
----------------------------
GUI-based Excel/CSV data analysis software.

Technologies:
    - Tkinter       -> GUI
    - Pandas        -> Data processing
    - Matplotlib    -> Charts
    - OpenPyXL      -> Excel file reading/writing

PyInstaller example:
    pyinstaller --onefile --windowed data_analysis_tool.py
"""

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

        # ---------------------------------------------------------
        # Main Window Configuration
        # ---------------------------------------------------------
        self.root.title("Corporate Data Analysis Tool")

        # Detect the available laptop/screen size
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Start the application at a responsive size
        window_width = int(screen_width * 0.90)
        window_height = int(screen_height * 0.85)

        self.root.geometry(f"{window_width}x{window_height}")
        self.root.minsize(900, 550)

        # ---------------------------------------------------------
        # Application Variables
        # ---------------------------------------------------------
        self.file_path = ""
        self.df = None
        self.report_df = None
        self.figure = None
        self.canvas = None

        # ---------------------------------------------------------
        # Styling
        # ---------------------------------------------------------
        self.setup_style()

        # ---------------------------------------------------------
        # Build GUI
        # ---------------------------------------------------------
        self.create_gui()

    # =============================================================
    # STYLE
    # =============================================================
    def setup_style(self):
        style = ttk.Style()

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 20, "bold")
        )

        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10)
        )

        style.configure(
            "Section.TLabel",
            font=("Segoe UI", 12, "bold")
        )

        style.configure(
            "TButton",
            font=("Segoe UI", 10),
            padding=7
        )

        style.configure(
            "TCombobox",
            padding=5
        )

        style.configure(
            "Treeview",
            rowheight=28,
            font=("Segoe UI", 9)
        )

        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold")
        )

    # =============================================================
    # MAIN GUI
    # =============================================================
    def create_gui(self):

        # ---------------------------------------------------------
        # Responsive Layout
        # ---------------------------------------------------------
        # The upper controls are placed inside a scrollable area.
        # The analysis output stays visible in the lower part of the window.
        # This prevents the Preview area from disappearing on small screens.
        self.root.grid_rowconfigure(0, weight=0)
        self.root.grid_rowconfigure(1, weight=1, minsize=300)
        self.root.grid_rowconfigure(2, weight=0)
        self.root.grid_columnconfigure(0, weight=1)

        # Keep the control panel compact so the Analysis Output always
        # receives a large, visible area. The controls themselves remain
        # scrollable on smaller screens.
        controls_height = 340

        controls_area = ttk.Frame(self.root, height=controls_height)
        controls_area.grid(row=0, column=0, sticky="ew")
        controls_area.grid_propagate(False)
        controls_area.grid_columnconfigure(0, weight=1)
        controls_area.grid_rowconfigure(0, weight=1)

        controls_canvas = tk.Canvas(
            controls_area,
            highlightthickness=0,
            borderwidth=0
        )
        controls_canvas.grid(row=0, column=0, sticky="nsew")

        controls_scroll = ttk.Scrollbar(
            controls_area,
            orient="vertical",
            command=controls_canvas.yview
        )
        controls_scroll.grid(row=0, column=1, sticky="ns")

        controls_canvas.configure(yscrollcommand=controls_scroll.set)

        controls_inner = ttk.Frame(controls_canvas)
        controls_window = controls_canvas.create_window(
            (0, 0),
            window=controls_inner,
            anchor="nw"
        )

        def update_controls_scrollregion(event=None):
            controls_canvas.configure(
                scrollregion=controls_canvas.bbox("all")
            )

        def resize_controls_width(event):
            controls_canvas.itemconfigure(
                controls_window,
                width=event.width
            )

        controls_inner.bind("<Configure>", update_controls_scrollregion)
        controls_canvas.bind("<Configure>", resize_controls_width)

        # Mouse-wheel scrolling for the controls area.
        def scroll_controls(event):
            controls_canvas.yview_scroll(
                int(-1 * (event.delta / 120)),
                "units"
            )

        controls_canvas.bind_all("<MouseWheel>", scroll_controls)

        # ---------------------------------------------------------
        # Header
        # ---------------------------------------------------------
        header = ttk.Frame(controls_inner, padding=(15, 8))
        header.pack(fill="x")

        title = ttk.Label(
            header,
            text="Corporate Data Analysis Tool",
            style="Title.TLabel"
        )
        title.pack(anchor="w")

        subtitle = ttk.Label(
            header,
            text="Analyze CSV and Excel files without writing Python code.",
            style="Subtitle.TLabel"
        )
        subtitle.pack(anchor="w", pady=(3, 0))

        # ---------------------------------------------------------
        # File Selection Section
        # ---------------------------------------------------------
        file_frame = ttk.LabelFrame(
            controls_inner,
            text="1. File Selection",
            padding=15
        )
        file_frame.pack(
            fill="x",
            padx=15,
            pady=3
        )

        self.file_entry = ttk.Entry(
            file_frame,
            width=90
        )
        self.file_entry.grid(
            row=0,
            column=0,
            padx=(0, 10),
            sticky="ew"
        )

        browse_button = ttk.Button(
            file_frame,
            text="Browse",
            command=self.browse_file
        )
        browse_button.grid(
            row=0,
            column=1,
            padx=5
        )

        read_button = ttk.Button(
            file_frame,
            text="Read",
            command=self.read_file
        )
        read_button.grid(
            row=0,
            column=2,
            padx=5
        )

        file_frame.columnconfigure(0, weight=1)

        # ---------------------------------------------------------
        # Dataset Information Section
        # ---------------------------------------------------------
        info_frame = ttk.LabelFrame(
            controls_inner,
            text="Dataset Information",
            padding=12
        )
        info_frame.pack(
            fill="x",
            padx=15,
            pady=3
        )

        self.rows_label = ttk.Label(
            info_frame,
            text="Rows: -"
        )
        self.rows_label.grid(
            row=0,
            column=0,
            padx=15,
            sticky="w"
        )

        self.columns_label = ttk.Label(
            info_frame,
            text="Columns: -"
        )
        self.columns_label.grid(
            row=0,
            column=1,
            padx=15,
            sticky="w"
        )

        self.headings_label = ttk.Label(
            info_frame,
            text="Column Headings: -"
        )
        self.headings_label.grid(
            row=1,
            column=0,
            columnspan=2,
            padx=15,
            pady=(8, 0),
            sticky="w"
        )

        # ---------------------------------------------------------
        # Report Builder
        # ---------------------------------------------------------
        report_frame = ttk.LabelFrame(
            controls_inner,
            text="2. Report Builder",
            padding=15
        )
        report_frame.pack(
            fill="x",
            padx=15,
            pady=3
        )

        # Group By
        ttk.Label(
            report_frame,
            text="Group By Column:"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.group_var = tk.StringVar()

        self.group_combo = ttk.Combobox(
            report_frame,
            textvariable=self.group_var,
            state="readonly",
            width=12
        )
        self.group_combo.grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        # Aggregation
        ttk.Label(
            report_frame,
            text="Aggregation:"
        ).grid(
            row=0,
            column=2,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.aggregation_var = tk.StringVar()

        aggregation_values = [
            "sum",
            "mean",
            "average",
            "max",
            "min",
            "count",
            "median"
        ]

        self.aggregation_combo = ttk.Combobox(
            report_frame,
            textvariable=self.aggregation_var,
            values=aggregation_values,
            state="readonly",
            width=18
        )
        self.aggregation_combo.grid(
            row=0,
            column=3,
            padx=5,
            pady=5
        )

        # Value Column
        ttk.Label(
            report_frame,
            text="Value Column:"
        ).grid(
            row=0,
            column=4,
            padx=5,
            pady=5,
            sticky="w"
        )

        self.value_var = tk.StringVar()

        self.value_combo = ttk.Combobox(
            report_frame,
            textvariable=self.value_var,
            state="readonly",
            width=18
        )
        self.value_combo.grid(
            row=0,
            column=5,
            padx=5,
            pady=5
        )

        # Preview Report
        preview_report_button = ttk.Button(
            report_frame,
            text="Preview Report",
            command=self.preview_report
        )
        preview_report_button.grid(
            row=1,
            column=0,
            columnspan=2,
            pady=10
        )

        # Export Format
        ttk.Label(
            report_frame,
            text="Export Format:"
        ).grid(
            row=1,
            column=2,
            padx=5
        )

        self.export_format_var = tk.StringVar(
            value="Excel (.xlsx)"
        )

        self.export_combo = ttk.Combobox(
            report_frame,
            textvariable=self.export_format_var,
            values=[
                "Excel (.xlsx)",
                "CSV (.csv)"
            ],
            state="readonly",
            width=18
        )
        self.export_combo.grid(
            row=1,
            column=3,
            padx=5
        )

        export_report_button = ttk.Button(
            report_frame,
            text="Export Report",
            command=self.export_report
        )
        export_report_button.grid(
            row=1,
            column=4,
            columnspan=2,
            padx=5
        )

        # ---------------------------------------------------------
        # Chart Builder
        # ---------------------------------------------------------
        chart_frame = ttk.LabelFrame(
            controls_inner,
            text="3. Chart Builder",
            padding=15
        )
        chart_frame.pack(
            fill="x",
            padx=15,
            pady=3
        )

        ttk.Label(
            chart_frame,
            text="Chart Type:"
        ).grid(
            row=0,
            column=0,
            padx=5,
            pady=5
        )

        self.chart_var = tk.StringVar()

        self.chart_combo = ttk.Combobox(
            chart_frame,
            textvariable=self.chart_var,
            values=[
                "Bar chart",
                "Column chart",
                "Line chart",
                "Pie chart"
            ],
            state="readonly",
            width=15
        )
        self.chart_combo.grid(
            row=0,
            column=1,
            padx=5,
            pady=5
        )

        preview_chart_button = ttk.Button(
            chart_frame,
            text="Preview Chart",
            command=self.preview_chart
        )
        preview_chart_button.grid(
            row=0,
            column=2,
            padx=10
        )

        export_chart_button = ttk.Button(
            chart_frame,
            text="Export Chart",
            command=self.export_chart
        )
        export_chart_button.grid(
            row=0,
            column=3,
            padx=10
        )


        # ---------------------------------------------------------
        # Output Area
        # ---------------------------------------------------------
        output_frame = ttk.LabelFrame(
            self.root,
            text="4. Analysis Output",
            padding=8
        )
        output_frame.grid(
            row=1,
            column=0,
            sticky="nsew",
            padx=15,
            pady=(3, 8)
        )

        self.notebook = ttk.Notebook(output_frame)
        self.notebook.pack(
            fill="both",
            expand=True
        )

        # ---------------------------------------------------------
        # Report Tab
        # ---------------------------------------------------------
        report_tab = ttk.Frame(self.notebook)
        self.notebook.add(
            report_tab,
            text="Report"
        )

        table_container = ttk.Frame(report_tab)
        table_container.pack(
            fill="both",
            expand=True
        )

        self.report_tree = ttk.Treeview(
            table_container,
            show="headings"
        )

        vertical_scroll = ttk.Scrollbar(
            table_container,
            orient="vertical",
            command=self.report_tree.yview
        )

        horizontal_scroll = ttk.Scrollbar(
            table_container,
            orient="horizontal",
            command=self.report_tree.xview
        )

        self.report_tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set
        )

        self.report_tree.grid(
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

        table_container.rowconfigure(
            0,
            weight=1
        )

        table_container.columnconfigure(
            0,
            weight=1
        )

        # ---------------------------------------------------------
        # Chart Tab
        # ---------------------------------------------------------
        chart_tab = ttk.Frame(self.notebook)
        self.notebook.add(
            chart_tab,
            text="Chart"
        )

        self.chart_container = chart_tab

        # ---------------------------------------------------------
        # Status Bar
        # ---------------------------------------------------------
        self.status_var = tk.StringVar(
            value="Ready"
        )

        status_bar = ttk.Label(
            self.root,
            textvariable=self.status_var,
            relief="sunken",
            anchor="w",
            padding=5
        )
        status_bar.grid(
            row=2,
            column=0,
            sticky="ew"
        )

    # =============================================================
    # BROWSE FILE
    # =============================================================
    def browse_file(self):

        file_path = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                (
                    "Data Files",
                    "*.csv *.xlsx *.xls"
                ),
                (
                    "CSV Files",
                    "*.csv"
                ),
                (
                    "Excel Files",
                    "*.xlsx *.xls"
                ),
                (
                    "All Files",
                    "*.*"
                )
            ]
        )

        if not file_path:
            return

        self.file_path = file_path

        self.file_entry.delete(
            0,
            tk.END
        )

        self.file_entry.insert(
            0,
            file_path
        )

        # Reset old data when a new file is selected
        self.reset_outputs()

        self.status_var.set(
            "File selected. Click Read to load the dataset."
        )

    # =============================================================
    # READ FILE
    # =============================================================
    def read_file(self):

        if not self.file_path:

            messagebox.showerror(
                "File Required",
                "Please select a CSV or Excel file first."
            )

            return

        try:

            extension = os.path.splitext(
                self.file_path
            )[1].lower()

            if extension == ".csv":

                self.df = pd.read_csv(
                    self.file_path
                )

            elif extension in [".xlsx", ".xls"]:

                self.df = pd.read_excel(
                    self.file_path
                )

            else:

                messagebox.showerror(
                    "Unsupported File",
                    "Please select a CSV or Excel file."
                )

                return

            # Clean column names
            self.df.columns = [
                str(column).strip()
                for column in self.df.columns
            ]

            # Detect columns
            self.detect_columns()

            # Dataset information
            total_rows = len(self.df)
            total_columns = len(self.df.columns)

            headings = ", ".join(
                map(str, self.df.columns)
            )

            self.rows_label.config(
                text=f"Rows: {total_rows:,}"
            )

            self.columns_label.config(
                text=f"Columns: {total_columns:,}"
            )

            self.headings_label.config(
                text=f"Column Headings: {headings}"
            )

            self.status_var.set(
                f"Dataset loaded successfully: "
                f"{total_rows:,} rows × "
                f"{total_columns:,} columns"
            )

            messagebox.showinfo(
                "Success",
                "Dataset loaded successfully."
            )

        except Exception as error:

            self.df = None

            messagebox.showerror(
                "Read Error",
                f"Could not read the file.\n\n{error}"
            )

            self.status_var.set(
                "Error while reading the file."
            )

    # =============================================================
    # COLUMN DETECTION
    # =============================================================
    def detect_columns(self):

        if self.df is None:
            return

        text_columns = []
        numeric_columns = []

        for column in self.df.columns:

            series = self.df[column]

            # Already numeric
            if pd.api.types.is_numeric_dtype(series):

                numeric_columns.append(column)

            else:

                # Try numeric conversion
                converted = pd.to_numeric(
                    series,
                    errors="coerce"
                )

                # If most non-empty values can be converted
                non_empty_count = series.notna().sum()

                if non_empty_count > 0:

                    numeric_count = converted.notna().sum()

                    numeric_ratio = (
                        numeric_count /
                        non_empty_count
                    )

                    if numeric_ratio >= 0.80:

                        numeric_columns.append(column)

                    else:

                        text_columns.append(column)

                else:

                    text_columns.append(column)

        self.text_columns = text_columns
        self.numeric_columns = numeric_columns

        # Update dropdowns
        self.group_combo["values"] = text_columns
        self.value_combo["values"] = numeric_columns

        self.group_var.set("")
        self.value_var.set("")
        self.aggregation_var.set("")
        self.chart_var.set("")

    # =============================================================
    # PREVIEW REPORT
    # =============================================================
    def preview_report(self):

        # Check file
        if not self.file_path:

            messagebox.showerror(
                "File Required",
                "Please select a file first."
            )

            return

        # Check dataset
        if self.df is None:

            messagebox.showerror(
                "Read Required",
                "Please click the Read button before building a report."
            )

            return

        # Get selections
        group_column = self.group_var.get()
        aggregation = self.aggregation_var.get()
        value_column = self.value_var.get()

        # Validate Group By
        if not group_column:

            messagebox.showerror(
                "Selection Required",
                "Please select a Group By Column."
            )

            return

        # Validate aggregation
        if not aggregation:

            messagebox.showerror(
                "Selection Required",
                "Please select an Aggregation Method."
            )

            return

        # Validate value column
        if not value_column:

            messagebox.showerror(
                "Selection Required",
                "Please select a Value Column."
            )

            return

        try:

            working_df = self.df.copy()

            # -----------------------------------------------------
            # Convert numeric-looking value column
            # -----------------------------------------------------
            working_df[value_column] = pd.to_numeric(
                working_df[value_column],
                errors="coerce"
            )

            # Remove rows where group is missing
            working_df = working_df[
                working_df[group_column].notna()
            ]

            # -----------------------------------------------------
            # GroupBy
            # -----------------------------------------------------
            grouped = working_df.groupby(
                group_column,
                dropna=False
            )[value_column]

            # -----------------------------------------------------
            # Aggregation
            # -----------------------------------------------------
            if aggregation == "sum":

                result = grouped.sum()

            elif aggregation in ["mean", "average"]:

                result = grouped.mean()

            elif aggregation == "max":

                result = grouped.max()

            elif aggregation == "min":

                result = grouped.min()

            elif aggregation == "count":

                result = grouped.count()

            elif aggregation == "median":

                result = grouped.median()

            else:

                messagebox.showerror(
                    "Invalid Aggregation",
                    "Selected aggregation method is not supported."
                )

                return

            # -----------------------------------------------------
            # Convert Series to DataFrame
            # -----------------------------------------------------
            self.report_df = result.reset_index()

            # Rename result column
            self.report_df.rename(
                columns={
                    value_column: f"{aggregation}_{value_column}"
                },
                inplace=True
            )

            result_column = (
                f"{aggregation}_{value_column}"
            )

            # Sort descending
            self.report_df.sort_values(
                by=result_column,
                ascending=False,
                inplace=True
            )

            self.report_df.reset_index(
                drop=True,
                inplace=True
            )

            # -----------------------------------------------------
            # Reset old outputs
            # -----------------------------------------------------
            self.clear_report_table()
            self.clear_chart()

            # -----------------------------------------------------
            # Display result
            # -----------------------------------------------------
            self.display_report()

            self.status_var.set(
                "Report generated successfully."
            )

            # Automatically switch to Report tab
            self.notebook.select(0)

        except Exception as error:

            messagebox.showerror(
                "Report Error",
                f"Could not generate the report.\n\n{error}"
            )

    # =============================================================
    # DISPLAY REPORT
    # =============================================================
    def display_report(self):

        if self.report_df is None:
            return

        self.clear_report_table()

        # Configure columns
        columns = list(
            self.report_df.columns
        )

        self.report_tree["columns"] = columns

        for column in columns:

            self.report_tree.heading(
                column,
                text=str(column)
            )

            self.report_tree.column(
                column,
                width=180,
                minwidth=100,
                anchor="center"
            )

        # Insert rows
        for _, row in self.report_df.iterrows():

            values = []

            for value in row:

                if pd.isna(value):
                    values.append("")
                elif isinstance(value, float):
                    values.append(
                        f"{value:,.2f}"
                    )
                else:
                    values.append(
                        str(value)
                    )

            self.report_tree.insert(
                "",
                tk.END,
                values=values
            )

    # =============================================================
    # CLEAR REPORT TABLE
    # =============================================================
    def clear_report_table(self):

        for item in self.report_tree.get_children():

            self.report_tree.delete(
                item
            )

        self.report_tree["columns"] = ()

    # =============================================================
    # EXPORT REPORT
    # =============================================================
    def export_report(self):

        if self.df is None:

            messagebox.showerror(
                "Read Required",
                "Please select and read a file first."
            )

            return

        if self.report_df is None:

            messagebox.showerror(
                "Report Required",
                "Please preview a report before exporting it."
            )

            return

        try:

            input_directory = os.path.dirname(
                self.file_path
            )

            input_filename = os.path.splitext(
                os.path.basename(
                    self.file_path
                )
            )[0]

            aggregation = self.aggregation_var.get()
            group_column = self.group_var.get()

            safe_group = self.make_safe_filename(
                group_column
            )

            safe_aggregation = self.make_safe_filename(
                aggregation
            )

            base_name = (
                f"{input_filename}_"
                f"{safe_group}_"
                f"{safe_aggregation}_report"
            )

            selected_format = (
                self.export_format_var.get()
            )

            if selected_format == "Excel (.xlsx)":

                output_path = os.path.join(
                    input_directory,
                    base_name + ".xlsx"
                )

                self.report_df.to_excel(
                    output_path,
                    index=False
                )

            else:

                output_path = os.path.join(
                    input_directory,
                    base_name + ".csv"
                )

                self.report_df.to_csv(
                    output_path,
                    index=False,
                    encoding="utf-8-sig"
                )

            messagebox.showinfo(
                "Export Successful",
                f"Report exported successfully.\n\n"
                f"Location:\n{output_path}"
            )

            self.status_var.set(
                f"Report exported: {output_path}"
            )

        except Exception as error:

            messagebox.showerror(
                "Export Error",
                f"Could not export the report.\n\n{error}"
            )

    # =============================================================
    # PREVIEW CHART
    # =============================================================
    def preview_chart(self):

        if self.df is None:

            messagebox.showerror(
                "Read Required",
                "Please select and read a file first."
            )

            return

        if self.report_df is None:

            messagebox.showerror(
                "Report Required",
                "Please preview a report before creating a chart."
            )

            return

        chart_type = self.chart_var.get()

        if not chart_type:

            messagebox.showerror(
                "Selection Required",
                "Please select a chart type."
            )

            return

        try:

            self.clear_chart()

            group_column = self.group_var.get()

            aggregation = self.aggregation_var.get()

            value_column = self.value_var.get()

            result_column = (
                f"{aggregation}_{value_column}"
            )

            chart_data = self.report_df.copy()

            # Limit pie chart categories if there are too many
            if chart_type == "Pie chart" and len(chart_data) > 15:

                chart_data = chart_data.head(15)

            # -----------------------------------------------------
            # Create Matplotlib Figure
            # -----------------------------------------------------
            # Fit the Matplotlib figure to the actual space available
            # in the chart preview area. This prevents the chart title,
            # x-axis labels, or bottom area from being cropped on
            # smaller laptop screens.
            self.chart_container.update_idletasks()

            container_width = self.chart_container.winfo_width()
            container_height = self.chart_container.winfo_height()

            # Keep safe minimums while adapting to the real container size.
            figure_width = max(7.0, container_width / 100)
            figure_height = max(3.2, container_height / 100)

            self.figure = plt.Figure(
                figsize=(figure_width, figure_height),
                dpi=100,
                constrained_layout=True
            )

            ax = self.figure.add_subplot(111)

            # -----------------------------------------------------
            # Bar Chart - horizontal
            # -----------------------------------------------------
            if chart_type == "Bar chart":

                ax.barh(
                    chart_data[group_column].astype(str),
                    chart_data[result_column]
                )

                ax.set_xlabel(
                    result_column
                )

                ax.set_ylabel(
                    group_column
                )

                ax.invert_yaxis()

            # -----------------------------------------------------
            # Column Chart - vertical
            # -----------------------------------------------------
            elif chart_type == "Column chart":

                ax.bar(
                    chart_data[group_column].astype(str),
                    chart_data[result_column]
                )

                ax.set_xlabel(
                    group_column
                )

                ax.set_ylabel(
                    result_column
                )

                plt.setp(
                    ax.get_xticklabels(),
                    rotation=45,
                    ha="right"
                )

            # -----------------------------------------------------
            # Line Chart
            # -----------------------------------------------------
            elif chart_type == "Line chart":

                ax.plot(
                    chart_data[group_column].astype(str),
                    chart_data[result_column],
                    marker="o"
                )

                ax.set_xlabel(
                    group_column
                )

                ax.set_ylabel(
                    result_column
                )

                plt.setp(
                    ax.get_xticklabels(),
                    rotation=45,
                    ha="right"
                )

            # -----------------------------------------------------
            # Pie Chart
            # -----------------------------------------------------
            elif chart_type == "Pie chart":

                ax.pie(
                    chart_data[result_column],
                    labels=chart_data[
                        group_column
                    ].astype(str),
                    autopct="%1.1f%%",
                    startangle=90
                )

                ax.axis("equal")

            else:

                messagebox.showerror(
                    "Invalid Chart",
                    "Selected chart type is not supported."
                )

                return

            # -----------------------------------------------------
            # Chart Title
            # -----------------------------------------------------
            ax.set_title(
                f"{aggregation.title()} of "
                f"{value_column} by "
                f"{group_column}"
            )

            # Grid for non-pie charts
            if chart_type != "Pie chart":

                ax.grid(
                    axis="y",
                    alpha=0.3
                )

            # constrained_layout=True automatically reserves space
            # for titles, axis labels, and tick labels.
            # No fixed-size tight_layout call is needed here.

            # -----------------------------------------------------
            # Display Figure in Tkinter
            # -----------------------------------------------------
            self.canvas = FigureCanvasTkAgg(
                self.figure,
                master=self.chart_container
            )

            self.canvas.draw()

            self.canvas.get_tk_widget().pack(
                fill="both",
                expand=True
            )

            # Switch to chart tab
            self.notebook.select(1)

            self.status_var.set(
                "Chart generated successfully."
            )

        except Exception as error:

            messagebox.showerror(
                "Chart Error",
                f"Could not generate chart.\n\n{error}"
            )

    # =============================================================
    # CLEAR CHART
    # =============================================================
    def clear_chart(self):

        if self.canvas is not None:

            self.canvas.get_tk_widget().destroy()

            self.canvas = None

        if self.figure is not None:

            plt.close(
                self.figure
            )

            self.figure = None

    # =============================================================
    # EXPORT CHART
    # =============================================================
    def export_chart(self):

        if self.df is None:

            messagebox.showerror(
                "Read Required",
                "Please select and read a file first."
            )

            return

        if self.report_df is None:

            messagebox.showerror(
                "Report Required",
                "Please preview a report first."
            )

            return

        if self.figure is None:

            messagebox.showerror(
                "Chart Required",
                "Please preview a chart before exporting it."
            )

            return

        try:

            input_directory = os.path.dirname(
                self.file_path
            )

            input_filename = os.path.splitext(
                os.path.basename(
                    self.file_path
                )
            )[0]

            chart_type = self.chart_var.get()

            safe_chart_type = self.make_safe_filename(
                chart_type
            )

            output_path = os.path.join(
                input_directory,
                f"{input_filename}_{safe_chart_type}.png"
            )

            self.figure.savefig(
                output_path,
                dpi=300,
                bbox_inches="tight"
            )

            messagebox.showinfo(
                "Chart Exported",
                f"Chart exported successfully.\n\n"
                f"Location:\n{output_path}"
            )

            self.status_var.set(
                f"Chart exported: {output_path}"
            )

        except Exception as error:

            messagebox.showerror(
                "Chart Export Error",
                f"Could not export the chart.\n\n{error}"
            )

    # =============================================================
    # RESET OUTPUTS
    # =============================================================
    def reset_outputs(self):

        # Reset dataset
        self.df = None
        self.report_df = None

        # Reset information
        self.rows_label.config(
            text="Rows: -"
        )

        self.columns_label.config(
            text="Columns: -"
        )

        self.headings_label.config(
            text="Column Headings: -"
        )

        # Reset dropdowns
        self.group_combo["values"] = []
        self.value_combo["values"] = []

        self.group_var.set("")
        self.value_var.set("")
        self.aggregation_var.set("")
        self.chart_var.set("")

        # Clear report
        self.clear_report_table()

        # Clear chart
        self.clear_chart()

    # =============================================================
    # SAFE FILE NAME
    # =============================================================
    @staticmethod
    def make_safe_filename(text):

        text = str(text)

        # Remove invalid Windows filename characters
        text = re.sub(
            r'[<>:"/\\|?*]',
            "_",
            text
        )

        # Replace spaces
        text = text.replace(
            " ",
            "_"
        )

        return text


# =================================================================
# APPLICATION ENTRY POINT
# =================================================================
def main():

    root = tk.Tk()

    app = DataAnalysisApp(
        root
    )

    root.mainloop()


if __name__ == "__main__":
    main()