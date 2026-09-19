# Python Data Analysis & Report Builder

A Python-based desktop application that allows users to load CSV and Excel datasets, inspect dataset information, build dynamic summary reports, and create visualizations through a graphical user interface.

## Project Overview

The application was developed as a practical Python data analysis project to make basic dataset exploration, reporting, and visualization easier through a single desktop interface.

Users can select a CSV or Excel file, view basic dataset information, choose grouping and aggregation options, generate a summary report, export the report, and visualize the generated results.

## Key Features

### 1. File Selection & Data Loading

* Supports CSV files.
* Supports Excel files (`.xlsx` and `.xls`).
* Allows users to browse and select a dataset through the graphical interface.
* Displays the selected file name.
* Loads the dataset using Pandas.

### 2. Dataset Information

After loading a dataset, the application displays:

* Total number of rows
* Total number of columns
* Column headings

The application also detects text and numeric columns dynamically based on the dataset structure.

### 3. Dynamic Report Builder

Users can create summary reports by selecting:

* A **Group By Column**
* An **Aggregation Method**
* A **Value Column**

Supported aggregation methods:

* Sum
* Mean
* Average
* Maximum
* Minimum
* Count
* Median

The generated report is sorted in descending order based on the calculated value and displayed inside the application.

### 4. Report Export

Generated reports can be exported in:

* Excel (`.xlsx`)
* CSV (`.csv`)

The exported report is automatically named based on the original dataset file name.

### 5. Chart Builder

The application can visualize generated reports using:

* Bar Chart
* Column Chart
* Line Chart
* Pie Chart

Charts are displayed in a separate preview window and can also be exported as PNG files.

### 6. Error Handling & Validation

The application includes validation and error messages for situations such as:

* No dataset selected
* Unsupported file format
* Missing report selections
* Attempting to export before generating a report
* Attempting to export a chart before previewing one

## Technologies Used

* **Python**
* **Pandas** — data loading, grouping, aggregation, and report generation
* **Tkinter** — graphical user interface
* **Matplotlib** — chart creation and visualization

## What I Practiced

Through this project, I practiced:

* Reading CSV and Excel datasets using Pandas
* Working with different dataset structures
* Detecting numeric and text columns dynamically
* Grouping and aggregating data
* Generating summary reports
* Sorting analytical results
* Exporting reports to Excel and CSV
* Creating Bar, Column, Line, and Pie charts
* Exporting charts as PNG files
* Building a desktop GUI using Tkinter
* Handling user input and application errors

## Testing

The application was tested with different CSV datasets to explore its reporting, export, and visualization features and to check how it handled datasets with different structures.

## Project File

`Data Analysis App.py`
