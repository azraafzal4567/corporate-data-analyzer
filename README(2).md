# Corporate Data Analyzer

A GUI-based desktop application for analyzing CSV and Excel files without requiring users to write Python code.

## Overview

**Corporate Data Analyzer** is a Python desktop application designed to make basic data analysis and visualization easier for non-technical users.

Users can load a CSV or Excel file, explore the dataset, build summary reports, create charts, preview results inside the application, and export their analysis.

## Features

- Load CSV and Excel files
- View total rows and columns
- Display dataset column headings
- Automatically detect text and numeric columns
- Build reports using Group By and aggregation
- Sum, Mean, Max, Min, Count and Median
- Preview reports inside the application
- Create Bar, Column, Line and Pie charts
- Preview charts inside the application
- Export reports to Excel or CSV
- Export charts as PNG
- GUI designed for non-technical users
- Responsive layout for different laptop screen sizes
- Can be packaged as a standalone Windows `.exe` using PyInstaller

## Technologies Used

- Python
- Tkinter
- Pandas
- Matplotlib
- OpenPyXL
- xlrd
- PyInstaller

## Project Structure

```text
Corporate-Data-Analyzer/
│
├── project.py
├── README.md
└── requirements.txt
```

## Installation

Install the required libraries:

```bash
pip install -r requirements.txt
```

## Run the Application

```bash
python project.py
```

## Creating a Windows Executable

```bash
pyinstaller --onefile --windowed --name "Corporate Data Analyzer" project.py
```

The executable will be created inside the `dist` folder.

## How to Use

1. Open the application.
2. Click **Browse** and select a CSV or Excel file.
3. Click **Read** to load the dataset.
4. Review the dataset information.
5. Use **Report Builder** to create a summary report.
6. Click **Preview Report**.
7. Use **Chart Builder** to create a visualization.
8. Click **Preview Chart**.
9. Export the report or chart when required.

## Purpose

The main purpose of this project is to provide a simple graphical interface for basic data analysis so that users can work with data without directly writing or running Python code.

## Learning Outcomes

This project provided practical experience with:

- Python GUI development
- Tkinter
- Pandas DataFrames
- Data type detection
- Grouping and aggregation
- Data visualization
- CSV and Excel processing
- File handling and error handling
- Responsive GUI layout
- Packaging Python applications with PyInstaller

## Author

**Azra Afzal**

## Project Type

Python Desktop Application / Data Analysis Tool
