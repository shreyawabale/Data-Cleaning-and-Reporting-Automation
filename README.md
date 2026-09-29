# Data Cleaning & Reporting Automation

## 📌 Project Overview

This project automates the data cleaning and reporting workflow using Python and the Tableau Superstore dataset.

The objective is to clean raw sales data, handle missing values and duplicate records, standardize data types and text values, and automatically generate reports and visual summaries.

This project demonstrates practical data preprocessing, quality checking, reporting automation, and data visualization.

---

## 🎯 Objectives

- Clean and preprocess raw sales data.
- Identify missing values and duplicate records.
- Standardize column names and data types.
- Handle missing values automatically.
- Standardize inconsistent text values.
- Generate data quality reports.
- Generate an automated Excel report.
- Create sales summary reports.
- Visualize sales performance by category.

---

## 📊 Dataset

Dataset: Tableau Superstore Dataset

The dataset contains sales transaction information including:

- Customer ID
- Customer Name
- Order ID
- Order Date
- Ship Date
- Sales
- Quantity
- Discount
- Profit
- Category
- Sub-Category
- Region
- State
- City

The dataset contains:

- 10,194 transaction records
- 21 columns

---

## 🔍 Data Cleaning Process

### 1. Data Loading

The Superstore Excel dataset is loaded using Pandas.

### 2. Data Quality Check

The project checks:

- Missing values
- Duplicate records
- Data types
- Column information

### 3. Column Name Cleaning

Column names are standardized by:

- Removing leading and trailing spaces
- Replacing spaces with underscores
- Replacing hyphens with underscores

### 4. Duplicate Removal

Duplicate records are identified and automatically removed.

For this dataset:

- Duplicate records found: 0

### 5. Missing Value Handling

Missing numerical values are replaced using the median value of the respective column.

Missing text values are replaced with:

`Unknown`

For this dataset:

- Missing values before cleaning: 0
- Missing values after cleaning: 0

### 6. Data Type Standardization

Important fields such as:

- Order Date
- Ship Date
- Sales
- Profit
- Quantity
- Discount

are converted into appropriate data types.

### 7. Text Standardization

Text fields are stripped of unnecessary leading and trailing spaces to improve consistency.

---

## 📈 Automated Reporting

The project automatically generates:

### Data Quality Report

An Excel workbook containing:

- Data quality before cleaning
- Data quality after cleaning
- Summary information
- Missing value information
- Sales summary

### CSV Reports

The project generates:

- Cleaned dataset
- Missing values report
- Summary report

### Visualization

A sales summary visualization is generated showing total sales by category.

---

## 📁 Generated Files

```text
Data-Cleaning-and-Reporting-Automation
│
├── data_cleaning_automation.py
├── cleaned_superstore.csv
├── missing_values_report.csv
├── summary_report.csv
├── data_quality_report.xlsx
├── sales_summary.png
└── README.md

📊 Results

The automation successfully processed the Superstore dataset.

Metric	Result
Original Rows	10,194
Original Columns	21
Duplicate Rows Removed	0
Final Rows	10,194
Final Columns	21
Missing Values Before	0
Missing Values After	0
🛠️ Technologies Used
Python
Pandas
NumPy
Matplotlib
OpenPyXL
Excel
CSV
Data Cleaning
Data Preprocessing
Data Visualization
Reporting Automation
▶️ How to Run
1. Clone the repository
git clone https://github.com/shreyawabale/Data-Cleaning-and-Reporting-Automation.git
2. Navigate to the project folder
cd Data-Cleaning-and-Reporting-Automation
3. Install required libraries
pip install pandas numpy matplotlib openpyxl xlrd
4. Run the Python script
python data_cleaning_automation.py
5. Generated reports

After execution, the following files will be generated:

cleaned_superstore.csv
missing_values_report.csv
summary_report.csv
data_quality_report.xlsx
sales_summary.png
🎯 Expected Outcome

This project demonstrates how data cleaning and reporting tasks can be automated using Python.

It reduces repetitive manual processing and provides structured reports and visual summaries that can support data analysis and business reporting.

👩‍💻 Author

Shreya Wabale

Computer Engineering Student
