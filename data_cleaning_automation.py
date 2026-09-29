import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

print("=" * 60)
print("DATA CLEANING & REPORTING AUTOMATION")
print("=" * 60)

# ---------------------------------------------------------
# 1. LOAD DATASET
# ---------------------------------------------------------

print("\nLoading dataset...")

file_name = "sample_-_superstore.xls"

df = pd.read_excel(
    file_name,
    sheet_name="Orders",
    engine="xlrd"
)

print("Dataset loaded successfully!")
print("Original rows:", len(df))
print("Original columns:", len(df.columns))

# ---------------------------------------------------------
# 2. CREATE DATA QUALITY REPORT - BEFORE CLEANING
# ---------------------------------------------------------

print("\nChecking data quality...")

missing_before = df.isnull().sum()
duplicates_before = df.duplicated().sum()

quality_before = pd.DataFrame({
    "Column": df.columns,
    "Missing_Values": [df[col].isnull().sum() for col in df.columns],
    "Missing_Percentage": [
        round(df[col].isnull().mean() * 100, 2)
        for col in df.columns
    ],
    "Data_Type": [str(df[col].dtype) for col in df.columns]
})

# ---------------------------------------------------------
# 3. CLEAN COLUMN NAMES
# ---------------------------------------------------------

print("\nCleaning column names...")

df.columns = (
    df.columns
      .str.strip()
      .str.replace(" ", "_")
      .str.replace("-", "_")
)

# ---------------------------------------------------------
# 4. REMOVE DUPLICATES
# ---------------------------------------------------------

print("\nChecking duplicate records...")

duplicate_count = df.duplicated().sum()

print("Duplicate rows found:", duplicate_count)

df = df.drop_duplicates()

print("Duplicate rows removed:", duplicate_count)

# ---------------------------------------------------------
# 5. HANDLE MISSING VALUES
# ---------------------------------------------------------

print("\nHandling missing values...")

# Numeric columns: fill missing values with median
numeric_columns = df.select_dtypes(
    include=np.number
).columns

for column in numeric_columns:
    if df[column].isnull().sum() > 0:
        df[column] = df[column].fillna(
            df[column].median()
        )

# Text columns: fill missing values with "Unknown"
text_columns = df.select_dtypes(
    include=["object", "string"]
).columns

for column in text_columns:
    if df[column].isnull().sum() > 0:
        df[column] = df[column].fillna("Unknown")

# ---------------------------------------------------------
# 6. FIX DATA TYPES
# ---------------------------------------------------------

print("\nStandardizing data types...")

if "Order_Date" in df.columns:
    df["Order_Date"] = pd.to_datetime(
        df["Order_Date"],
        errors="coerce"
    )

if "Ship_Date" in df.columns:
    df["Ship_Date"] = pd.to_datetime(
        df["Ship_Date"],
        errors="coerce"
    )

# Convert important numerical columns
for column in ["Sales", "Profit", "Quantity", "Discount"]:
    if column in df.columns:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )

# ---------------------------------------------------------
# 7. HANDLE REMAINING MISSING VALUES
# ---------------------------------------------------------

for column in numeric_columns:
    if column in df.columns and df[column].isnull().sum() > 0:
        df[column] = df[column].fillna(
            df[column].median()
        )

for column in text_columns:
    if column in df.columns and df[column].isnull().sum() > 0:
        df[column] = df[column].fillna("Unknown")

# ---------------------------------------------------------
# 8. STANDARDIZE TEXT VALUES
# ---------------------------------------------------------

print("\nStandardizing text values...")

for column in df.select_dtypes(
    include=["object", "string"]
).columns:

    df[column] = df[column].astype(str).str.strip()

# ---------------------------------------------------------
# 9. FINAL DATA QUALITY CHECK
# ---------------------------------------------------------

missing_after = df.isnull().sum()
duplicates_after = df.duplicated().sum()

quality_after = pd.DataFrame({
    "Column": df.columns,
    "Missing_Values": [df[col].isnull().sum() for col in df.columns],
    "Missing_Percentage": [
        round(df[col].isnull().mean() * 100, 2)
        for col in df.columns
    ],
    "Data_Type": [str(df[col].dtype) for col in df.columns]
})

# ---------------------------------------------------------
# 10. SAVE CLEANED DATA
# ---------------------------------------------------------

print("\nSaving cleaned dataset...")

df.to_csv(
    "cleaned_superstore.csv",
    index=False
)

print("Saved: cleaned_superstore.csv")

# ---------------------------------------------------------
# 11. CREATE SUMMARY REPORT
# ---------------------------------------------------------

print("\nCreating summary report...")

summary_data = {
    "Metric": [
        "Original Rows",
        "Original Columns",
        "Duplicate Rows Removed",
        "Final Rows",
        "Final Columns",
        "Missing Values Before",
        "Missing Values After"
    ],
    "Value": [
        len(df) + duplicate_count,
        len(df.columns),
        duplicate_count,
        len(df),
        len(df.columns),
        missing_before.sum(),
        missing_after.sum()
    ]
}

summary_report = pd.DataFrame(summary_data)

summary_report.to_csv(
    "summary_report.csv",
    index=False
)

# ---------------------------------------------------------
# 12. MISSING VALUES REPORT
# ---------------------------------------------------------

missing_report = quality_before[
    ["Column", "Missing_Values", "Missing_Percentage"]
].copy()

missing_report.to_csv(
    "missing_values_report.csv",
    index=False
)

# ---------------------------------------------------------
# 13. SALES SUMMARY
# ---------------------------------------------------------

print("\nCreating sales summary...")

sales_summary = pd.DataFrame()

if "Category" in df.columns and "Sales" in df.columns:
    sales_summary = (
        df.groupby("Category")["Sales"]
          .sum()
          .sort_values(ascending=False)
          .reset_index()
    )

    sales_summary.columns = [
        "Category",
        "Total_Sales"
    ]

# ---------------------------------------------------------
# 14. AUTOMATED EXCEL REPORT
# ---------------------------------------------------------

print("\nGenerating automated Excel report...")

with pd.ExcelWriter(
    "data_quality_report.xlsx",
    engine="openpyxl"
) as writer:

    quality_before.to_excel(
        writer,
        sheet_name="Before_Cleaning",
        index=False
    )

    quality_after.to_excel(
        writer,
        sheet_name="After_Cleaning",
        index=False
    )

    summary_report.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    missing_report.to_excel(
        writer,
        sheet_name="Missing_Values",
        index=False
    )

    if not sales_summary.empty:
        sales_summary.to_excel(
            writer,
            sheet_name="Sales_Summary",
            index=False
        )

print("Saved: data_quality_report.xlsx")

# ---------------------------------------------------------
# 15. SALES VISUALIZATION
# ---------------------------------------------------------

if not sales_summary.empty:

    print("\nCreating sales visualization...")

    plt.figure(figsize=(10, 6))

    plt.bar(
        sales_summary["Category"],
        sales_summary["Total_Sales"]
    )

    plt.title("Sales by Category")
    plt.xlabel("Category")
    plt.ylabel("Total Sales")

    plt.tight_layout()

    plt.savefig(
        "sales_summary.png",
        dpi=300
    )

    plt.show()

# ---------------------------------------------------------
# 16. FINAL REPORT
# ---------------------------------------------------------

print("\n" + "=" * 60)
print("DATA CLEANING & REPORTING AUTOMATION COMPLETED")
print("=" * 60)

print("\nOriginal rows:", len(df) + duplicate_count)
print("Duplicates removed:", duplicate_count)
print("Final rows:", len(df))
print("Final columns:", len(df.columns))
print("Missing values before:", missing_before.sum())
print("Missing values after:", missing_after.sum())

print("\nFiles created:")
print("1. cleaned_superstore.csv")
print("2. missing_values_report.csv")
print("3. summary_report.csv")
print("4. data_quality_report.xlsx")
print("5. sales_summary.png")

print("\nAutomation completed successfully!")