import pandas as pd
import os

# Full path to your Excel file
file_path = r"C:\Users\atten\Desktop\UTAR\Degree\Year 3\Sem 3\Data Visualisation and Reporting\Assignment\UHospital Sales and Purchase.xlsx"

# Load all sheets
sheets = pd.read_excel(file_path, sheet_name=None)

# Helper function: translate Python types into simple English
def type_to_english(py_type):
    if py_type == str:
        return "Text"
    elif py_type == float:
        return "Number"
    elif py_type == int:
        return "Whole Number"
    elif "datetime" in str(py_type).lower():
        return "Date/Time"
    else:
        return "Other"

all_results = {}

for sheet_name, df in sheets.items():
    issues_list = []

    # Duplicate column names
    if df.columns.duplicated().any():
        issues_list.append(["(Whole Sheet)", "Duplicate column names found"])

    # Completely empty rows → list row numbers
    empty_rows = df.index[df.isnull().all(axis=1)].tolist()
    for r in empty_rows:
        issues_list.append(["(Whole Sheet)", f"Row {r+2} is completely empty"])  
        # +2 so it matches Excel row number (since header is row 1)

    # Completely empty columns → list column names
    empty_cols = df.columns[df.isnull().all(axis=0)].tolist()
    for c in empty_cols:
        issues_list.append([c, "This column is completely empty"])

    # Column-level checks
    for col in df.columns:
        col_data = df[col]

        # Missing values
        missing = col_data.isnull().sum()
        if missing > 0 and missing < len(col_data):  
            issues_list.append([col, f"{missing} missing values"])

        # Mixed types
        types = col_data.map(type).value_counts()
        if len(types) > 1:
            type_summary = [f"{count} {type_to_english(t)} values" for t, count in types.items()]
            issues_list.append([col, "mixed data types: " + " and ".join(type_summary) + " (should be consistent)"])

        # Numbers stored as text
        if col_data.dtype == "object":
            numeric_as_text = col_data.str.isnumeric().sum()
            if numeric_as_text > 0:
                issues_list.append([col, f"{numeric_as_text} values look like numbers but are stored as text"])

        # Whitespace issues
        whitespace = col_data.astype(str).str.contains(r"^\s|\s$|  +", regex=True).sum()
        if whitespace > 0:
            issues_list.append([col, f"{whitespace} values with leading, trailing, or extra spaces"])

    # Duplicate rows → list row numbers
    dup_rows = df.index[df.duplicated()].tolist()
    for r in dup_rows:
        issues_list.append(["(Whole Sheet)", f"Row {r+2} is a duplicate of a previous row"])

    all_results[sheet_name] = issues_list

# Save report
output_path = os.path.join(os.path.dirname(file_path), "data_quality_report_2.xlsx")
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    for sheet, issues in all_results.items():
        if not issues:
            issues = [["(Whole Sheet)", "No issues found"]]
        pd.DataFrame(issues, columns=["Column Name", "Issue"]).to_excel(writer, sheet_name=sheet, index=False)

print(f"✅ Data quality check complete. Report saved here:\n{output_path}")
