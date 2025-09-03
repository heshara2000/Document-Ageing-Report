import pandas as pd
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font
import os

# --- Helper function: Excel serial date ---
def date_to_excel_serial(dt):
    if pd.isna(dt):
        return None
    base = datetime(1899, 12, 30).date()
    if isinstance(dt, datetime):
        dt = dt.date()
    elif isinstance(dt, pd.Timestamp):
        dt = dt.to_pydatetime().date()
    elif isinstance(dt, date):
        pass
    else:
        return None
    return (dt - base).days

# --- Sample data generator ---
def generate_sample_data(file_path):
    sample_data = {
        "Comapany": ["UN0100", "UN0200", "UN0300"],
        "Account": ["100100", "200200", "300300"],
        "Entry Date": ["2020-06-02", "2020-07-15", "2020-08-20"],
        "Document Date": ["2020-06-01", "2020-07-30", "2020-08-25"],
        "Document Type": ["DZ", "FZ", "DZ"],
        "Text": ["Sales invoice", "Customer refund", "Customer payment"],
        "Document currency": ["USD", "EUR", "LKR"],
        "Amount in doc. curr.": [1050.75, -500.00, 2000.00],
        "Local Currency": ["USD","USD","USD"],
        "Amount in local currency": [190000.00, -90000.00, 360000.00],
        "Year/month": ["2020/06", "2020/07", "2020/08"],
    }
    df = pd.DataFrame(sample_data)
    # Save as xlsx
    file_path = file_path.replace(".xls", ".xlsx")
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    df.to_excel(file_path, index=False, engine="openpyxl")
    print(f"✅ Sample data written to {file_path}")
    return file_path

# --- Step 1: Read input file ---
input_file = r"E:\dil_copies\Document-Ageing-Report-\data\export.xls"
FORCE_SAMPLE = True  # Set True to overwrite/generate sample every run

if FORCE_SAMPLE or not os.path.exists(input_file.replace(".xls", ".xlsx")):
    input_file = generate_sample_data(input_file)
else:
    input_file = input_file.replace(".xls", ".xlsx")

# Read file
df = pd.read_excel(input_file, engine="openpyxl")
print(f"📄 Reading from: {input_file}")
print(f"🔢 Rows loaded: {len(df)}")
print("🧾 Columns:", df.columns.tolist())

# Normalize column names
df.columns = (
    df.columns
    .str.replace(r"[ .]", "_", regex=True)
    .str.replace(r"_+", "_", regex=True)
    .str.strip("_")
)

# Parse document date
if "Document_Date" in df.columns:
    df["Document_Date"] = pd.to_datetime(df["Document_Date"], errors="coerce")

# --- Step 2: Today’s date & serial ---
today = datetime.now().date()
today_str = today.strftime("%d.%m.%Y")
today_serial = date_to_excel_serial(today)

# Add ageing column
df["Doc_Serial"] = df["Document_Date"].apply(date_to_excel_serial)
df["Doc_Ageing"] = today_serial - df["Doc_Serial"]

# --- Step 3: Create workbook & Summary sheet ---
wb = openpyxl.Workbook()
summary_ws = wb.active
summary_ws.title = "Summary"

# Title row
summary_ws.cell(row=2, column=2).value = f"Document Ageing Report as at {today_str}"
summary_ws.cell(row=2, column=2).font = Font(bold=True, size=14)

# Headers (row 4)
summary_headers = [
    "Company", "Account", "Document_currency", 
    "Amount_in_doc_curr", "Local_Currency", "Amount_in_local_currency"
]
header_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")
thin_border = Border(
    left=Side(style='thin'),
    right=Side(style='thin'),
    top=Side(style='thin'),
    bottom=Side(style='thin')
)
for col, header in enumerate(summary_headers, start=2):
    summary_ws.cell(row=4, column=col).value = header
    summary_ws.cell(row=4, column=col).fill = header_fill
    summary_ws.cell(row=4, column=col).border = thin_border

# --- Step 4: Build summary content ---
group = df.groupby(["Comapany", "Account", "Document_currency", "Local_Currency"])
sums = group.agg({
    "Amount_in_doc_curr": "sum",
    "Amount_in_local_currency": "sum"
}).reset_index()
sums = sums[abs(sums["Amount_in_local_currency"]) > 1e-5]

for i, row in sums.iterrows():
    summary_ws.cell(row=i+5, column=2).value = row["Comapany"]
    summary_ws.cell(row=i+5, column=3).value = row["Account"]
    summary_ws.cell(row=i+5, column=4).value = row["Document_currency"]
    summary_ws.cell(row=i+5, column=5).value = row["Amount_in_doc_curr"]
    summary_ws.cell(row=i+5, column=6).value = row["Local_Currency"]
    summary_ws.cell(row=i+5, column=7).value = row["Amount_in_local_currency"]
    for col in range(2, 8):
        summary_ws.cell(row=i+5, column=col).border = thin_border

# --- Step 5: Create per-account sheets ---
def auto_adjust_column_width(ws):
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[column].width = max_length + 2

unique_accounts = df["Account"].unique()
for account in unique_accounts:
    account_df = df[df["Account"] == account].copy()
    ws = wb.create_sheet(title=str(account))

    headers = [
        "Company", "Account", "Document_Date", "Document_Type", "Text",
        "Document_currency", "Amount_in_doc_curr", 
        "Local_Currency", "Amount_in_local_currency", "Doc_Ageing"
    ]
    for col, header in enumerate(headers, start=1):
        ws.cell(row=1, column=col).value = header

    for idx, row in enumerate(account_df.itertuples(index=False), start=2):
        row_dict = row._asdict()
        ws.cell(row=idx, column=1).value = row_dict.get("Comapany")
        ws.cell(row=idx, column=2).value = row_dict.get("Account")
        doc_date = row_dict.get("Document_Date")
        ws.cell(row=idx, column=3).value = pd.to_datetime(doc_date).strftime("%d.%m.%Y") if pd.notna(doc_date) else None
        ws.cell(row=idx, column=4).value = row_dict.get("Document_Type")
        ws.cell(row=idx, column=5).value = row_dict.get("Text")
        ws.cell(row=idx, column=6).value = row_dict.get("Document_currency")
        ws.cell(row=idx, column=7).value = row_dict.get("Amount_in_doc_curr")
        ws.cell(row=idx, column=8).value = row_dict.get("Local_Currency")
        ws.cell(row=idx, column=9).value = row_dict.get("Amount_in_local_currency")
        ws.cell(row=idx, column=10).value = row_dict.get("Doc_Ageing")

    # Total row
    total_row = len(account_df) + 2
    ws.cell(row=total_row, column=7).value = account_df["Amount_in_doc_curr"].sum()
    ws.cell(row=total_row, column=9).value = account_df["Amount_in_local_currency"].sum()

    # Adjust column widths after writing
    auto_adjust_column_width(ws)

# Adjust summary sheet columns after writing
auto_adjust_column_width(summary_ws)

# --- Step 6: Save output ---
out_path = os.path.join(os.path.dirname(input_file), "Final Report.xlsx")
wb.save(out_path)
print(f"✅ Final Report.xlsx generated successfully at: {out_path}")
