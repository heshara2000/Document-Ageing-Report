import pandas as pd
from openpyxl import Workbook
from datetime import datetime

# Step 1: Load export.xls (must be in the same folder as this script)
#df = pd.read_excel("export.xls", engine="xlrd")
df = pd.read_excel(r"E:\dil_copies\Document-Ageing-Report-\data\export.xls", engine="xlrd")

# Step 2: Create new workbook
wb = Workbook()
summary = wb.active
summary.title = "Summary"

# Step 3: Add report title and today's date
# summary["A1"] = "Document Ageing Report"
# summary["B2"] = datetime.today().strftime("%Y-%m-%d")

summary["A1"] = f"Document Ageing Report as at {datetime.today().strftime('%d.%m.%Y')}"
summary["B2"] = datetime.today().strftime("%Y-%m-%d")


# Step 4: Loop through accounts and create sheets
for account, acc_df in df.groupby("Account"):
    ws = wb.create_sheet(title=str(account))

    # Write headers
    for col_idx, col_name in enumerate(acc_df.columns, 1):
        ws.cell(row=1, column=col_idx, value=col_name)

    # Write rows
    for r_idx, row in enumerate(acc_df.itertuples(index=False), 2):
        for c_idx, value in enumerate(row, 1):
            ws.cell(row=r_idx, column=c_idx, value=value)

    # Step 5: Add Column J with date difference formula
    # Replace 'H' with the column letter of your date column
    for r in range(2, len(acc_df) + 2):
        ws.cell(row=r, column=10, value=f"=TODAY()-H{r}")

    # Update summary
    summary.append([account, len(acc_df)])

# Step 6: Save file
final_name = f"Final_Report_{datetime.today().strftime('%Y%m%d')}.xlsx"
wb.save(final_name)
print(f"✅ Report saved as {final_name}")
