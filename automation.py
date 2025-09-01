import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from datetime import datetime

# Load the source data
df = pd.read_excel(r"E:\dil_copies\Document-Ageing-Report-\data\export.xls", engine="xlrd")
print(df.columns)

# # Create a new workbook
# wb = Workbook()
# summary = wb.active
# summary.title = "Summary"

# # Title in C2
# summary["C2"] = f"Document Ageing Report as at {datetime.today().strftime('%d.%m.%Y')}"
# summary["C2"].font = Font(bold=True, size=14)
# summary["C2"].alignment = Alignment(horizontal="center")

# # Add table headers in row 4
# headers = ["Company", "Account", "Document currency", "Amount in doc. curr.", "Local Currency", "Amount in local currency"]
# for col_idx, col_name in enumerate(headers, 3):  # start at column C (3rd col)
#     summary.cell(row=4, column=col_idx, value=col_name).font = Font(bold=True)

# # Fill summary rows (row 5 onwards)
# row_num = 5
# for _, row in df.iterrows():
#     summary.cell(row=row_num, column=3, value=row["Company"])
#     summary.cell(row=row_num, column=4, value=row["Account"])
#     summary.cell(row=row_num, column=5, value=row["Document currency"])
#     summary.cell(row=row_num, column=6, value=row["Amount in doc. curr."])
#     summary.cell(row=row_num, column=7, value=row["Local Currency"])
#     summary.cell(row=row_num, column=8, value=row["Amount in local currency"])
#     row_num += 1

# # Save
# final_name = f"Final_Report_{datetime.today().strftime('%Y%m%d')}.xlsx"
# wb.save(final_name)
# print(f"✅ Report saved as {final_name}")
