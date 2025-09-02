import pandas as pd
from datetime import datetime

# Sample data
data = [
    {
        "Comapany": "ABC Ltd",
        "Account": "1001",
        "Document Date": datetime(2024, 8, 1),
        "Document Type": "Invoice",
        "Text": "Office Supplies",
        "Document currency": "USD",
        "Amount in doc. curr.": 1200,
        "Local Currency": "LKR",
        "Amount in local currency": 440000,
        "Year/month": "2024/08",
    }
    # {
    #     "Comapany": "ABC Ltd",
    #     "Account": "1002",
    #     "Document Date": datetime(2024, 7, 15),
    #     "Document Type": "Invoice",
    #     "Text": "Equipment",
    #     "Document currency": "EUR",
    #     "Amount in doc. curr.": 800,
    #     "Local Currency": "LKR",
    #     "Amount in local currency": 300000,
    #     "Year/month": "2024/07",
    # },
    # {
    #     "Comapany": "XYZ Pvt",
    #     "Account": "1001",
    #     "Document Date": datetime(2024, 6, 10),
    #     "Document Type": "Credit Note",
    #     "Text": "Return",
    #     "Document currency": "USD",
    #     "Amount in doc. curr.": -200,
    #     "Local Currency": "LKR",
    #     "Amount in local currency": -74000,
    #     "Year/month": "2024/06",
    # },
]

# Convert to DataFrame
df = pd.DataFrame(data)

# Save to Excel (xlsx format)
output_path = r"E:\dil_copies\Document-Ageing-Report-\data\export.xlsx"
df.to_excel(output_path, index=False, engine="openpyxl")

print(f"✅ Sample data written to {output_path}")

