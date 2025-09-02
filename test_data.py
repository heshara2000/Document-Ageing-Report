import pandas as pd
from datetime import datetime

# Sample data
data = [
    {
        "Comapany": "UN015O",
        "Account": "63010503",
        "Document Date": datetime(2024, 8, 1),
        "Document Type": "FZ",
        "Text": "0601530BCC701 PRCEEDS F PURC",
        "Document currency": "USD",
        "Amount in doc. curr.": "12431.22 USD",
        "Local Currency": "USD",
        "Amount in local currency": "12431",
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

# Save to Excel in old .xls format
output_path = r"E:\dil_copies\Document-Ageing-Report-\data\export.xls"
df.to_excel(output_path, index=False, engine="xlwt")


print(f"✅ Sample data written to {output_path}")

