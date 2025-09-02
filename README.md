# Document-Ageing-Report-

⚙️ Features

Reads .xls Excel input files

Normalizes column names automatically

Calculates document ageing using today's date

Generates a Summary sheet and per-account sheets

Applies table headers, borders, and formatting for better readability

Handles multiple currencies and sums totals


Document-Ageing-Report/
│
├─ data/
│   └─ export.xls           # Input Excel file with sample data
│
├─ automation.py            # Main script for generating report
├─ Final Report.xlsx        # Generated report
└─ README.md               # This documentation


🚀 Installation

1.Clone this repository:
git clone <your-repo-url>
cd Document-Ageing-Report

2.Create a virtual environment
python -m venv env
env\Scripts\activate      # Windows
source env/bin/activate   # Linux/macOS

3.Install dependencies:
pip install -r requirements.txt

4.Run the main automation script:
python automation.py

