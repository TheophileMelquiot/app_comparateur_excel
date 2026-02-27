📊 Excel Comparison Tool – Smart Key-Based Comparison with Highlighted Differences
🚀 Overview

This project is an advanced Excel comparison tool built with Python and Gradio.

It allows you to:

📂 Upload two Excel files

📑 Select specific sheets

🔑 Choose a key column for comparison

🔍 Detect duplicate occurrences safely

🎯 Compare values with numeric tolerance

🎨 Generate a new Excel file highlighting differences in red

📊 Identify the most unique columns (Top 5 unique IDs detection)

This tool is especially useful for:

Data reconciliation

Financial data validation

Regulatory reporting control

Data quality auditing

Operational risk checks

🖥️ Interface Preview

The application is built with Gradio Blocks UI, providing:

Dynamic sheet detection

Column auto-loading

Header row configuration

Unique ID analysis

Automatic Excel export with formatting

⚙️ Features
🔎 1. Unique ID Detection

Computes number of unique values per column

Displays Top 5 most discriminant columns

Shows uniqueness percentage

🔁 2. Smart Comparison Logic

The comparison engine:

Normalizes column names (case, accents, spaces)

Handles duplicate keys using occurrence indexing

Cleans strings (whitespace, line breaks, casing)

Handles numeric rounding tolerance

Compares safely using a custom safe_compare() function

🎨 3. Highlighted Excel Output

Differences are highlighted in red

Column widths auto-adjusted

Clean side-by-side structure:

merge_key | col1_file1 | col1_file2 | col2_file1 | col2_file2 | ...
🏗️ Project Structure
📁 excel-comparison-tool
│
├── comparaison_final_tool.py
├── README.md
└── requirements.txt
🛠️ Installation
1️⃣ Clone the repository
git clone https://github.com/your-username/excel-comparison-tool.git
cd excel-comparison-tool
2️⃣ Create virtual environment (recommended)
python -m venv venv
source venv/bin/activate   # Mac/Linux
venv\Scripts\activate      # Windows
3️⃣ Install dependencies
pip install -r requirements.txt
▶️ Run the Application
python comparaison_final_tool.py

The app will automatically open in your browser.

📦 Dependencies

pandas

numpy

gradio

openpyxl

unidecode

(See requirements.txt below)

🧠 Technical Highlights

Custom merge strategy using:

merge_key

occurrence index per duplicate key

Safe numeric comparison with rounding tolerance

Automatic column normalization using unidecode

Excel generation via openpyxl

Temporary file handling with tempfile

📈 Use Cases

✔ Banking reconciliation
✔ Regulatory reporting checks
✔ Data migration validation
✔ ERP comparison
✔ Internal audit controls

👨‍💻 Author

Théophile Melquiot
Master IASD – Artificial Intelligence & Big Data
Data Analysis & Data Management

📜 License

MIT License (or specify your preferred license)
