Excel Automation with Python

A Python automation project for processing Excel spreadsheets, correcting prices, and generating charts automatically. This project demonstrates how to automate repetitive Excel tasks that would normally take hours or weeks to complete manually.

📋 Project Overview

This automation tool processes Excel spreadsheets to:

· Apply bulk price corrections (10% reduction)
· Add calculated columns automatically
· Generate visual charts
· Handle thousands of rows in seconds

🚀 Features

· Bulk Data Processing: Process thousands of spreadsheet rows instantly
· Price Correction: Automatically apply 10% price reductions across all products
· Chart Generation: Create bar charts directly in Excel files
· File Management: Process single files or batch process entire directories
· Professional Code Structure: Clean, reusable functions for production use

🛠 Technologies Used

· Python - Core programming language
· OpenPyXL - Excel file manipulation
· Pandas - Data processing (implied)

📁 Project Structure

```
excel-automation/
├── process_spreadsheets.py  # Main automation script
├── transactions.xlsx        # Input file (example)
├── transactions2.xlsx       # Output file (processed)
└── README.md
```

⚡ Quick Start

Prerequisites

```bash
pip install openpyxl pandas
```

Basic Usage

```python
from process_spreadsheets import process_workbook

# Process a single spreadsheet
process_workbook('transactions.xlsx')

# For batch processing multiple files:
import os
for file in os.listdir('spreadsheets_directory'):
    if file.endswith('.xlsx'):
        process_workbook(file)
```

🔧 How It Works

1. Load Spreadsheet

```python
import openpyxl as excel
workbook = excel.load_workbook('transactions.xlsx')
sheet = workbook['Sheet1']
```

2. Price Correction

· Iterates through all rows (skipping headers)
· Applies 10% reduction to prices in column 3
· Adds corrected prices to new column 4

3. Chart Generation

· Creates bar charts using OpenPyXL
· Positions charts adjacent to data (e.g., cell E2)
· Customizable chart types and styles

4. Save Results

· Overwrites original file or creates new version
· Maintains data integrity and formatting

💡 Key Code Snippets

Main Processing Function

```python
def process_workbook(filename):
    workbook = excel.load_workbook(filename)
    sheet = workbook['Sheet1']
    
    # Correct prices and add new column
    for row in range(2, sheet.max_row + 1):
        price = sheet.cell(row, 3).value
        corrected_price = price * 0.9
        corrected_price_cell = sheet.cell(row, 4)
        corrected_price_cell.value = corrected_price
    
    # Add chart and save
    # ... chart configuration code ...
    workbook.save(filename)
```

📊 Sample Workflow

Input:

Transaction ID Product ID Price
1 A100 5.95
2 A101 6.95

Output:

Transaction ID Product ID Price Corrected Price Chart
1 A100 5.95 5.36 █████
2 A101 6.95 6.26 ██████

🎯 Use Cases

· E-commerce: Bulk price updates during sales
· Finance: Automated financial report generation
· Data Analysis: Rapid data transformation and visualization
· Inventory Management: Mass product price adjustments

📝 Notes

· The project starts with basic scripting and evolves into professional, reusable functions
· Charts are customizable (colors, types, positioning)
· Error handling can be added for production use
· Compatible with CSV and JSON files with minor modifications

🔮 Future Enhancements

· Add support for multiple chart types
· Implement error handling and logging
· Add configuration files for different business rules
· Create web interface for non-technical users
· Add email notification for completed batches

📚 Learning Resources

Based on Python automation concepts from Mosh Hamedani's programming courses. This project demonstrates real-world application of:

· File I/O operations
· Excel manipulation
· Data transformation
· Code refactoring and organization

---

⭐ Star this repo if you found it helpful for your automation needs!