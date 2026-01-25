# Data Terminal UI

This small UI helps you load an Excel/CSV file, preview rows, view summary statistics, and quickly plot columns to help make data decisions.

Quick start:

1. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

2. Run the UI:

```bash
python ui_terminal.py
```

Features:
- Load `.xlsx`, `.xls`, `.csv` files
- Preview first 100 rows
- Show `describe()` summary
- Select a column and plot (line plot for numeric, value counts fallback for categorical)

If you want: I can wire this to your PDF-to-Excel flow so users can load the converted output directly.