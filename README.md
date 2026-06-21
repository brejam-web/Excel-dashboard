# Excel Insight Dashboard

A Streamlit app that turns any uploaded Excel file into an interactive,
visual dashboard — automatically.

## Features
- Upload `.xlsx` / `.xls` / `.xlsm` files (multi-sheet support)
- Auto-detects numeric, categorical, and date columns
- Descriptive statistics (mean, median, std, variance, skew, quartiles)
- Categorical breakdowns with frequency tables and bar charts
- Interactive visual explorer: histograms, box plots, scatter plots,
  bar charts, line charts, pie charts
- Correlation matrix + ranked list of strongest relationships
- Sidebar filters: categorical multiselect, numeric range sliders, date ranges
- Download the filtered data back out as CSV

## Setup

```bash
pip install -r requirements.txt
```

## Run

```bash
streamlit run app.py
```

This opens the app in your browser at `http://localhost:8501`.

## Notes
- Works with multi-sheet workbooks — pick the sheet from the sidebar.
- The app auto-cleans data lightly: drops fully empty rows/columns,
  trims column names, and tries to detect numeric/date columns that were
  stored as text.
- For very large files (100k+ rows), consider sampling before upload for
  best performance, since all charts run client-side via Plotly.
