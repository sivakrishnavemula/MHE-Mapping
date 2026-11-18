# Carrier Change Dashboard

This dashboard combines Excel transaction files and displays interactive summaries. You do not need to edit the Python code to update the data.

## 1. Install Python

If Python is not installed, download and install Python 3.10 or newer from [python.org](https://www.python.org/downloads/). During installation on Windows, select **Add Python to PATH**.

## 2. Add your Excel files

Put the `.xlsx` or `.xls` files in the `data` folder next to `app.py`. The dashboard reads all supported files and removes exact duplicate rows that occur across files.

## 3. Install the dashboard requirements

Open Command Prompt in this project folder and run:

```text
pip install -r requirements.txt
```

## 4. Start the dashboard

In the same Command Prompt window, run:

```text
streamlit run app.py
```

Your browser should open the dashboard. If it does not, copy the local address printed in the Command Prompt (usually `http://localhost:8501`) into your browser.

On Windows, you can instead double-click `run_dashboard.bat`. The first run may take a little while as the required Python packages are installed.

## Update the data

Add new Excel files to the `data` folder or replace existing files. The dashboard checks file names and modified times and reloads changed files automatically when the page is refreshed. You can also click **Refresh data** in the sidebar or reload the browser page.

If a workbook is open in Excel or damaged, close or replace it and click **Refresh data**. The dashboard will show a message naming files it could not read. Rows with an empty or unreadable transaction time are kept in the detail view but are not included in date-based charts and totals.

## Reporting details

The reporting day runs from 07:00 until 07:00 the following day. Shift 1 is 07:00–15:00, Shift 2 is 15:00–23:00, and Shift 3 is 23:00–07:00. The date range, source, user first name, and shift filters in the sidebar apply to all dashboard sections. Type part of a source name in the Source filter (for example, `MHT`) to match source names containing that text, regardless of letter case. On the Overview tab, use **Download full Overview as PNG** to save the KPIs, daily trend, and daily summary together. Use the Data tab to search, sort, and download the filtered rows as CSV or Excel.
# MHE-Mapping
