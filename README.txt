SMARTSPEND - STUDENT EXPENSE ANALYZER

Local Windows use:
1. Install Python from https://www.python.org/downloads/ and check Add Python to PATH.
2. Unzip the package, open its folder, double-click START_WINDOWS.bat.
3. Installation needs internet the first time. A browser tab will open to the app.
4. Choose an example workbook from samples/ or your own Excel .xlsx file.
5. View the dashboard and click Download complete Excel report.

The ordinary user does not use Google Colab. On your own computer the Streamlit
server has to run (START_WINDOWS.bat); you can later deploy this folder to
Streamlit Community Cloud using a GitHub repository so visitors only open a URL.

Upload columns: Date, Category, Amount (Description and Payment Mode optional).
A positive Amount is an expense. Missing, invalid, duplicate or nonpositive
records are omitted from calculations. The report lists how many were excluded.
Budget amounts are manually adjustable on the sidebar. Default demo budgets
are NOT financial advice. There is no bank connection or automatic mobile alert.

Main files: app.py (user interface), expense_core.py (analysis/export),
requirements.txt (libraries), samples/ (synthetic demo spreadsheets).
