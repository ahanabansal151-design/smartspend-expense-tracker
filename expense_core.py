"""Testable processing and one-workbook export for SmartSpend."""
from io import BytesIO
import numpy as np
import pandas as pd

REQUIRED = {"Date", "Category", "Amount"}
DEFAULT_BUDGETS = {"Food": 6500.0, "Travel": 6500.0, "Shopping": 8000.0, "Entertainment": 6000.0}

def process_expenses(file_source, budgets=None):
    budgets = dict(DEFAULT_BUDGETS if budgets is None else budgets)
    try:
        original = pd.read_excel(file_source, sheet_name=0, engine="openpyxl")
    except Exception as exc:
        raise ValueError(f"Could not read the Excel file: {exc}") from exc
    original.columns = original.columns.astype(str).str.strip()
    missing = REQUIRED - set(original.columns)
    if missing:
        raise ValueError("Missing required columns: " + ", ".join(sorted(missing)) + ". Required: Date, Category, Amount.")
    if original.empty:
        raise ValueError("The uploaded Excel sheet is empty.")
    df = original.copy()
    df["Category"] = df["Category"].astype("string").str.strip().str.title()
    df["Category"] = df["Category"].replace("", pd.NA)
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce")
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce", dayfirst=False)
    invalid = df["Amount"].isna() | (df["Amount"] <= 0) | df["Date"].isna() | df["Category"].isna()
    invalid_count = int(invalid.sum())
    df = df.loc[~invalid].drop_duplicates().copy()
    duplicate_count = int((~invalid).sum() - len(df))
    if df.empty:
        raise ValueError("No valid transactions remain after cleaning. Check dates, categories and amounts.")
    df = df.sort_values("Date").reset_index(drop=True)
    df["Month"] = df["Date"].dt.strftime("%Y-%m")
    monthly = df.groupby(["Month", "Category"], as_index=False)["Amount"].sum()
    months = sorted(df["Month"].unique().tolist())
    categories = sorted(set(df["Category"].tolist()) | set(budgets.keys()))
    full_index = pd.MultiIndex.from_product([months, categories], names=["Month", "Category"])
    monthly = monthly.set_index(["Month", "Category"]).reindex(full_index, fill_value=0).reset_index()
    monthly["Budget"] = monthly["Category"].map(budgets)
    monthly["Remaining"] = monthly["Budget"] - monthly["Amount"]
    monthly["Status"] = np.select(
       [monthly["Budget"].isna(), monthly["Amount"] > monthly["Budget"], monthly["Amount"] >= .8 * monthly["Budget"]],
       ["NO BUDGET", "OVER BUDGET", "WARNING"], default="SAFE")
    monthly["Usage %"] = np.where(monthly["Budget"].fillna(0) > 0, 100 * monthly["Amount"] / monthly["Budget"], np.nan)
    by_category = df.groupby("Category", as_index=False)["Amount"].sum().sort_values("Amount", ascending=False)
    by_month = df.groupby("Month", as_index=False)["Amount"].sum().sort_values("Month")
    heatmap = df.pivot_table(index="Month", columns="Category", values="Amount", aggfunc="sum", fill_value=0)
    top_category = str(by_category.iloc[0]["Category"])
    top_month = str(by_month.sort_values("Amount", ascending=False).iloc[0]["Month"])
    over = int((monthly["Status"] == "OVER BUDGET").sum())
    insights = [
      f"Total spend: ₹{df['Amount'].sum():,.0f} across {len(df)} valid transactions.",
      f"Highest-spend category: {top_category} (₹{by_category.iloc[0]['Amount']:,.0f}).",
      f"Highest-spend month: {top_month}.",
      f"Over-budget category-month combinations: {over}.",
      f"Recommendation: review {top_category} spending and set realistic monthly budgets before the next month."
    ]
    quality = pd.DataFrame({"Metric":["Original rows", "Invalid rows removed", "Duplicate rows removed", "Valid rows"], "Count":[len(original), invalid_count, duplicate_count, len(df)]})
    return dict(clean=df, monthly=monthly, by_category=by_category, by_month=by_month, heatmap=heatmap, insights=insights, quality=quality, budgets=budgets, months=months)

def make_report(result):
    """Creates exactly one XLSX with native Excel charts and analysis tabs."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        book = writer.book
        summary = book.add_worksheet("Dashboard")
        writer.sheets["Dashboard"] = summary
        summary.hide_gridlines(2)
        title_fmt = book.add_format({"bold":True,"font_size":18,"font_color":"#17365D"})
        heading = book.add_format({"bold":True,"bg_color":"#DCEAF5","border":0})
        money = book.add_format({"num_format":"₹#,##0.00"})
        summary.write("B2", "SMARTSPEND — STUDENT EXPENSE REPORT", title_fmt)
        summary.write("B4", "Total Spending", heading)
        summary.write_number("C4", float(result["clean"]["Amount"].sum()), money)
        summary.write("B5", "Transactions", heading)
        summary.write_number("C5", len(result["clean"]))
        summary.write("B6", "Over-Budget Alerts", heading)
        summary.write_number("C6", int((result["monthly"]["Status"] == "OVER BUDGET").sum()))
        summary.set_column("A:A", 3)
        summary.set_column("B:B", 28)
        summary.set_column("C:C", 18)
        summary.set_column("D:P", 13)
        for key, sheet in [("clean", "Cleaned Data"),("monthly", "Budget Alerts"),("by_category", "Category Summary"),("by_month", "Monthly Summary"),("quality", "Data Quality")]:
            result[key].to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes(1,0)
            ws.autofilter(0,0,len(result[key]),len(result[key].columns)-1)
            ws.set_column(0,len(result[key].columns)-1,20)
            ws.set_row(0,22,heading)
        heat = result["heatmap"]
        heat.to_excel(writer,sheet_name="Heatmap")
        hws=writer.sheets["Heatmap"]
        hws.set_column(0,heat.shape[1],18)
        hws.conditional_format(1,1,len(heat),len(heat.columns),{"type":"3_color_scale"})
        hws.freeze_panes(1,1)
        ins=book.add_worksheet("Insights")
        for i, sentence in enumerate(result["insights"]):
            ins.write(i+1,1,sentence)
        ins.set_column("B:B",105)
        cat_len=len(result["by_category"])
        month_len=len(result["by_month"])
        def insert_chart(kind, title, cat_sheet, n, pos, categories_col=0,values_col=1):
            c=book.add_chart({"type":kind})
            c.add_series({"name":title,"categories":[cat_sheet,1,categories_col,n,categories_col],"values":[cat_sheet,1,values_col,n,values_col]})
            c.set_title({"name":title})
            c.set_legend({"none":True})
            c.set_size({"width":580,"height":290})
            summary.insert_chart(pos,c)
        insert_chart("column","Spending by Category","Category Summary",cat_len,"B9")
        insert_chart("line","Monthly Spending Trend","Monthly Summary",month_len,"K9")
        insert_chart("pie","Share of Spending","Category Summary",cat_len,"B26")
        # This chart compares all-month actuals to matched total budget over observed months.
        compare=result["by_category"].copy()
        compare["Total Budget"] = compare["Category"].map(result["budgets"]) * len(result["months"])
        compare.to_excel(writer,sheet_name="Spending vs Budget",index=False)
        cws=writer.sheets["Spending vs Budget"]
        cws.set_column("A:C",22)
        chart=book.add_chart({"type":"column"})
        chart.add_series({"name":"Actual Spending","categories":["Spending vs Budget",1,0,len(compare),0],"values":["Spending vs Budget",1,1,len(compare),1]})
        chart.add_series({"name":"Total Budget","categories":["Spending vs Budget",1,0,len(compare),0],"values":["Spending vs Budget",1,2,len(compare),2]})
        chart.set_title({"name":"Spending vs Budget (Observed Months)"})
        chart.set_size({"width":580,"height":290})
        summary.insert_chart("K26",chart)
    return output.getvalue()
