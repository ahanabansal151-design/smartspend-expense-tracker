import pandas as pd
import plotly.express as px
import streamlit as st
from expense_core import DEFAULT_BUDGETS, process_expenses, make_report

st.set_page_config(page_title="SmartSpend | Student Expense Tracker", page_icon="💸", layout="wide")
st.title("💸 SmartSpend — Student Expense Tracker")
st.caption("Upload an Excel workbook → inspect your spending and budget alerts → download one complete Excel report.")
with st.sidebar:
    st.header("Monthly budgets (₹)")
    st.caption("Change these values to match your real budget. Demo defaults are illustrative.")
    budgets={name:float(st.number_input(name,min_value=1.0,value=float(value),step=500.0)) for name,value in DEFAULT_BUDGETS.items()}
file=st.file_uploader("Upload your Excel expense sheet (.xlsx)",type=["xlsx"])
st.info("Expected columns: Date, Category, Amount. Optional: Description and Payment Mode. No artificial transactions are added.")
if file is not None:
    try:
        result=process_expenses(file,budgets)
        excel=make_report(result)
    except Exception as exc:
        st.error(str(exc))
        st.stop()
    df=result["clean"]
    top=result["by_category"].iloc[0]["Category"]
    a,b,c,d=st.columns(4)
    a.metric("Total spent",f"₹{df['Amount'].sum():,.0f}")
    b.metric("Transactions",len(df))
    c.metric("Top category",top)
    d.metric("Over-budget alerts",int((result["monthly"]["Status"]=="OVER BUDGET").sum()))
    st.download_button("⬇️ Download complete Excel report (all sheets + charts)",data=excel,file_name="Student_Expense_Analysis_Report.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",type="primary")
    st.subheader("Charts")
    left,right=st.columns(2)
    with left:
        st.plotly_chart(px.bar(result["by_category"],x="Category",y="Amount",title="Spending by Category"),use_container_width=True)
        st.plotly_chart(px.pie(result["by_category"],names="Category",values="Amount",title="Category Spending Share"),use_container_width=True)
    with right:
        st.plotly_chart(px.line(result["by_month"],x="Month",y="Amount",markers=True,title="Monthly Spending Trend"),use_container_width=True)
        heat=result["heatmap"]
        st.plotly_chart(px.imshow(heat,text_auto=".0f",aspect="auto",title="Monthly Category Heatmap"),use_container_width=True)
    st.subheader("Monthly Budget Alerts")
    st.dataframe(result["monthly"],use_container_width=True,hide_index=True)
    st.subheader("Insights")
    for line in result["insights"]: st.write("•",line)
    with st.expander("Data cleaning details and cleaned transactions"):
        st.dataframe(result["quality"],hide_index=True)
        st.dataframe(result["clean"],hide_index=True)
else:
    st.write("Choose one of the included example Excel files to try the app.")
