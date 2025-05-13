
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Avolon Data Dashboard", layout="wide")
st.title("📊 Avolon Monthly Performance Dashboard")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)

    # Load required sheets
    expenses_df = xls.parse('Expenses', skiprows=3)
    scrap_df = xls.parse('Scrap Proceeds')
    wire_df = xls.parse('Wire History', skiprows=2)

    # --- Monthly Expenses ---
    expenses_df['Date'] = pd.to_datetime(expenses_df.iloc[:, 0], errors='coerce')
    expenses_df['Amount'] = pd.to_numeric(expenses_df['Unnamed: 3'], errors='coerce')
    monthly_expenses = expenses_df.dropna(subset=['Date', 'Amount']).groupby(expenses_df['Date'].dt.to_period('M')).sum()['Amount']

    # --- Wire (Revenue) History ---
    wire_df.columns = ['Period', 'Date', 'Amount']
    wire_df['Date'] = pd.to_datetime(wire_df['Date'], errors='coerce')
    wire_df['Amount'] = pd.to_numeric(wire_df['Amount'], errors='coerce')
    monthly_revenue = wire_df.dropna(subset=['Date', 'Amount']).groupby(wire_df['Date'].dt.to_period('M')).sum()['Amount']

    # --- Scrap Analysis for Fast-Selling Parts & Margin ---
    scrap_df['QTY'] = pd.to_numeric(scrap_df['QTY'], errors='coerce')
    scrap_df['PROCEEDS'] = pd.to_numeric(scrap_df['PROCEEDS'], errors='coerce')
    scrap_df['COST'] = pd.to_numeric(scrap_df['COST'], errors='coerce')
    scrap_df['Margin'] = scrap_df['PROCEEDS'] - scrap_df['COST']

    top_parts = scrap_df.groupby('PART NUMBER').agg({
        'QTY': 'sum',
        'PROCEEDS': 'sum',
        'COST': 'sum',
        'Margin': 'sum'
    }).sort_values('QTY', ascending=False).head(10)

    avg_margin = scrap_df['Margin'].mean()

    # --- Dashboard Layout ---
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📈 Monthly Revenue vs Expenses")
        perf_df = pd.DataFrame({
            'Revenue': monthly_revenue,
            'Expenses': monthly_expenses
        }).fillna(0)
        st.line_chart(perf_df)

    with col2:
        st.subheader("💰 Average Margin on Scrap Sales")
        st.metric("Average Margin", f"${avg_margin:,.2f}")

        st.subheader("🚀 Top 10 Fastest Selling Parts")
        st.dataframe(top_parts.reset_index())

else:
    st.info("Please upload an Excel file to begin analysis.")
