import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl


# Load Data
@st.cache_data
def load_data():
    file_path = "Task Assignment Data.xlsx"
    xls = pd.ExcelFile(file_path)

    # Load Sheets
    df_data = pd.read_excel(xls, sheet_name="Data")
    df_charge_offs = pd.read_excel(xls, sheet_name="Charge-Offs")

    # Convert Date Columns
    df_data["EntryDate"] = pd.to_datetime(df_data["EntryDate"], errors='coerce')
    df_charge_offs["Month"] = pd.to_datetime(df_charge_offs["Month"], errors='coerce')

    return df_data, df_charge_offs

df_data, df_charge_offs = load_data()

# Streamlit Dashboard Layout
st.title("Loan Portfolio & Charge-Off Dashboard")

# KPI Metrics
st.subheader("📊 Key Loan Metrics")
total_loans = df_data["LoanNumber"].count()
total_approved_amount = df_data["AmountApproved"].sum()
avg_credit_score = df_data["CreditScore"].mean()

st.metric(label="Total Loans", value=f"{total_loans:,}")
st.metric(label="Total Approved Amount ($)", value=f"${total_approved_amount:,.2f}")
st.metric(label="Average Credit Score", value=f"{avg_credit_score:.1f}")

# Loan Volume Over Time
st.subheader("📈 Loan Volume Over Time")
loan_trend = df_data.groupby(df_data["EntryDate"].dt.to_period("M"))["LoanNumber"].count()

fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(loan_trend.index.astype(str), loan_trend, marker="o", color="blue")
ax.set_title("Loan Volume Over Time")
ax.set_xlabel("Month")
ax.set_ylabel("Number of Loans")
ax.grid(True)
st.pyplot(fig)

# Dealer Performance
st.subheader("🏢 Dealer Performance")
dealer_performance = df_data.groupby("ClinicName").agg(
    Total_Loan_Count=("LoanNumber", "count"),
    Total_Approved_Amount=("AmountApproved", "sum")
).sort_values(by="Total_Loan_Count", ascending=False)

st.write("### Top 10 Dealers by Loan Volume")
st.dataframe(dealer_performance.head(10))

# Charge-Off Trends
st.subheader("⚠️ Charge-Off Rate Trends")
fig, ax = plt.subplots(figsize=(10, 5))
ax.plot(df_charge_offs["Month"], df_charge_offs["Unsecured Gross Charge-Offs (%)"], marker="o", label="Unsecured", color="red")
ax.plot(df_charge_offs["Month"], df_charge_offs["Secured Gross \nCharge-Offs (%)"], marker="s", label="Secured", color="blue")
ax.set_title("Charge-Off Rate Trends")
ax.set_xlabel("Month")
ax.set_ylabel("Charge-Off Rate (%)")
ax.legend()
ax.grid(True)
st.pyplot(fig)

st.write("### Total Charge-Off Amounts")
charge_off_summary = df_charge_offs[["Unsecured Gross Charge-Offs ($)", "Secured Gross \nCharge-Offs ($)"]].sum()
st.dataframe(charge_off_summary.to_frame())

st.success("✅ Dashboard Loaded Successfully!")
