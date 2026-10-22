import pandas as pd

FILE="novaretail_dataset.xlsx"

#load all 3 sheets
regional = pd.read_excel(FILE,
                         sheet_name="Regional_Performance")
customer=pd.read_excel(FILE,
                       sheet_name="Customer_Experience")
risk=pd.read_excel(FILE,
                   sheet_name="Operational_Risk")

first_quarter="2025-Q3"
latest_quarter="2026-Q2"
previous_quarter="2026-Q1"

def compare_quarters(df,metric):
    """Compare each region between first and latest quarters"""
    pivot = df.pivot(index="Region",
                     columns="Quarter",
                     values=metric)
    
    result=pd.DataFrame({
        "First quarter":pivot[first_quarter],
        "Latest quarter":pivot[latest_quarter]
    })

    result["Change"]=(
        result["Latest quarter"]-result["First quarter"]
    )

    result["Change_pct"]=(
        result["Change"]/result["First quarter"]*100
    )

    return result.round(2)

#1.Revenue Performance
print("\n===== REVENUE PERFORMANCE =====")

quarterly_revenue=(
    regional.groupby("Quarter")["Revenue_USD_M"].sum()
)

print ("\n Total revenue by quarter:")
print(quarterly_revenue)

latest_revenue = quarterly_revenue[latest_quarter]
previous_revenue = quarterly_revenue[previous_quarter]

qoq_growth = (
    (latest_revenue - previous_revenue)/previous_revenue *100
)

print(f"\nLatest quarter revenue: ${latest_revenue:,.0f}M")
print(f"Quarter-on-quarter growth:{qoq_growth:.2f}%")

print("\nrevenue change by region:")
print(compare_quarters(regional,"Revenue_USD_M"))

#2. Customer experience
print("\n==== CUSTOMER EXPERIENCE ====")

for metric in ["NPS","CSAT_pct"]:
    print(f"\n{metric}change by region:")
    print(compare_quarters(customer,metric))

print("\nAvergae delivery days by region:")
print(compare_quarters(customer,"Avg_Delivery_Days"))

#3. Operational Risk
print("\n==== OPERATIONAL RISK ====")
for metric in[
    "Stockout_Rate_pct",
    "Warehouse_Capacity_Util_pct",
    "Fraud_Incidents"
]:
    print(f"\n{metric} change by region:")
    print(compare_quarters(risk,metric))

    print("\nAnalysis complete:")