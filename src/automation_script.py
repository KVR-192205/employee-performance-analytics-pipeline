import pandas as pd
import numpy as np

# Load dataset
file_path = "data/employee_data.xlsx"
df = pd.read_excel(file_path)

# Clean column names
df.columns = [col.strip().replace(" ", "_") for col in df.columns]

# -----------------------------
# FEATURE ENGINEERING
# -----------------------------

# Performance Percentage
df["Performance_Percentage"] = (df["Performance_Score"] / 10) * 100

# Performance Status
def categorize(score):
    if score >= 9:
        return "Top Performer"
    elif score >= 7:
        return "High Performer"
    elif score >= 4:
        return "Average"
    else:
        return "Low Performer"

df["Performance_Status"] = df["Performance_Score"].apply(categorize)

# Efficiency Score
df["Efficiency_Score"] = df["Performance_Percentage"] / (df["Experience_Years"] + 1)

# Ranking
df["Rank"] = df["Performance_Percentage"].rank(ascending=False)

# Risk Flag
df["Risk_Flag"] = df["Performance_Status"].apply(
    lambda x: "At Risk" if x == "Low Performer" else "Stable"
)

# -----------------------------
# OUTLIER DETECTION
# -----------------------------

mean_perf = df["Performance_Percentage"].mean()
std_perf = df["Performance_Percentage"].std()

df["Outlier_Flag"] = df["Performance_Percentage"].apply(
    lambda x: "Outlier" if abs(x - mean_perf) > 2 * std_perf else "Normal"
)

# -----------------------------
# KPI METRICS
# -----------------------------

total_employees = len(df)
avg_performance = df["Performance_Percentage"].mean()
high_performers = (df["Performance_Status"] == "High Performer").sum()
low_performers = (df["Performance_Status"] == "Low Performer").sum()

kpi_df = pd.DataFrame({
    "Metric": ["Total Employees", "Average Performance", "High Performers", "Low Performers"],
    "Value": [total_employees, avg_performance, high_performers, low_performers]
})

# -----------------------------
# DEPARTMENT ANALYSIS
# -----------------------------

dept_summary = df.groupby("Department")["Performance_Percentage"].mean().reset_index()
dept_summary = dept_summary.sort_values(by="Performance_Percentage", ascending=False)

# -----------------------------
# TOP / LOW PERFORMERS
# -----------------------------

top_performers = df.sort_values(by="Performance_Percentage", ascending=False).head(10)
low_performers_df = df[df["Performance_Status"] == "Low Performer"]

# -----------------------------
# INSIGHT GENERATION
# -----------------------------

best_dept = dept_summary.iloc[0]["Department"]
best_score = dept_summary.iloc[0]["Performance_Percentage"]

insights = [
    f"Total Employees: {total_employees}",
    f"Average Performance: {round(avg_performance,2)}%",
    f"Best Performing Department: {best_dept} ({round(best_score,2)}%)",
    f"High Performers Count: {high_performers}",
    f"Low Performers Count: {low_performers}",
]

insights_df = pd.DataFrame({"Insights": insights})

# -----------------------------
# SUMMARY ROW
# -----------------------------

summary_row = ["Average"] + [""]*(len(df.columns)-2) + [avg_performance]
df.loc[len(df)] = summary_row

# -----------------------------
# SAVE TO EXCEL
# -----------------------------

with pd.ExcelWriter("output/processed_employee_data.xlsx", engine="openpyxl", mode="w") as writer:
    df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
    dept_summary.to_excel(writer, sheet_name="Dept_Analysis", index=False)
    kpi_df.to_excel(writer, sheet_name="KPI_Summary", index=False)
    top_performers.to_excel(writer, sheet_name="Top_Performers", index=False)
    low_performers_df.to_excel(writer, sheet_name="Low_Performers", index=False)
    insights_df.to_excel(writer, sheet_name="Insights", index=False)

print("Full advanced analytics pipeline completed successfully!")