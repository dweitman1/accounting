import pandas as pd
import numpy as np

mileageRate = 0.655

output = open("output.xlsx", "rb")

a = pd.read_excel(output, sheet_name="Entries")
a["Totals"] = (a["Hours"] * a["Rate"] * a["Multiplier"]) + (mileageRate * a["Miles"])

weeklySummary = a.groupby(["Company", "Name", "Week"])[["Hours", "Miles"]].sum(numeric_only=True)
regHours = np.where(weeklySummary["Hours"] > 40, 40, weeklySummary["Hours"])
otHours = np.where(weeklySummary["Hours"] > 40, weeklySummary["Hours"] - 40, 0)
weeklySummary.insert(1, "Reg Hours", regHours)
weeklySummary.insert(2, "OT Hours", otHours)



monthlySummary = a.groupby(["Company", "Name", "Rate", "Multiplier"])[["Hours", "Miles"]].sum(numeric_only=True)
monthlySummary = monthlySummary.reset_index(level=["Rate", "Multiplier"])
monthlySummary.insert(3, "Reg Hours", weeklySummary.groupby(["Company", "Name"])["Reg Hours"].sum(numeric_only=True).values)
monthlySummary.insert(4, "OT Hours", weeklySummary.groupby(["Company", "Name"])["OT Hours"].sum(numeric_only=True).values)
monthlySummary["Reg Hours Total"] = monthlySummary["Reg Hours"] * monthlySummary["Rate"]
monthlySummary["OT Hours Total"] = monthlySummary["OT Hours"] * monthlySummary["Rate"] * 1.5
monthlySummary["Multiplier Total"] = (monthlySummary["Reg Hours Total"] + monthlySummary["OT Hours Total"]) * monthlySummary["Multiplier"]
monthlySummary["Miles Total"] = monthlySummary["Miles"] * mileageRate
monthlySummary["Total"] = monthlySummary["Multiplier Total"] + monthlySummary["Miles Total"]



companySummary = monthlySummary.drop(columns=["Rate", "Multiplier"]).groupby(["Company"]).sum(numeric_only=True)

totals = companySummary.sum(numeric_only=True)

jobSummary = a.drop(columns=["Rate", "Multiplier"]).groupby(["Job"]).sum(numeric_only=True)

with pd.ExcelWriter(
    "output.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    
    a.to_excel(writer, sheet_name="Entries", index=False)
    weeklySummary.to_excel(writer, sheet_name="Weekly")
    monthlySummary.to_excel(writer, sheet_name="Monthly")
    companySummary.to_excel(writer, sheet_name="CompanySummary")
    totals.to_excel(writer, sheet_name="Totals", header=False)
    jobSummary.to_excel(writer, sheet_name="JobSummary")