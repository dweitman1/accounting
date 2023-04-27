import pandas as pd

output = open("output.xlsx", "rb")

a = pd.read_excel(output, sheet_name="Sheet1")
b = pd.read_excel(output, sheet_name="Sheet3")

print(a.groupby(["Project"]).sum(numeric_only=True))
print(b.groupby(["Company"]).sum(numeric_only=True))
print("---------\n", b.sum(numeric_only=True))

with pd.ExcelWriter(
    "output.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    
    pd.DataFrame(b.groupby(["Company"]).sum(numeric_only=True)).to_excel(writer, sheet_name="CompanySummary")
    pd.DataFrame(a.groupby(["Project"]).sum(numeric_only=True)).to_excel(writer, sheet_name="JobSummary")
    pd.DataFrame(b.sum(numeric_only=True)).to_excel(writer, sheet_name="Totals")