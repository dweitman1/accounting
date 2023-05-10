import pandas as pd

output = open("output.xlsx", "rb")

jobSummary = pd.read_excel(output, sheet_name="JobSummary")
a = jobSummary.groupby(["Fund", "Facility", "Activity"]).sum(numeric_only=True)
print(a)


with pd.ExcelWriter(
    "output.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer: 
    a.to_excel(writer, sheet_name="Sheet1")
    #totals.to_excel(writer, sheet_name="Totals", header=False)