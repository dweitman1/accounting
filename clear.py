import pandas as pd

with pd.ExcelWriter(
    "output.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    
    pd.DataFrame(data={}, columns=["Name", "Company", "Week", "Date", "Project", "Open", "Reg", "OT", "Miles", "Total"]).to_excel(writer, sheet_name="Sheet1", index=False)
    pd.DataFrame(data={}, columns=["Name", "Company", "Week", "Date", "Reg", "OT", "Miles", "Total"]).to_excel(writer, sheet_name="Sheet2", index=False)
    pd.DataFrame(data={}, columns=["Name", "Company", "Month", "Reg", "OT", "Miles", "Total"]).to_excel(writer, sheet_name="Sheet3", index=False)