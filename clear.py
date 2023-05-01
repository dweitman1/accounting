import pandas as pd

with pd.ExcelWriter(
    "output.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    
    pd.DataFrame(data={}, columns=["Name", "Company", "Rate", "Week", "Date", "Project", "Open", "Hours", "Miles", "Total"]).to_excel(writer, sheet_name="Entries", index=False)