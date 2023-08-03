import pandas as pd

with pd.ExcelWriter(
    "summary.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    pd.DataFrame(data={}, columns=["Name", "Company", "Rate", "Multiplier", "Week", "Fund", "Activity", "Facility", "Job", "Open", "Hours", "Miles", "Total"]).to_excel(writer, sheet_name="Entries", index=False)