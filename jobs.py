import pandas as pd

accounting = open("accounting.xlsx", "rb")

jobs = pd.read_excel(accounting, sheet_name="Job List")

def test(s):
    return ['background-color: green' if s_ else 'background-color: red' for s_ in s]

out = jobs.style.apply(test, subset=["Open"], axis=1)

with pd.ExcelWriter(
    "accounting.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    
    out.to_excel(writer, sheet_name="Sheet1", index=False)

