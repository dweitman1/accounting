import pandas as pd
import time
import random

while True:
    
    print(random.randint(1000000000000, 9999999999999))
    time.sleep(5)


with pd.ExcelWriter(
    "output.xlsx",
    mode="a",
    engine="openpyxl",
    if_sheet_exists="replace",
) as writer:
    
    pd.DataFrame(data={}, columns=["Name", "Company", "Rate", "Multiplier", "Week", "Date", "Job", "Open", "Hours", "Miles", "Total"]).to_excel(writer, sheet_name="Entries", index=False)