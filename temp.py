import pandas as pd

output = open("output.xlsx", "rb")

x = pd.read_excel(output, sheet_name="Entries")

print(x.groupby(["Project"]).sum())