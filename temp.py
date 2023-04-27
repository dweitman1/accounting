import pandas as pd

output = open("output.xlsx", "rb")

x = pd.read_excel(output, sheet_name="Sheet1")

print(x.groupby(["Project"]).sum())