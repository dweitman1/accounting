import pandas as pd

accounting = open("accounting.xlsx", "rb")

openJobs = pd.read_excel(accounting, sheet_name="Open Jobs")

while True:
    openJob = openJobs[openJobs["JOB NUMBER"] == input("Job Number: ")]

    if openJob.empty:
        print("Job is closed")
    else:
        print("Job is open")


