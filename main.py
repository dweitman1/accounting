import sys
import pandas as pd
import datetime as dt

accounting = open("accounting.xlsx", "rb")
jobs = pd.read_excel(accounting, sheet_name="Open Jobs")
ic = pd.read_excel(accounting, sheet_name="WSP-IC2018")

#---Find contractor---#
contractor = pd.DataFrame()
while contractor.empty:
    contractorName = input("Enter contractor name: ")
    contractor = ic[ic["Name"] == contractorName]

    if contractor.empty:
        print("Contractor not found...")

numWeeks = int(sys.argv[1])
weekEnding = (dt.datetime(int(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4])) + dt.timedelta(days=6))
currentWeek = 1

while numWeeks > 0:
    currentEntry = 1
    weeklyEntries = int(input(f"Number of entries for week {currentWeek}: "))

    while weeklyEntries > 0:    
        job = pd.DataFrame()
        isOpen = True
        
        jobNumber = input("Job Number: ")
        job = jobs[jobs["JOB NUMBER"] == jobNumber]

        if job.empty:
            print(f"Job {jobNumber} not found")
            isOpen = False
            project = pd.DataFrame([{"JOB NUMBER": jobNumber}])

        hours = float(input("Hours: "))
        miles = float(input("Miles: "))

        with pd.ExcelWriter("output.xlsx", mode="a", if_sheet_exists="overlay") as writer:
            entry = {"Name": contractor["Name"].iloc[0],
            "Company": contractor["Company"].iloc[0],
            "Rate": contractor["Rate"].iloc[0],
            "Multiplier": contractor["Multiplier"].iloc[0],
            "Week": "Week " + str(currentWeek),
            "Date": weekEnding.strftime("%b %d %Y"),
            "Job": job["JOB NUMBER"].iloc[0],
            "Open": isOpen,
            "Hours": hours,
            "Miles": miles,
            "Total": None}
            pd.DataFrame([entry]).to_excel(writer, sheet_name="Entries", startrow=writer.sheets["Entries"].max_row, header=False, index=False)
         
        weeklyEntries -= 1
        currentEntry += 1
    numWeeks -=1
    currentWeek += 1
    weekEnding = (weekEnding + dt.timedelta(days=7))
