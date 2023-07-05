import sys
import pandas as pd
import datetime as dt

accounting = open("accounting.xlsx", "rb")
ic = pd.read_excel(accounting, sheet_name="WSP-IC2020")

currentWeek = 1
numWeeks = int(sys.argv[1])
weekEnding = (dt.datetime(int(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4])) + dt.timedelta(days=6))

#---Get Contractor Data---#
contractor = pd.DataFrame()
while contractor.empty:
    contractorName = input("Enter contractor name: ")
    contractor = ic[ic["Name"] == contractorName]

    if contractor.empty:
        print("Contractor not found...")

#---Timesheet Entries---#
repeatFlag = False
while numWeeks > 0:
    if currentWeek > 1:
        repeatFlag = input("Repeat entries? ")

    if repeatFlag:
        savedEntries["Week"] = "Week " + str(currentWeek)
        savedEntries["Date"] = weekEnding.strftime("%b %d %Y")
        for i, row in savedEntries.iterrows():
            x = row["Job"]
            savedEntries.at[i, "Hours"] = input(f"{x} Hours: ")
            savedEntries.at[i, "Miles"] = input(f"{x} Miles: ")
        
        with pd.ExcelWriter("output.xlsx", mode="a", if_sheet_exists="overlay") as writer:
            savedEntries.to_excel(writer, sheet_name="Entries", startrow=writer.sheets["Entries"].max_row, header=False, index=False)

    else:
        while True:
            try:
                weeklyEntries = int(input(f"Number of entries for week {currentWeek}: "))
                savedEntries = pd.DataFrame()
            except:
                print("Invalid input!")
                continue
            else:
                break

        while weeklyEntries > 0:    
            jobs = pd.read_excel(accounting, sheet_name="Job List")
            jobNumber = input("Job Number: ")
            job = pd.DataFrame()
            job = jobs[jobs["Job"] == jobNumber]
            isOpen = True

            if job.empty:
                print(f"Job {jobNumber} not found")
                openJobs = pd.read_excel(accounting, sheet_name="Open Jobs")
                openJob = openJobs[openJobs["JOB NUMBER"] == jobNumber]
                if openJob.empty:
                    isOpen = False

                job = pd.DataFrame([{
                "Fund": input("Fund: "),
                "Activity": input("Activity: "),
                "Facility": input("Facility: "),
                "Job": jobNumber,
                "Open": isOpen}])

                with pd.ExcelWriter("accounting.xlsx", mode="a", if_sheet_exists="overlay") as writer:
                    pd.DataFrame(job).to_excel(writer, sheet_name="Job List", startrow=writer.sheets["Job List"].max_row, header=False, index=False)
            
            hours = float(input("Hours: "))
            miles = float(input("Miles: "))

            entry = pd.DataFrame([{
                "Name": contractor["Name"].iloc[0],
                "Company": contractor["Company"].iloc[0],
                "Rate": contractor["Rate"].iloc[0],
                "Week": "Week " + str(currentWeek),
                "Date": weekEnding.strftime("%b %d %Y"),
                "Fund": job["Fund"].iloc[0],
                "Activity": job["Activity"].iloc[0],
                "Facility": job["Facility"].iloc[0],
                "Job": job["Job"].iloc[0],
                "Open": isOpen,
                "Hours": hours,
                "Miles": miles,
                "Total": None}])
            
            savedEntries= pd.concat([savedEntries, entry], ignore_index=True)
            
            #---Write timesheet entry to Excel---#
            with pd.ExcelWriter("output.xlsx", mode="a", if_sheet_exists="overlay") as writer:
                entry.to_excel(writer, sheet_name="Entries", startrow=writer.sheets["Entries"].max_row, header=False, index=False)
            weeklyEntries -= 1

    numWeeks -=1
    currentWeek += 1
    weekEnding = (weekEnding + dt.timedelta(days=7))
accounting.close()