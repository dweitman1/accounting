import sys
import pandas as pd

def getWeeklyEntries(jobEntries=pd.DataFrame()):
    while True:
        isOpen = True
        jobList = pd.read_excel(accounting, sheet_name="Job List")
        jobNumber = input(f"Job Number: ")
        if not jobNumber:                 
            weeklyEntries = pd.DataFrame(data={}, columns=["Name", "Company", "Rate", "Multiplier", "Week", "Fund", "Activity", "Facility", "Job", "Open", "Hours", "Miles", "Total"])
            weeklyEntries["Fund"] = jobEntries["Fund"]
            weeklyEntries["Activity"] = jobEntries["Activity"]
            weeklyEntries["Facility"] = jobEntries["Facility"]
            weeklyEntries["Job"] = jobEntries["Job"]
            weeklyEntries["Open"] = jobEntries["Open"]
            weeklyEntries["Name"] = contractor.iloc[0]["Name"]
            weeklyEntries["Company"] = contractor.iloc[0]["Company"]
            weeklyEntries["Rate"] = contractor.iloc[0]["Rate"]
            weeklyEntries["Multiplier"] = contractor.iloc[0]["Multiplier"]
            return weeklyEntries
        job = pd.DataFrame()
        job = jobList[jobList["Job"] == jobNumber]

        if job.empty:
            print(f"Job {jobNumber} not found")
            openJobs = pd.read_excel(accounting, sheet_name="Open Jobs")
            openJob = openJobs[openJobs["JOB NUMBER"] == jobNumber]
            if openJob.empty:
                isOpen = False
                print(f"Job {jobNumber} not open")
            else:
                jobCodes = pd.read_excel(accounting, sheet_name="Job Codes")
                print(jobCodes.to_string(index=False))

            fund = input("Fund: ")
            activity = input("Activity: ")
            facility = input("Facility: ")
            if fund and activity and facility:
                job = pd.DataFrame([{
                    "Fund": fund,
                    "Activity": activity,
                    "Facility": facility,
                    "Job": jobNumber,
                    "Open": isOpen}])
                
                with pd.ExcelWriter("accounting.xlsx", mode="a", if_sheet_exists="overlay") as writer:
                    pd.DataFrame(job).to_excel(writer, sheet_name="Job List", startrow=writer.sheets["Job List"].max_row, header=False, index=False)
        
        if not job.empty and job["Job"].iloc[0]=="BLANK":
            acts = pd.DataFrame({"Type": ["General", "Training", "Management"],"Activity": [17522, 28172, 28123]})
            print(acts)
            try:
                ind = int(input("Enter Activity Type: "))
            except:
                ind = 0
            job.iloc[0, job.columns.get_loc("Activity")] = acts["Activity"].iloc[ind]; 

        jobEntries = pd.concat([jobEntries, job], ignore_index=True)
        print(jobEntries[["Job"]])

accounting = open("accounting.xlsx", "rb")
ic = pd.read_excel(accounting, sheet_name=sys.argv[1])
numWeeks = int(sys.argv[2])
currentWeek = 1









#---Get Contractor Data---#
contractor = pd.DataFrame()
while contractor.empty:
    contractorName = input("Enter contractor name: ")
    contractor = ic[ic["Name"] == contractorName]
    contractor = contractor[["Name", "Company", "Rate", "Multiplier"]]
    if contractor.empty:
        print("Contractor not found...")

#---Input Job Entries---#
weeklyEntries = getWeeklyEntries()
contractorEntries = pd.DataFrame()

while currentWeek <= numWeeks:
    if currentWeek > 1:
        print(weeklyEntries[["Job"]])

        choice = True
        while choice:
            choice = input(f"Week {currentWeek} Add (a) | Remove (r): ")

            if choice == 'a':
                print("Adding...")
                weeklyEntries = getWeeklyEntries(weeklyEntries[["Fund", "Activity", "Facility", "Job", "Open"]])
                    
            if choice == 'r':
                index = True
                while index:
                    print(weeklyEntries[["Job"]])
                    index = input("Row number to remove: ")
                    try:
                        weeklyEntries = weeklyEntries.drop([int(index)])
                        weeklyEntries.reset_index(drop=True, inplace=True)
                    except:
                        index = False

    weeklyEntries["Week"] = "Week " + str(currentWeek)
    currentWeek += 1
    hours = []
    miles = []

    try:
        for i, row in weeklyEntries.iterrows():
            x = row["Job"]
            hours.append(float(input(f"Enter hours for {x}: ")))
            miles.append(float(input(f"Enter miles for {x}: ")))
        weeklyEntries["Hours"] = hours
        weeklyEntries["Miles"] = miles
        contractorEntries = pd.concat([contractorEntries, weeklyEntries])
        contractorEntries.reset_index(drop=True, inplace=True)
    except:
        continue
print(contractorEntries)
with pd.ExcelWriter("summary.xlsx", mode="a", if_sheet_exists="overlay") as writer:
    contractorEntries.to_excel(writer, sheet_name="Entries", startrow=writer.sheets["Entries"].max_row, header=False, index=False)