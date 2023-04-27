import sys
import pandas as pd
import datetime as dt

accounting = open("accounting.xlsx", "rb")
projects = pd.read_excel(accounting, sheet_name="Open Jobs")
ic = pd.read_excel(accounting, sheet_name="WSP-IC2018")

mileageRate = 0.655

#---Find contractor---#
contractor = pd.DataFrame()
while contractor.empty:
    contractorName = input("Enter contractor name: ")
    contractor = ic[ic["Name"] == contractorName]

    if contractor.empty:
        print("Contractor not found...")

#---Monthly entry---#
numWeeks = int(sys.argv[1])
weekEnding = (dt.datetime(int(sys.argv[2]), int(sys.argv[3]), int(sys.argv[4])) + dt.timedelta(days=6))
currentWeek = 1
monthlyTotals = pd.DataFrame(
    [{"Name": contractor["Name"].iloc[0],
        "Company": contractor["Company"].iloc[0],
        "Month": (weekEnding+dt.timedelta(days=14)).strftime("%B"),
        "Regular Hours": 0,
        "Overtime Hours": 0,
        "Miles": 0,
        "Total": 0}]
    )

while numWeeks > 0:
    #---Weekly entry---# 
    currentEntry = 1
    weeklyEntries = 1#int(input(f"Number of entries for week {currentWeek}: "))
    weeklyTotals = pd.DataFrame(
        [{"Name": contractor["Name"].iloc[0],
        "Company": contractor["Company"].iloc[0],
        "Week": "Week " + str(currentWeek),
        "Date": weekEnding.strftime("%b %d %Y"),
        "Regular Hours": 0,
        "Overtime Hours": 0,
        "Miles": 0,
        "Total": 0}]
        )
    
    while weeklyEntries > 0:    
        project = pd.DataFrame()
        isOpen = True
        
        projectNumber = "test"#input("Project Number: ")
        project = projects[projects["JOB NUMBER"] == projectNumber]

        if project.empty:
            print(f"Project {projectNumber} not found")
            isOpen = False
            project = pd.DataFrame([{"JOB NUMBER": projectNumber}])

        hours = float(input("Hours: "))
        miles = 0#float(input("Miles: "))

        otHours = 0
        regHours = 0
        totalHours = weeklyTotals["Regular Hours"].loc[0]
        currentHours = totalHours + hours
        rate = contractor["Rate"]
        subtotal = 0

        if totalHours > 40:
            otHours = hours
        elif currentHours > 40:
            otHours = currentHours - 40
            regHours = hours - otHours
        elif currentHours <= 40:
            regHours = hours

        subtotal += regHours * rate
        subtotal += otHours * (rate * 1.5)
        subtotal += miles * mileageRate

        weeklyTotals["Regular Hours"] += regHours
        weeklyTotals["Overtime Hours"] += otHours
        weeklyTotals["Miles"] += miles
        weeklyTotals["Total"] += float(subtotal)

        monthlyTotals["Regular Hours"] += regHours
        monthlyTotals["Overtime Hours"]+= otHours
        monthlyTotals["Miles"]+= miles
        monthlyTotals["Total"]+= float(subtotal)

        with pd.ExcelWriter("output.xlsx", mode="a", if_sheet_exists="overlay") as writer:
            entry = {"Name": contractor["Name"].iloc[0],
            "Company": contractor["Company"].iloc[0],
            "Week": "Week " + str(currentWeek),
            "Date": weekEnding.strftime("%b %d %Y"),
            "Project": project["JOB NUMBER"].iloc[0],
            "Open": isOpen,
            "Regular Hours": regHours,
            "Overtime Hours": otHours,
            "Miles": miles,
            "Total": subtotal}

            pd.DataFrame(entry).to_excel(writer, startrow=writer.sheets["Sheet1"].max_row, header=False, index=False)
            if weeklyEntries == 1:
                pd.DataFrame(weeklyTotals).to_excel(writer, sheet_name="Sheet2", startrow=writer.sheets["Sheet2"].max_row, header=False, index=False)
                if numWeeks == 1:
                    pd.DataFrame(monthlyTotals).to_excel(writer, sheet_name="Sheet3", startrow=writer.sheets["Sheet3"].max_row, header=False, index=False)
        weeklyEntries -= 1
        currentEntry += 1
    numWeeks -=1
    currentWeek += 1
    weekEnding = (weekEnding + dt.timedelta(days=7))
