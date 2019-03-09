import csv

filePath = "Resources/budget_data.csv"
    
with open(filePath, 'r') as csvfile:

    csvreader = csv.reader(csvfile, delimiter=',')
    header = next(csvreader)

    i = 0
    globalSummatory = 0
    changeSummatory = 0
    greatestIncreaseAmount = 0
    greatestIncreasePeriod = ""
    greatestDecreaseAmount = 0
    greatestDecreasePeriod = ""

    for row in csvreader:
        globalSummatory += int(row[1])
        if(i > 0):
            change = int(row[1]) - int(lastRow[1])
            changeSummatory += change
            if(change > greatestIncreaseAmount):
                greatestIncreaseAmount = change
                greatestIncreasePeriod = row[0]
            elif(change < greatestDecreaseAmount):
                 greatestDecreaseAmount = change
                 greatestDecreasePeriod = row[0]


        lastRow = row
        i += 1

    print("Financial Analysis")
    print("----------------------------")
    print(f"Total Months: {i}")
    print(f"Total: ${globalSummatory}")
    print(f"Average  Change: ${round(changeSummatory / (i - 1), 2)}")
    print(f"Greatest Increase in Profits: {greatestIncreasePeriod} $({greatestIncreaseAmount})")
    print(f"Greatest Decrease in Profits: {greatestDecreasePeriod} $({greatestDecreaseAmount})")