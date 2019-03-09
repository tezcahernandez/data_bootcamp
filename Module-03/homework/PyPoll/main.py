import csv

filePath = "Resources/election_data.csv"

def printLine():
    print("----------------------------------")
    
with open(filePath, 'r') as csvfile:

    csvreader = csv.reader(csvfile, delimiter=',')
    header = next(csvreader)

    i = 0
    candidateCounters = { "Khan": 0, "Correy": 0, "Li": 0, "O'Tooley": 0}

    for row in csvreader:
        i += 1
        candidateCounters[row[2]] += 1

    print("Election Results")
    printLine()
    print(f"Total votes {i}")
    printLine()
    for key, value in candidateCounters.items():
        print(f" -  {key}:  {round( value / i * 100, 3)}% ({value} votes)")
    printLine()
    print(f"Winner: {sorted(candidateCounters.items(), key=lambda x: x[1], reverse=True)[0][0]}")
    printLine()
    # map(lambda x : print("*"), range(20))
    