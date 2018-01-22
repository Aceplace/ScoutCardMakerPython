import csv

with open("Book1.csv") as csvfile:
    csvreader = csv.reader(csvfile)
    for row in csvreader:
        print(row)
