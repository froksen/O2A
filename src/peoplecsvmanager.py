import csv
import pandas as pd

class PeopleCsvManager():    
    def __init__(self) -> None:
        pass

    def readFile(self, csvFile="personer.csv"):
        #TODO: Tilføj flere checks her...
        if csvFile == None:
            return None


        people = []
        with open(csvFile, mode='r') as csv_file:
            csv_reader = csv.DictReader(csv_file,delimiter=";")
            line_count = 0
            for row in csv_reader:
                if line_count == 0:
                    print(f'Column names are {"; ".join(row)}')
                    line_count += 1

                person = {
                    "outlook_name" : row["Outlook navn"],
                    "aula_name" : row["AULA navn"]
                }

                people.append(person)

                print(f'\t{row["Outlook navn"]} works in the {row["AULA navn"]} .')
                line_count += 1

            print(people)
            print(f'Processed {line_count} lines.')

            return csv_reader


pClass = PeopleCsvManager()
pClass.readFile()
