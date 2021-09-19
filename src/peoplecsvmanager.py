import csv

class PeopleCsvManager():    
    def __init__(self, csv_file="personer.csv") -> None:
        self.__people = self.__readFile(csv_file)

    def getPersonData(self,person_outlook_name):
        print(f"Searching for {person_outlook_name} in CSV register")

        for person in self.__people:
            if person["outlook_name"] == person_outlook_name:
                aula_name = person["aula_name"]
                print(f"FOUND and should be replaced with {aula_name}")
                return aula_name

        print("NOT FOUND")
        return None

    def __readFile(self, csv_file="personer.csv"):
        #TODO: Tilføj flere checks her...
        if csv_file == None:
            print("File not found.")
            return None

        people = []
        with open(csv_file, mode='r') as csv_file:
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

            return people

#pClass = PeopleCsvManager(csv_file="personer.csv")
#print(pClass.getPersonData("Fiktiv Fiktivsen"))