import csv
import logging

class PeopleCsvManager():    
    def __init__(self, csv_file="personer.csv") -> None:
        self.logger = logging.getLogger('O2A')
        self.__people = self.__readFile(csv_file)

    def getPersonData(self,person_outlook_name):
        self.logger.debug(f"Searching for {person_outlook_name} in CSV register")

        for person in self.__people:
            if person["outlook_name"] == person_outlook_name:
                aula_name = person["aula_name"]
                self.logger.debug(f"FOUND and should be replaced with {aula_name}")
                return aula_name

        self.logger.debug("NOT FOUND")
        return None

    def __readFile(self, csv_file="personer.csv"):
        people = []

        try:
            with open(csv_file, mode='r') as csv_file:
                csv_reader = csv.DictReader(csv_file,delimiter=";")
                line_count = 0
                for row in csv_reader:
                    if line_count == 0:
                        self.logger.debug(f'Column names are {"; ".join(row)}')
                        line_count += 1

                    person = {
                        "outlook_name" : row["Outlook navn"],
                        "aula_name" : row["AULA navn"]
                    }

                    people.append(person)

                    self.logger.debug(f'\t{row["Outlook navn"]} works in the {row["AULA navn"]} .')
                    line_count += 1

                self.logger.debug(people)
                self.logger.debug(f'Processed {line_count} lines.')
        except FileNotFoundError as e:
            self.logger.warning(f"CSV file '{csv_file}'' was not found. Continuing without.")
            self.logger.debug(e)

        return people

#pClass = PeopleCsvManager(csv_file="personer.csv")
#print(pClass.getPersonData("Fiktiv Fiktivsen"))