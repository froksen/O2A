from aulamanager import AulaManager
import csv
import logging
import sys
from setupmanager import SetupManager
import time

class ContactsChecker():
    def __init__(self,csv_file="contacts_to_check.csv"):
        self.aulamgr = AulaManager()
        self.logger = logging.getLogger('O2A')
        self.setupmanager = SetupManager()
        self.__people = self.__readFile(csv_file)

        self.login_to_aula()



    def login_to_aula(self):
        #Gets AULA password and username from keyring
        aula_usr = self.setupmanager.get_aula_username()
        aula_pwd = self.setupmanager.get_aula_password()

        #Login to AULA
        if not self.aulamanager.login(aula_usr,aula_pwd) == True:
            self.logger.critical("Program stopped because unable to log in to AULA.")
            sys.exit()
            return

    def searchForPeople(self):
        for person in self.__people:
            #Searching for name in AULA
            search_result = self.aulamanager.findRecipient(person)

            if not search_result == None:
                self.logger.info("      Attendee %s was found in AULA!" %(person))
            else:
                self.logger.info("      Attendee %s was NOT found in AULA!" %(person))

            time.sleep(0.5)
            
            #self.logger.debug(f"Searching for {person_outlook_name} in AULA")

           # for person in self.__people:
               # if person["outlook_name"] == person_outlook_name:
                  #  aula_name = person["aula_name"]
                   # self.logger.debug(f"FOUND and should be replaced with {aula_name}")
                    #return aula_name

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
                        }

                        people.append(person)

                        self.logger.debug(f'\t{row["Outlook navn"]} found in CSV.')
                        line_count += 1

                    self.logger.debug(people)
                    self.logger.debug(f'Processed {line_count} lines.')
            except FileNotFoundError as e:
                self.logger.warning(f"CSV file '{csv_file}'' was not found. Continuing without.")
                self.logger.debug(e)

            return people

