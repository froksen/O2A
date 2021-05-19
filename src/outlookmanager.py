import win32com.client
import datetime as dt
import re
import sys
import logging
import time

class OutlookManager:
    def __init__(self):
        print("Outlook Manager Initialized")
        self.logger = logging.getLogger('O2A')

    def get_aulaevents_from_outlook(self,begin,end):

        def format_as_aula_date(outlook_date_time):
            outlook_date_time = str(outlook_date_time)
            date_part = outlook_date_time.split(" ")[0]
            date_part = date_part.split("-")

            date_part = date_part[2]+"/"+date_part[1]+"/"+date_part[0] 

            return date_part.strip()

        def format_as_aula_time(outlook_date_time):
            outlook_date_time = str(outlook_date_time)
            time_part = outlook_date_time.split(" ")[1]
            time_part = time_part.split(":")
            time_part = time_part[0] + ":" + time_part[1]
            return time_part.strip()
            #2021-03-04 10:00:00+00:00

        aulaEvents = {}

        events = self.get_personal_calendar(begin,end) #Finds all events
        
        for event in events: #Loops through
            categories = event.categories.split(";") #If event has multiple categories, then split

            for category in categories: #Loops through categories
                category = category.strip()

                if(category == "AULA"): #If category is AULA
                    #print("Subject")
                    #print(event.Subject)
                    #print("CurrentUser")
                    #print(self.get_personal_calendar_username())
                    #print("RequiredAttendees")
                    #print(event.RequiredAttendees)
                    
                    #attendeesList =event.RequiredAttendees.split(";")
                    addToInstitutionCalendar = False
                    for category2 in categories: #Loops through categories
                        category2 = category2.strip()
                        if category2 == "AULA: Institutionskalender":
                            addToInstitutionCalendar = True

                    aulaEvents[event.GlobalAppointmentID] = {"appointmentitem":event, 
                        "aula_startdate": format_as_aula_date(event.start),
                        "aula_enddate": format_as_aula_date(event.end),
                        "aula_starttime": format_as_aula_time(event.start),
                        "aula_endtime": format_as_aula_time(event.end),
                        "addToInstitutionCalendar" : addToInstitutionCalendar
                    }

        return aulaEvents

    def get_personal_calendar_username(self):
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        return ns.CurrentUser

    def get_personal_calendar(self,begin,end):
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")
        calendar = ns.GetDefaultFolder(9).Items

        return self.__get_calendar(calendar,begin,end)
        
    def __get_calendar(self,calendar,begin,end):
        calendar.Sort('[Start]')
        restriction = "[Start] >= '" + begin.strftime('%d/%m/%Y') + "' AND [END] <= '" + end.strftime('%d/%m/%Y') + "'"
        calendar = calendar.Restrict(restriction)
        
        return calendar
