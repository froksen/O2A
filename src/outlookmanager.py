import win32com.client
import datetime as dt
from datetime import timedelta
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
            categories_org = event.categories.split(";") #If event has multiple categories, then split

            #Makes sure that there are no whitespaces before or after
            categories = []
            for category in categories_org:
                #print(category)
                categories.append(str(category).strip())

            # If has category "AULA" then it should be added to AULA
            if 'AULA' in categories:
                addToInstitutionCalendar = False

                #If it also has category "AULA: Institutionskalender" then the event should be added to the instituionCalendar
                if 'AULA Institutionskalender' in categories: #Loops through categories
                    addToInstitutionCalendar = True

                #Fixes issue, where end in Allday events are pushed one day forward.
                #TODO: Make a better fix. 
                if event.AllDayEvent == True:
                    endDateTime_fix = event.end - timedelta(days=1)
                    event.end = endDateTime_fix

                    startDateTime_fix = event.start + timedelta(days=1)
                    event.start = startDateTime_fix

                #Array containing event information
                aulaEvents[event.GlobalAppointmentID] = {"appointmentitem":event, 
                    "aula_startdate": format_as_aula_date(event.start),
                    "aula_enddate": format_as_aula_date(event.end),
                    "aula_starttime": format_as_aula_time(event.start),
                    "aula_endtime": format_as_aula_time(event.end),
                    "addToInstitutionCalendar" : addToInstitutionCalendar
                }

                #print("ENDDATE")
                #print(aulaEvents[event.GlobalAppointmentID]["appointmentitem"].subject)
                #print(aulaEvents[event.GlobalAppointmentID]["aula_enddate"])
                #print(event.end)
                #time.sleep(2)

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
