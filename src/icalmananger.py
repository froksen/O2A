import requests
from icalendar import Calendar
from setupmanager import SetupManager
import re
import sys
import logging

class IcalManager:

    def __init__(self):
        
        self.setupMgr = SetupManager()
        self.logger = logging.getLogger('O2A')


    def readAulaCalendarEvents(self, begin_date=False,end_date=False):
        self.logger.info("Reading Week calendar")
        week_calendar_events = self.readCalendarEvents(title="Ugekalender",url=self.setupMgr.get_aula_week_calendar_url())
        self.logger.info("Reading year calendar")
        year_calendar_events = self.readCalendarEvents(title="Årskalender",url=self.setupMgr.get_aula_year_calendar_url())

        self.logger.debug("Combining")
        return week_calendar_events | year_calendar_events

    def readCalendarEvents(self,title,url):
        #Downloading calendar
        try:
            self.logger.info("Downloading iCal-file for \"%s\"" %(title))
            r = requests.get(url)
            #Setting correct charset
            self.logger.debug("Setting charset")
            r.encoding = r.apparent_encoding
        except requests.exceptions.ConnectionError:
            self.logger.critical("UNABLE to download iCal-file. Is Url correct or are you connected to the internet?")
            self.logger.critical("Process stopped!")
            sys.exit()

        #print(r.text)

        #reading content of ical
        self.logger.info("Reading iCal-file for \"%s\"" %(title))
        gcal = Calendar.from_ical(r.text)

        #going through ical
        events = {}
        for component in gcal.walk():
            if component.name == "VEVENT":
                #print(component.get('summary'))
                #print(component.get('description'))
                #print(component.get('dtstart').dt)
                #print(component.get('dtend'))
                #print()
                #print(component.get('dtstamp'))

                class appointmentitem(object):
                    pass

                appointmentitem.subject = component.get('SUMMARY')
                appointmentitem.body = component.get('DESCRIPTION')
                appointmentitem.start = component.get('DTSTART').dt
                appointmentitem.end = component.get('DTEND').dt
                appointmentitem.location = component.get('LOCATION')

                #appointmentitem = {
                #    "subject" : component.get('SUMMARY'),
                #    "body" : component.get('DESCRIPTION'),
                #    "start" : component.get('DTSTART').dt,
                #    "end" : component.get('DTEND').dt,
                ##    "location" : component.get('LOCATION')
                #}

                description = component.get('DESCRIPTION')

                #Finds AULA-Url for event
                m = re.search('(?P<url>https?://[^\s]+)', description)
                if m:
                    aula_calendar_url = m.group("url").replace(",","")

                #Find GAID in description
                m1 = re.search('o2a_outlook_GlobalAppointmentID=\S*', description)
                if m1:
                    outlook_GlobalAppointmentID = m1.group(0)
                    outlook_GlobalAppointmentID = outlook_GlobalAppointmentID.split("=")[1].strip()

                #FINDS LMT in description
                m2 = re.search('o2a_outlook_LastModificationTime=\S* \S*\S\S:\S\S', description)
                if m2:
                    outlook_LastModificationTime = m2.group(0)
                    outlook_LastModificationTime = outlook_LastModificationTime.split("=")[1].strip()

                #if both GAID and LMT exists then add item to dict. 
                if m1 and m2:
                    events[outlook_GlobalAppointmentID]={
                        "appointmentitem":appointmentitem,
                        "aula_event_url":aula_calendar_url,
                        "outlook_GlobalAppointmentID":outlook_GlobalAppointmentID,
                        "outlook_LastModificationTime":outlook_LastModificationTime
                    }

        self.logger.info("Completed downloading and reading calendar \"%s\"" %(title))
        return events