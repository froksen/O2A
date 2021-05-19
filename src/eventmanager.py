from outlookmanager import OutlookManager
from aulamanager import AulaManager
import datetime as dt
import time
import logging
import sys
import win32com.client
import keyring
from icalmananger import IcalManager
from setupmanager import SetupManager


class EventManager:
    def __init__(self):
        #Managers are init.
        self.aulamanager = AulaManager()
        self.outlookmanager = OutlookManager()
        self.icalmanager = IcalManager()
        self.setupmanager = SetupManager()

        #Sets logger
        self.logger = logging.getLogger('O2A')

    def update_aula_calendar(self, changes):


        #If no changes, then do nothing
        if len(changes['events_to_create'] or changes['events_to_remove']) <= 0:
            self.logger.info("No changes. Process completed")
            return

        # Create a browser
        #self.aulamanager.setBrowser(self.aulamanager.createBrowser(headless=True))

        #Gets AULA password and username from keyring
        aula_usr = self.setupmanager.get_aula_username()
        aula_pwd = self.setupmanager.get_aula_password()
        

        #Login to AULA
        if not self.aulamanager.login(aula_usr,aula_pwd) == True:
            self.logger.critical("Program stopped because unable to log in to AULA.")
            sys.exit()
            return

        for event_to_remove in changes['events_to_remove']:
            event_title = event_to_remove["appointmentitem"].subject
            event_url = event_to_remove["aula_event_url"]
            event_id = event_url.split("/")[7] #Should be regexp instead!
            #event_GlobalAppointmentID = event_to_remove["appointmentitem"].GlobalAppointmentID
            
            #Removing event
            self.logger.info("Attempting to REMOVE event: %s " %(event_title))
            #self.aulamanager.deleteEvent(event_id)

        #time.sleep(5)

        #Creation of event
        for event_to_create in changes['events_to_create']:
            event_title = event_to_create["appointmentitem"].subject
            start_date = event_to_create["aula_startdate"]
            end_date = event_to_create["aula_enddate"]
            start_time = event_to_create["aula_starttime"]
            end_time = event_to_create["aula_endtime"]
            start_dateTime = str(start_date).replace("/","-") + "T" + start_time + "+02:00"  # FORMAT: 2021-05-18T15:00:00+02:00
            end_dateTime = str(end_date).replace("/","-") + "T" + end_time + "+02:00" # FORMAT: 2021-05-18T15:00:00+02:00
            location = event_to_create["appointmentitem"].location 
            sensitivity = event_to_create["appointmentitem"].Sensitivity 
            description = "%s\no2a_outlook_GlobalAppointmentID=%s\n\no2a_outlook_LastModificationTime=%s" %(event_to_create["appointmentitem"].body,event_to_create["appointmentitem"].GlobalAppointmentID,event_to_create["appointmentitem"].LastModificationTime)
            allDay = event_to_create["appointmentitem"].AllDayEvent
            attendees = []
            attendee_ids = []
            isPrivate = False
            addToInstitutionCalendar = event_to_create["addToInstitutionCalendar"]

            #Sensitivity == 2 means private
            if sensitivity == 2:
                isPrivate = True

            #Only do this if organizer is current user
            #print("Current user: %s" %(self.outlookmanager.get_personal_calendar_username()))
            #print("Organizer: %s " %(event_to_create["appointmentitem"].Organizer))
            if str(self.outlookmanager.get_personal_calendar_username()).strip() == str(event_to_create["appointmentitem"].Organizer).strip(): 
                attendees = event_to_create["appointmentitem"].RequiredAttendees.split(";") #| event_to_create["appointmentitem"].OptionalAttendees.split(";") #Both optional and required attendees. In AULA they are the same.
                 
                for attendee in attendees:
                    if not self.aulamanager.findRecipient(attendee) == None:
                        attendee_ids.append(self.aulamanager.findRecipient(attendee))


            #Creating new event
            self.aulamanager.createEvent(title=event_title,description=description,startDateTime=start_dateTime,endDateTime=end_dateTime, attendee_ids = attendee_ids, addToInstitutionCalendar=addToInstitutionCalendar,allDay=allDay,isPrivate=isPrivate)



    def compare_calendars(self, begin, end):
        #Summary of changes
        self.logger.info(" ")
        self.logger.info("..:: Comparing Outlook and AULA events :: ...")
        self.logger.info("Between")
        self.logger.info(" Start datetime: %s" %(begin.strftime('%Y-%m-%d')))
        self.logger.info(" End datetime: %s" %(end.strftime('%Y-%m-%d')))
        self.logger.info(" ")

        if(begin.strftime('%Y-%m-%d') < dt.datetime.today().strftime('%Y-%m-%d')):
            self.logger.critical("Begin date must be today or in the future! Exitting.")
            sys.exit()

        #Finds all events from Outlook
        aulaevents_from_outlook = self.outlookmanager.get_aulaevents_from_outlook(begin, end)

        #Finds AULA events from ICal-calendar
        #outlookevents_from_aula = self.icalmanager.readAulaCalendarEvents()
        outlookevents_from_aula = self.aulamanager.getEvents(None,None)
        #events = self.getEvents(None, None)
        

        events_to_create = []
        events_to_remove = []

        self.logger.info("..:: CHANGES :: ...")

        #Checking for events that has been updated, and exists both places
        for key in aulaevents_from_outlook:
            if  key in outlookevents_from_aula:
                if str(aulaevents_from_outlook[key]["appointmentitem"].LastModificationTime) != outlookevents_from_aula[key]["outlook_LastModificationTime"]:
                    events_to_remove.append(outlookevents_from_aula[key])
                    self.logger.info("Event \"%s\" has been updated. Old entry will be removed, and a new will be created." %(outlookevents_from_aula[key]["appointmentitem"].subject))
                    self.logger.info(" - LastModificationTime from AULA: %s" %(outlookevents_from_aula[key]["outlook_LastModificationTime"]))
                    self.logger.info(" - LastModificationTime from Outlook: %s" %(aulaevents_from_outlook[key]["appointmentitem"].LastModificationTime))
                    self.logger.info(" - Outlook event GlobalAppointmentID: %s" %(aulaevents_from_outlook[key]["appointmentitem"].GlobalAppointmentID))
                    self.logger.info(" - AULA event GlobalAppointmentID: %s" %(outlookevents_from_aula[key]["outlook_GlobalAppointmentID"]))
                    #events_to_remove.append(outlookevents_from_aula[key])
                    events_to_create.append(aulaevents_from_outlook[key])

        #Checking for events that currently only exists in Outlook and should be created in AULA
        for key in aulaevents_from_outlook:
            if not key in outlookevents_from_aula:
                events_to_create.append(aulaevents_from_outlook[key])
                self.logger.info("Event \"%s\" width start date %s does not exists in AULA. Set to be created in AULA." %(aulaevents_from_outlook[key]["appointmentitem"].subject,aulaevents_from_outlook[key]["appointmentitem"].start))
                #self.logger.info(" - LastModificationTime from Outlook: %s" %(aulaevents_from_outlook[key]["appointmentitem"].LastModificationTime))
                #self.logger.info(" - Outlook event GlobalAppointmentID: %s" %(aulaevents_from_outlook[key]["appointmentitem"].GlobalAppointmentID))

        #Checking for events that currently only exists in AULA, and therefore should be deleted from AULA. 
        for key in outlookevents_from_aula:
            if not key in aulaevents_from_outlook:
                if not key in events_to_remove:
                    events_to_remove.append(outlookevents_from_aula[key])
                    self.logger.info("Event \"%s\"  only exists in AULA. Set to be removed from AULA." %(outlookevents_from_aula[key]["appointmentitem"].subject))
                    #self.logger.info(" - Appointment URL: %s" %(outlookevents_from_aula[key]["aula_event_url"]))
                    #self.logger.info(" - LastModificationTime from AULA: %s" %(outlookevents_from_aula[key]["outlook_LastModificationTime"]))
                    #self.logger.info(" - AULA event GlobalAppointmentID: %s" %(outlookevents_from_aula[key]["outlook_GlobalAppointmentID"]))

        #Summary of changes
        self.logger.info(" ")
        self.logger.info("..:: CHANGES SUMMARY :: ...")
        self.logger.info("Events to be created: %s" %(len(events_to_create)))
        self.logger.info("Events to be removed: %s" %(len(events_to_remove)))
        self.logger.info(" ")

        return {
                'events_to_create': events_to_create,
                'events_to_remove': events_to_remove
                }
    
        

