from dateutil.relativedelta import relativedelta
from aulaevent import AulaEvent
from outlookmanager import OutlookManager
from aulamanager import AulaManager
import datetime as dt
import time
import logging
import sys
import win32com.client
import keyring
from setupmanager import SetupManager
from peoplecsvmanager import PeopleCsvManager
import itertools
import pytz

class EventManager:
    def __init__(self):
        #Managers are init.
        self.aulamanager = AulaManager()
        self.outlookmanager = OutlookManager()
        self.setupmanager = SetupManager()
        self.peoplecsvmanager = PeopleCsvManager()

        #Sets logger
        self.logger = logging.getLogger('O2A')

        self.login_to_aula()

    def login_to_aula(self):
        #Gets AULA password and username from keyring
        aula_usr = self.setupmanager.get_aula_username()
        aula_pwd = self.setupmanager.get_aula_password()
        

        #Login to AULA
        login_response = self.aulamanager.login(aula_usr,aula_pwd)
        if not login_response.status == True:
            self.logger.critical("Program stopped because unable to log in to AULA.")
            self.outlookmanager.send_a_mail(login_response)
            sys.exit()
            return

    def calulate_day_of_the_week_mask(self):
        olFriday = 32    # Friday
        olMonday = 2     # Monday
        olSaturday = 64  # Saturday
        olSunday = 1     # Sunday
        olThursday = 16  # Thursday
        olTuesday = 4    # Tuesday
        olWednesday = 8  # Wednesday

        days_list = [olMonday, olTuesday, olWednesday, olThursday,
                        olFriday, olSaturday, olSunday]

        data = []
        #Used to convert from value to string
        def day_of_week_convert(x):
                            x = int(x)
                            return {
                                olSunday: "sunday",
                                olMonday: "monday",
                                olTuesday: "tuesday",
                                olWednesday: "wednesday",
                                olThursday: "thursday",
                                olFriday: "friday",
                                olSaturday: "saturday",
                            }.get(x, "unknown")

        #Find all combinations of the days_list, and creates a data dict
        for L in range(0, len(days_list)+1):
            for subset in itertools.combinations(days_list, L):
                sum = 0
                days_text = []
                for i in subset:
                    sum = sum + i
                    days_text.append(day_of_week_convert(i))

                days_info = {
                    "days_integer": subset,
                    "days_string": days_text,
                    "sum": sum
                }     

                data.append(days_info)

        return data

    def get_day_of_the_week_mask(self,sum):
        days_combinations = self.calulate_day_of_the_week_mask()

        for day in days_combinations:
            if sum == day["sum"]:
                return day["days_integer"]

        return False

    def _basic_aula_event_actions(self, event):
        #If event has been created by some one else. Set in description that its the case.
        if not str(self.outlookmanager.get_personal_calendar_username()).strip() == str(event.outlook_organizer).strip(): 
            self.logger.debug("Event was created by another user. Appending to description")
            event.description = "<p><b>OBS:</b> Begivenheden er oprindelig oprettet af: %s" %(str(event.outlook_organizer).strip()) + "</p>" +  event.description

        #Only attempt to add attendees to event if created by the user them self. 
        if str(self.outlookmanager.get_personal_calendar_username()).strip() == str(event.outlook_organizer).strip(): 

            self.logger.info("Searching in AULA for attendees:")
            for attendee in event.outlook_required_attendees:
                attendee = attendee.strip()

                if attendee == str(event.outlook_organizer) or attendee == "":
                    self.logger.debug("     Attendee is organizer - Skipping")
                    continue

                #Removes potential emails from contact name
                attendee = attendee.split("(")[0].strip()

                #Checks if person should be replaced with other name from CSV-file
                csv_aula_name = self.peoplecsvmanager.getPersonData(attendee)

                if not csv_aula_name == None:
                    self.logger.info("      NOTE: Attendee %s Outlook name was found in CSV-file was replaced with %s" %(attendee,csv_aula_name))
                    attendee = csv_aula_name

                #Searching for name in AULA
                search_result = self.aulamanager.findRecipient(attendee)

                if not search_result == None:
                    self.logger.info("      Attendee %s was found in AULA!" %(attendee))
                    event.attendee_ids.append(search_result)
                else:
                    self.logger.info("      Attendee %s was NOT found in AULA!" %(attendee))

                time.sleep(0.5)

            return event


    def update_aula_calendar(self, changes):

        #If no changes, then do nothing
        if len(changes['events_to_create']) <= 0 and len(changes['events_to_remove']) <= 0 and len(changes['events_to_update']) <= 0:
            self.logger.info("No changes. Process completed")
            return

        for event_to_remove in changes['events_to_remove']:
            event_title = event_to_remove["appointmentitem"].subject
            event_id = event_to_remove["appointmentitem"].aula_id #Should be regexp instead!
            self.logger.info("Attempting to REMOVE event: %s " %(event_title))
            self.aulamanager.deleteEvent(event_id)

        for event_to_update in changes["events_to_update"]:
            event_to_update = self._basic_aula_event_actions(event_to_update)
            self.aulamanager.updateEvent(event_to_update)

        #Creation of event
        for event_to_create in changes['events_to_create']:
            event_to_create = self._basic_aula_event_actions(event_to_create)

            #Creating new event
            is_Recurring = False #TODO: Gør via variable
            if is_Recurring:

                #Read more about patterns: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.olrecurrencetype?view=outlook-pia
                def outlook_pattern_to_aula_pattern(x):
                    x = int(x)
                    return {
                        0: "daily",
                        1: "weekly",
                        2: "monthly"
                    }.get(x, "never")

                def day_of_week_convert(x):
                    x = int(x)
                    #print("day_of_week_convert")
                    #print(x)
                    return {
                        0: "daily",
                        1: "weekly",
                        2: "monthly"
                    }.get(x, "never")

                def day_of_week_convert(x):
                    from collections import deque

                olFriday = 32    # Friday
                olMonday = 2     # Monday
                olSaturday = 64  # Saturday
                olSunday = 1     # Sunday
                olThursday = 16  # Thursday
                olTuesday = 4    # Tuesday
                olWednesday = 8  # Wednesday
                from collections import deque
                weekDays = deque((olMonday, olTuesday, olWednesday, olThursday,
                                olFriday, olSaturday, olSunday))

                recurrence_pattern = event_to_create["appointmentitem"].GetRecurrencePattern() #Gets the pattern of the event. How it is repeated.
                recurrence_pattern_aula = outlook_pattern_to_aula_pattern(recurrence_pattern.RecurrenceType) #Gets the type, like if its daily etc. And converts it from Outlook-naming to AULA. 
                max_date = str(recurrence_pattern.PatternEndDate).split(" ")[0] #Only the date part is needed. EX: 2022-02-11 00:00:00+00:00 --> 2022-02-11
                interval = recurrence_pattern.Interval #How often event should be repeated. 
                day_of_week_mask = recurrence_pattern.DayOfWeekMask
                day_of_week_mask_list = self.get_day_of_the_week_mask(day_of_week_mask)

                if recurrence_pattern.RecurrenceType == 5:
                    self.logger.warning(f"NOTICE: Event {event_title} is set to be repeated YEARLY in outlook. This is currently not supported by Aula! Event will not be created, there for proces skipped.")
                    continue
                #if not day_of_week_mask in weekDays:
                #    self.logger.warning(f"NOTICE: Event {event_title} is set to be repeated more than one day a week. This is currently not supported!. Event will not be repeated, and might not be created.")
                #    day_of_week_mask_list = []
                #TODO: Få dette til at virke igen.
                self.aulamanager.createRecuringEvent(title=event_title,description=description,startDateTime=start_dateTime,endDateTime=end_dateTime,maxDate=max_date,pattern=recurrence_pattern_aula,interval=interval,weekmask=day_of_week_mask_list, location=location, attendee_ids = attendee_ids, addToInstitutionCalendar=addToInstitutionCalendar,allDay=allDay,isPrivate=isPrivate,hideInOwnCalendar=hideInOwnCalendar)
            else:
                self.aulamanager.createSimpleEvent(event_to_create)

    def __from_outlookobject_to_aulaevent(self,outlookobject):
        aula_event = AulaEvent()

        aula_event.id = ""
        aula_event.outlook_global_appointment_id = outlookobject["appointmentitem"].GlobalAppointmentID
        aula_event.outlook_organizer = outlookobject["appointmentitem"].Organizer
        aula_event.institution_code = ""
        aula_event.creator_inst_profile_id = ""
        aula_event.title = outlookobject["appointmentitem"].subject
        aula_event.type = "event"
        aula_event.outlook_body = outlookobject["appointmentitem"].body
        aula_event.location = outlookobject["appointmentitem"].location 
        aula_event.start_date = outlookobject["aula_startdate"]
        aula_event.end_date = outlookobject["aula_enddate"]
        aula_event.start_time = outlookobject["aula_starttime"]
        aula_event.end_time = outlookobject["aula_endtime"]
        aula_event.start_timezone  = outlookobject["aula_startdate_timezone"]
        aula_event.end_timezone = outlookobject["aula_enddate_timezone"]
        aula_event.outlook_last_modification_time = outlookobject["appointmentitem"].LastModificationTime
        aula_event.all_day = outlookobject["appointmentitem"].AllDayEvent
        aula_event.private = True if outlookobject["appointmentitem"].Sensitivity == 2 else False #Værdien 2 betyder privat
        aula_event.is_recurring = outlookobject["appointmentitem"].IsRecurring
        aula_event.hide_in_own_calendar = outlookobject["hideInOwnCalendar"]
        aula_event.add_to_institution_calendar = outlookobject["addToInstitutionCalendar"]
        aula_event.is_private = True if outlookobject["appointmentitem"].Sensitivity == 2 else False #Værdien 2 betyder privat
        aula_event.outlook_required_attendees = outlookobject["appointmentitem"].RequiredAttendees.split(";")

        return aula_event

    def compare_calendars(self, begin, end, force_update_existing_events = False):
        #Summary of changes
        self.logger.info(" ")
        self.logger.info("..:: Sammenligner Outlook og AULA kalenderne :: ...")
        self.logger.info("Mellem datoerne")
        self.logger.info(" Start: %s" %(begin.strftime('%Y-%m-%d')))
        self.logger.info(" End: %s" %(end.strftime('%Y-%m-%d')))
        self.logger.info(" ")

        if(begin.strftime('%Y-%m-%d') < dt.datetime.today().strftime('%Y-%m-%d')):
            self.logger.critical("Kritisk fejl: Start datoen skal være senest dags dato.")
            sys.exit()

        #Finds all events from Outlook
        from datetime import timedelta
        aulaevents_from_outlook = self.outlookmanager.get_aulaevents_from_outlook(begin, end)

        #Finds AULA events from ICal-calendar
        aulabegin = dt.datetime(year=begin.year,month=begin.month,day=begin.day) + dt.timedelta(days=-1)
        #aulaend = dt.datetime(year=end.year,month=end.month,day=end.day-1)
        outlookevents_from_aula = self.aulamanager.getEvents(aulabegin,end)
        #events = self.getEvents(None, None)
        

        events_to_create = []
        events_to_remove = []
        events_to_update = []

        self.logger.info("..:: CHANGES :: ...")



        #Springer over OUTLOOK begivenheder der ligger med start dato før d.d.
        events_to_keep = {}
        for key in aulaevents_from_outlook:
            dateobj = aulaevents_from_outlook[key]["appointmentitem"].start.replace(tzinfo=pytz.UTC)

            if dateobj <= dt.datetime.today().replace(tzinfo=pytz.UTC):
                self.logger.info("Outlook event \"%s\" that begins at \"%s\" is in the past. Skipped." %(aulaevents_from_outlook[key]["appointmentitem"].subject, aulaevents_from_outlook[key]["appointmentitem"].start))
                continue

            if aulaevents_from_outlook[key]["appointmentitem"].IsRecurring and aulaevents_from_outlook[key]["appointmentitem"].GetRecurrencePattern().RecurrenceType == 5:
                self.logger.info("NOTICE: Outlook event \"%s\" that begins at \"%s\" is set to be repeated YEARLY in outlook. This is currently not supported by Aula! Event will not be created, there for proces skipped." %(aulaevents_from_outlook[key]["appointmentitem"].subject, aulaevents_from_outlook[key]["appointmentitem"].start))
                continue

            events_to_keep[key] = self.__from_outlookobject_to_aulaevent(aulaevents_from_outlook[key]) #Hvis begivenheden er d.d. eller senere, da overføres til denne liste.

        aulaevents_from_outlook = events_to_keep #Renavngives listen.

        #Springer over AULA begivenheder der ligger med start dato før d.d.
        events_to_keep = {}
        for key in outlookevents_from_aula:

            date_string = outlookevents_from_aula[key]["appointmentitem"].start
            dateobj = dt.datetime.strptime(date_string,'%Y-%m-%dT%H:%M:%S%z') #2020-08-10T10:05:00+00:00
            dateobj = dateobj + dt.timedelta(hours=2)

            if dateobj <= dt.datetime.today().replace(tzinfo=pytz.UTC):
                self.logger.info("AULA event \"%s\" that begins at \"%s\" is in the past. Skipped." %(outlookevents_from_aula[key]["appointmentitem"].subject, outlookevents_from_aula[key]["appointmentitem"].start))
                continue

            events_to_keep[key] = outlookevents_from_aula[key] #Hvis begivenheden er d.d. eller senere, da overføres til denne liste.

        outlookevents_from_aula = events_to_keep#Renavngives listen.


        # TJEKKER FOR DULETTER FRA AULA. Altså samme begivenhed er oprettet flere gange. Hvis da, da fjernes den ene udgave.
        for key in outlookevents_from_aula:
            if outlookevents_from_aula[key]["isDuplicate"] == True:
                events_to_remove.append(outlookevents_from_aula[key])
                self.logger.info("Event \"%s\" that begins at \"%s\" only is a dublicated entry. Set to be removed from AULA." %(outlookevents_from_aula[key]["appointmentitem"].subject, outlookevents_from_aula[key]["appointmentitem"].start))

        #Checking for events that has been updated or needs to be forced updated, and exists both places
        for key in aulaevents_from_outlook:
            if  key in outlookevents_from_aula:
                
                #If forceupdate is enabled
                if force_update_existing_events == True:
                    self.logger.info("Event \"%s\" will be force updated." %(outlookevents_from_aula[key]["appointmentitem"].subject))

                    #Adds AULA eventid to array
                    aulaevents_from_outlook[key]["appointmentitem"].id = outlookevents_from_aula[key]["appointmentitem"].aula_id
                    events_to_update.append(aulaevents_from_outlook[key]) 

                    #Prevents the same event to be set en both update metods. 
                    continue
  
                #If event has been updated, but force update is not set.
                if str(aulaevents_from_outlook[key].outlook_last_modification_time) != outlookevents_from_aula[key]["outlook_LastModificationTime"]:
                    #events_to_remove.append(outlookevents_from_aula[key])
                    self.logger.info("Event \"%s\" has been updated in Outlook. Will attempt to do the same in AULA." %(outlookevents_from_aula[key]["appointmentitem"].subject))
                    self.logger.info(" - LastModificationTime from AULA: %s" %(outlookevents_from_aula[key]["outlook_LastModificationTime"]))
                    self.logger.info(" - LastModificationTime from Outlook: %s" %(aulaevents_from_outlook[key].outlook_last_modification_time))
                    self.logger.info(" - Outlook event GlobalAppointmentID: %s" %(aulaevents_from_outlook[key].outlook_global_appointment_id))
                    self.logger.info(" - AULA event GlobalAppointmentID: %s" %(outlookevents_from_aula[key]["outlook_GlobalAppointmentID"]))
                    #events_to_remove.append(outlookevents_from_aula[key])
                    #events_to_create.append(aulaevents_from_outlook[key]) 

                    #Adds AULA eventid to array
                    aulaevents_from_outlook[key].id = outlookevents_from_aula[key]["appointmentitem"].aula_id
                    events_to_update.append(aulaevents_from_outlook[key]) 

        #Checking for events that currently only exists in Outlook and should be created in AULA
        for key in aulaevents_from_outlook:
            if not key in outlookevents_from_aula:
                events_to_create.append(aulaevents_from_outlook[key])
                self.logger.info("Event \"%s\" that begins at \"%s\" does not exists in AULA. Set to be created in AULA." %(aulaevents_from_outlook[key].title, aulaevents_from_outlook[key].start_date))

        #Checking for events that currently only exists in AULA, and therefore should be deleted from AULA. 
        for key in outlookevents_from_aula:

            if not key in aulaevents_from_outlook:
                if not key in events_to_remove:
                    events_to_remove.append(outlookevents_from_aula[key])
                    self.logger.info("Event \"%s\" that begins at \"%s\" only exists in AULA. Set to be removed from AULA." %(outlookevents_from_aula[key]["appointmentitem"].subject, outlookevents_from_aula[key]["appointmentitem"].start))

        #Summary of changes
        self.logger.info(" ")
        self.logger.info("..:: CHANGES SUMMARY :: ...")
        self.logger.info("Events to be created: %s" %(len(events_to_create)))
        self.logger.info("Events to be updated: %s" %(len(events_to_update)))
        self.logger.info("Events to be removed: %s" %(len(events_to_remove)))
        self.logger.info(" ")

        #time.sleep(10)

        return {
                'events_to_create': events_to_create,
                'events_to_remove': events_to_remove,
                'events_to_update' : events_to_update
                }