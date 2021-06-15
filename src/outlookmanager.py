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
            if 'AULA' in categories or 'AULA Institutionskalender' in categories:
                addToInstitutionCalendar = False
                hideInOwnCalendar = False

                if not 'AULA' in categories and 'AULA Institutionskalender' in categories:
                    hideInOwnCalendar = True

                #If it also has category "AULA: Institutionskalender" then the event should be added to the instituionCalendar
                if 'AULA Institutionskalender' in categories: #Loops through categories
                    addToInstitutionCalendar = True

                #Fixes issue, where end in Allday events are pushed one day forward.
                #TODO: Make a better fix. 
                if event.AllDayEvent == True:
                    try:
                        endDateTime_fix = event.end - timedelta(days=1)
                        event.end = endDateTime_fix

                        startDateTime_fix = event.start + timedelta(days=1)
                        event.start = startDateTime_fix
                    except: 
                        pass
                        #print("SKIPPED")


                #Array containing event information
                aulaEvents[event.GlobalAppointmentID] = {"appointmentitem":event, 
                    "aula_startdate": format_as_aula_date(event.start),
                    "aula_enddate": format_as_aula_date(event.end),
                    "aula_starttime": format_as_aula_time(event.start),
                    "aula_endtime": format_as_aula_time(event.end),
                    "hideInOwnCalendar" : hideInOwnCalendar,
                    "addToInstitutionCalendar" : addToInstitutionCalendar
                }
                
                #print("ENDDATE")
                #print(aulaEvents[event.GlobalAppointmentID]["appointmentitem"].subject)
                #print(aulaEvents[event.GlobalAppointmentID]["aula_enddate"])
                #print(event.end)
                #time.sleep(2)

        return aulaEvents

    def send_a_mail(self, message_to_send=""):
        #FROM: https://gist.github.com/vinovator/0a6d653c22c32ab67e11
        outlook = win32com.client.Dispatch("Outlook.Application")

        exchange_user = outlook.Session.CurrentUser.AddressEntry.GetExchangeUser()
        ownEmailAdress = exchange_user.PrimarySmtpAddress

        self.logger.debug("Exchange user " + str(exchange_user))
        self.logger.debug("Exchange user email " + ownEmailAdress)
        if ownEmailAdress == None:
            return

       #     Outlook VBA Reference 
       # 0 - olMailItem
       # 1 - olAppointmentItem
       # 2 - olContactItem
       # 3 - olTaskItem
       # 4 - olJournalItem
       # 5 - olNoteItem 
       # 6 - olPostItem
       # 7 - olDistributionListItem
        mail = outlook.CreateItem(0)

        mail.To = ownEmailAdress
        #mail.CC = "mail2@example.com"
        #mail.BCC = "mail3@example.com"

        mail.Subject = "Outlook2Aula - Opmærksomhed på fejl"

        # Using "Body" constructs body as plain text
        # mail.Body = "Test mail body from Python"

        """
        Using "HtmlBody" constructs body as html text
        default font size for most browser is 12
        setting font size to "-1" might set it to 10
        """
        mail.HTMLBody = f"""
        <html>
        <head></head>
        <body>
            <font color="DarkBlue" size=-1 face="Arial">
            <p>Hej {str(exchange_user)}!<br>
            I forbindelse med at Outlook2Aula overførselsprogrammet prøvede at køre på din computer, da skete der en fejl.<br>
            Problemet er, at <u>programmet ikke kan logge på AULA.</u> <br><br> 
            <u>Det er typisk fordi du har ændret din adgangskode</u>. Det kræves derfor, at du geninstaster din adgangskode i programmet.  <br><br>
            Hvis det ikke er tilfældet, og denne fejl bliver ved med at blive meldt, da kontakt Ole Frandsen (olfr@sonderborg.dk) eller Jesper Qvist (jeqv@sonderborg.dk).
            </p>
            <p>Venlig hilsen <br> Outlook2Aula overførselsprogrammet</p>
            </font>
        </body>
        </html>
        """

        """
        Set the format of mail
        1 - Plain Text
        2 - HTML
        3 - Rich Text
        """
        mail.BodyFormat = 2

        # Instead of sending the message, just display the compiled message
        # Useful for visual inspection of compiled message
        #mail.Display(True)

        # Send the mail
        # Use this directly if there is no need for visual inspection
        mail.Send()
        
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
