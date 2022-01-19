# Imports
from sys import getprofile
import requests                 # Perform http/https requests
from bs4 import BeautifulSoup   # Parse HTML pages
import json                     # Needed to print JSON API data
import logging
import re
from dateutil.relativedelta import relativedelta
import datetime
import time

#THIS CODE IS LARGELY INSPIRED BY CODE FROM https://helmstedt.dk/2020/05/et-lille-kig-paa-aulas-api/

class AulaManager:
    def __init__(self):
        # Start requests session
        print("AULA Manager Initialized")
        self.session = requests.Session()
        self.__profilesByLogin = ""
        
        #Sets logger
        self.logger = logging.getLogger('O2A')

    def setProfilesByLogin(self,profile):
        self.__profilesByLogin = profile

    def getProfilesByLogin(self):
        return self.__profilesByLogin

    def getProfileinstitutionCode(self):
        profiles = self.getProfilesByLogin()['data']['profiles']

        for profile in profiles:
            if profile['institutionProfiles'][0]['role'] == "employee":
                return profile['institutionProfiles'][0]['institutionCode']

        def getProfileId(self):
            profiles = self.getProfilesByLogin()['data']['profiles']

            for profile in profiles:
                if profile['institutionProfiles'][0]['role'] == "employee":
                    return profile['institutionProfiles'][0]['id']

        #return self.getProfilesByLogin()['data']['profiles'][0]['institutionProfiles'][0]['id']

    def getSession(self):
        return self.session

    def getAulaApiUrl(self):
        return 'https://www.aula.dk/api/v12/'

    def getEventsForInstitutions(self,profileId,instCodes, startDateTime, endDateTime):
        session = self.getSession()
        url = self.getAulaApiUrl()

        params = {
            'method': 'calendar.getEventsForInstitutions',
            "instCodes":[instCodes],
            "start":startDateTime,"end":endDateTime
            }

        events = []        
        #FORMAT:"2021-05-17 08:00:00.0000+02:00"

        url = url+"?method=calendar.getEventsForInstitutions&instCodes[]="+str(instCodes)+"&start="+startDateTime.replace("T+","%2B")+"&end="+endDateTime.replace("T+","%2B")

        response = session.get(url).json()
        #response = session.get(url).json()
        #print(json.dumps(response, indent=4))
        #print(response)
        for event in response["data"]:
            if(event["type"] == "event" and profileId == event["creatorInstProfileId"]):
                events.append(event)

        return events

    def getEventsByProfileIdsAndResourceIds(self,profileId, startDateTime, endDateTime):
        session = self.getSession()
        url = self.getAulaApiUrl()

        params = {
            'method': 'calendar.getEventsByProfileIdsAndResourceIds',
            }

        events = []        
        #FORMAT:"2021-05-17 08:00:00.0000+02:00"
        data = {"instProfileIds":[profileId],"resourceIds":[],"start":startDateTime,"end":endDateTime}

        response = session.post(url, params=params, json=data).json()
        #response = session.get(url).json()
        #print(json.dumps(response, indent=4))

        try:
            for event in response["data"]:
                if(event["type"] == "event" and profileId == event["creatorInstProfileId"]):
                    events.append(event)
        except TypeError as e:
            self.logger.critical("Der skete en fejl:")
            self.logger.critical(e)

        return events

    def getProfileId(self):
        # print(self.getProfilesByLogin()['data']['profiles'][0]['institutionProfiles'])

            profiles = self.getProfilesByLogin()['data']['profiles']

            for profile in profiles:
                if profile['institutionProfiles'][0]['role'] == "employee":
                    return profile['institutionProfiles'][0]['id']

    def getEventById(self,event_id):
        session = self.getSession()
        url = self.getAulaApiUrl()

        params = {
            'method': 'calendar.getEventById',
            "eventId": event_id,
            }

        response  = session.get(url, params=params).json()
        #print(json.dumps(response, indent=4))
        return response
        try:
            recipient_profileid = response["data"]["results"][0]["docId"] #Appearenly its docId and not profileId
            print(recipient_profileid)

            return int(recipient_profileid)

        except:
            return None

    def findRecipient(self,recipient):
        session = self.getSession()
        url = self.getAulaApiUrl()

        params = {
            'method': 'search.findRecipients',
            "text": recipient,
            "query": recipient,
            "id": str(self.getProfileId()),
            "typeahead": "true",
            "limit": "100",
            "scopeEmployeesToInstitution" : "true",
            "instCode": str(self.getProfileinstitutionCode()),
            "fromModule":"event",
            "docTypes[]":"Profile",
            "docTypes[]":"Group"
            }

        #url = " https://www.aula.dk/api/v11/?method=search.findRecipients&text=Stefan&query=Stefan&id=779467&typeahead=true&limit=100&scopeEmployeesToInstitution=false&fromModule=event&instCode=537007&docTypes[]=Profile&docTypes[]=Group"
        url = url+"?method=search.findRecipients&text="+recipient+"&query="+recipient+"&id="+str(self.getProfileId())+"&typeahead=true&limit=100&scopeEmployeesToInstitution=false&fromModule=event&instCode="+str(self.getProfileinstitutionCode())+"&docTypes[]=Profile&docTypes[]=Group"
        
        response  = session.get(url, params=params).json()
        #response = session.get(url).json()
        #print(json.dumps(response, indent=4))
        recipient_profileid = -1
        try:
            for result in response["data"]["results"]:
                if result["portalRole"] == "employee":
                    recipient_profileid = result["docId"] #Appearenly its docId and not profileId

                    return int(recipient_profileid)
           

        except:
            return None


    def deleteEvent(self, eventId):
        session = self.getSession()
        url = self.getAulaApiUrl()

        params = {
            'method': 'calendar.deleteEvent'
            }

        data = {
            "id":eventId
        }

        response  = session.post(url, params=params, json=data).json()
        #print(json.dumps(response, indent=4))

        if(response["status"]["message"] == "OK"):
            self.logger.info("Event was removed!")
        else:
            self.logger.warning("Event was not removed!")

    def teams_url_fixer(self,text):
        #Patterns for all the different parts of the Teams Meeting
        pattern_teams_meeting="Klik her for at deltage i mødet <https:\/\/teams.microsoft.com\/l\/meetup-join/.*" 
        pattern_know_more = "Få mere at vide <https:\/\/aka.ms\/JoinTeamsMeeting"
        pattern_meeting_options = "Mødeindstillinger <https:\/\/teams.microsoft.com\/meetingOptions.*"
        url_pattern = 'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'

        #Looks for all the parts
        teams_meeting = re.search(pattern_teams_meeting,text)
        know_more = re.search(pattern_know_more,text)
        meeting_options = re.search(pattern_meeting_options,text)

        if teams_meeting and know_more and meeting_options:
            self.logger.info("Microsoft Teams meeting detected. Fixing urls.")

        #If they are found, then do differnt things. 
        if teams_meeting:
            url = re.search(url_pattern,teams_meeting.group(0)).group(0).replace(">","")
            text = re.sub(pattern_teams_meeting,'<p><a href=\"%s" target=\"_blank\" rel=\"noopener\">%s</a></p>'%(url,"Klik her for at deltage i mødet"),text)

        if know_more:
            url = re.search(url_pattern,know_more.group(0)).group(0).replace(">","")
            text = re.sub(pattern_know_more,'<a href=\"%s" target=\"_blank\" rel=\"noopener\">%s</a>'%(url,"Få mere at vide"),text)

        if meeting_options:
            url = re.search(url_pattern,meeting_options.group(0)).group(0).replace(">","")
            text = re.sub(pattern_meeting_options,'<a href=\"%s" target=\"_blank\" rel=\"noopener\">%s</a>'%(url,"Mødeindstillinger"),text)

        return text

    def url_fixer(self,text):
        pattern_teams = "https:\/\/teams.microsoft.com\/l\/meetup-join"
        found = re.search(pattern_teams,text)

        if found:
            text = re.sub("<","",text)
            text = re.sub(">","",text)

        #print(text)

       # return
        #return text
        pattern = 'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+'

        urls_found = re.findall(pattern, text)

        #print(urls_found)
        if urls_found:
            for url in urls_found:
                #print("URL")
                #print(url)
                #print ("/URL")

                text = re.sub(re.escape(url),'<a href=\"%s" target=\"_blank\" rel=\"noopener\">%s</a>'%(url,url),text)
        return text
            #foundText = m1.group(0)

    def updateEvent(self, event_id, title, description, startDateTime, endDateTime, attendee_ids = [], location = "", addToInstitutionCalendar = False, allDay = False, isPrivate = False, hideInOwnCalendar = False):
        session = self.getSession()
        url = self.getAulaApiUrl()

        print("update")

        params = {
            'method': 'calendar.updateSimpleEvent'
            }

        description = self.teams_url_fixer(description)

        data = {
            "creator":{"id":self.getProfileId()},
            "institutionCode":self.getProfileinstitutionCode()
            ,"description":description,
            'primaryResource': location,
            'primaryResourceText' : location,
            "additionalResources":[],
            "additionalResourceText":None,
            "invitees":[],
            "invitedGroups":[],
            "attachments":[],
            "pendingMedia":False,
            "timeSlot":None,
            "vacationRegistration":None,
            "isDeleted":False,
            "eventClass":"basic",
            "responseDeadline":None,
            "isDeadlineExceeded":False,
            "hideInOwnCalendar":hideInOwnCalendar,
            "invitedGroupHomeChildren":[],
            "id":event_id,
            "title":title,
            "allDay":allDay,
            "startDateTime": startDateTime , #"2021-10-03T10:10:00.0000+02:00",
            "endDateTime":endDateTime, #"2021-10-03T12:00:00.0000+02:00",
            #"oldEndDateTime":"2021-10-03T10:00:00+00:00",
            #"oldStartDateTime":"2021-10-03T01:10:00+00:00",
            "responseRequired":True,
            "private":isPrivate,
            "type":"event",
            "addedToInstitutionCalendar":False,
            #"start":"2021-10-03T01:10:00+00:00",
            #"end":"2021-10-03T10:00:00+00:00",
            "invitedGroupHomes":[],
            "addToInstitutionCalendar":addToInstitutionCalendar,
            "additionalLocations":[],
            "resources":[],
            "pattern":"never",
            "occurenceLimit":0,
            "weekdayMask":[False,False,False,False,False,False,False],
            "maxDate":None,
            "interval":0,
            "eventId":event_id,
            "isPrivate":isPrivate,
            "inviteeIds": attendee_ids, #[],
            "invitedGroupIds":[],
            "resourceIds":[],
            "additionalLocationIds":[],
            "additionalResourceIds":[],
            "attachmentIds":[],
            "isEditEvent":True
            }

        response_calendar = session.post(url, params=params, json=data).json()
        #print(json.dumps(response_calendar, indent=4))

        if(response_calendar["status"]["message"] == "OK"):
            self.logger.info("Event \"%s\" with start date %s was SUCCESSFULLY updated" %(title,startDateTime))
        else:
            self.logger.warning("Event \"%s\" with start date %s was UNSUCCESSFULLY updated" %(title,startDateTime))

    def createSimpleEvent(self, title, description, startDateTime, endDateTime, attendee_ids = [], location = "", addToInstitutionCalendar = False, allDay = False, isPrivate = False, hideInOwnCalendar = False):
        session = self.getSession()
        
        #print("START: %s" %(startDateTime))
        #print("END: %s" %(endDateTime))
        #return

        # All API requests go to the below url
        # Each request has a number of parameters, of which method is always included
        # Data is returned in JSON
        url = self.getAulaApiUrl()
        
        ### First example API request ###
        params = {
            'method': 'calendar.createSimpleEvent'
        }

        description = self.teams_url_fixer(description)

        data = {
            'title': title,
            'description': description,
            'startDateTime': startDateTime, # 2021-05-18T14:30:00.0000+02:00
            'endDateTime': endDateTime, # '2021-05-18T15:00:00.0000+02:00'
            'startDate': datetime.datetime.today().strftime('%Y-%m-%d'), #Is always today
            'endDate': datetime.datetime.today().strftime('%Y-%m-%d'), # is always today
            #'startTime': '12:00:19', 
            #'endTime': '12:30:19',
            'id': '',
            'institutionCode': self.getProfileinstitutionCode(),
            'creatorInstProfileId': self.getProfileId(),
            'type': 'event',
            'allDay': allDay,
            'private': isPrivate,
            'primaryResource': location,
            'primaryResourceText' : location,
            'additionalLocations': [],
            'invitees': [],
            'invitedGroups': [],
            'invitedGroupIds': [],
            'invitedGroupHomes': [],
            'responseRequired': True,
            'responseDeadline': None,
            'resources': [],
            'attachments': [],
            'oldStartDateTime': '',
            'oldEndDateTime': '',
            'isEditEvent': False,
            'addToInstitutionCalendar': addToInstitutionCalendar,
            'hideInOwnCalendar': hideInOwnCalendar,
            'inviteeIds': attendee_ids,
            'additionalResources': [],
            'pattern': 'never',
            'occurenceLimit': 0,
            'weekdayMask': [
                False,
                False,
                False,
                False,
                False,
                False,
                False
            ],
            'maxDate': None,
            'interval': 0,
            'lessonId': '',
            'noteToClass': '',
            'noteToSubstitute': '',
            'eventId': '',
            'isPrivate': isPrivate,
            'resourceIds': [],
            'additionalLocationIds': [],
            'additionalResourceIds': [],
            'attachmentIds': []
        }

        response_calendar = session.post(url, params=params, json=data).json()
        #print(json.dumps(response_calendar, indent=4))

        if(response_calendar["status"]["message"] == "OK"):
            self.logger.info("Event \"%s\" with start date %s was SUCCESSFULLY created" %(title,startDateTime))
        else:
            self.logger.warning("Event \"%s\" with start date %s was UNSUCCESSFULLY created" %(title,startDateTime))


    def createRecuringEvent(self, title, description, startDateTime, endDateTime, maxDate, pattern, interval, weekmask = [], attendee_ids = [], location = "", addToInstitutionCalendar = False, allDay = False, isPrivate = False, hideInOwnCalendar = False):
        olFriday = 32    # Friday
        olMonday = 2     # Monday
        olSaturday = 64  # Saturday
        olSunday = 1     # Sunday
        olThursday = 16  # Thursday
        olTuesday = 4    # Tuesday
        olWednesday = 8  # Wednesday
        
        session = self.getSession()        
        #print("START: %s" %(startDateTime))
        #print("END: %s" %(endDateTime))
        #return

        # All API requests go to the below url
        # Each request has a number of parameters, of which method is always included
        # Data is returned in JSON
        url = self.getAulaApiUrl()
        
        ### First example API request ###
        params = {
            'method': 'calendar.createRepeatingEvent'
        }

        description = self.teams_url_fixer(description)

        data = {
            'title': title,
            'description': description,
            'startDateTime': startDateTime, # 2021-05-18T14:30:00.0000+02:00
            'endDateTime': endDateTime, # '2021-05-18T15:00:00.0000+02:00'
            'startDate': datetime.datetime.today().strftime('%Y-%m-%d'), #Is always today
            'endDate': datetime.datetime.today().strftime('%Y-%m-%d'), # is always today
            #'startTime': '12:00:19', 
            #'endTime': '12:30:19',
            'id': '',
            'institutionCode': self.getProfileinstitutionCode(),
            'creatorInstProfileId': self.getProfileId(),
            'type': 'event',
            'allDay': allDay,
            'private': isPrivate,
            'primaryResource': location,
            'primaryResourceText' : location,
            'additionalLocations': [],
            'invitees': [],
            'invitedGroups': [],
            'invitedGroupIds': [],
            'invitedGroupHomes': [],
            'responseRequired': True,
            'responseDeadline': None,
            'resources': [],
            'attachments': [],
            'oldStartDateTime': '',
            'oldEndDateTime': '',
            'isEditEvent': False,
            'addToInstitutionCalendar': addToInstitutionCalendar,
            'hideInOwnCalendar': hideInOwnCalendar,
            "repeating": {
                "pattern": "daily",
                "interval": 1,
                "weekdayMask": [
                False if not olSunday in weekmask else True,
                False if not olMonday in weekmask else True,
                False if not olTuesday in weekmask else True,
                False if not olWednesday in weekmask else True,
                False if not olThursday in weekmask else True,
                False if not olFriday in weekmask else True,
                False if not olSaturday in weekmask else True,
                ],
                "occurenceLimit": 0,
                "maxDate": maxDate
            },
            'inviteeIds': attendee_ids,
            'additionalResources': [],
            'pattern': pattern, #EX "daily"
            'occurenceLimit': 0,
            'weekdayMask': [
                False if not olSunday in weekmask else True,
                False if not olMonday in weekmask else True,
                False if not olTuesday in weekmask else True,
                False if not olWednesday in weekmask else True,
                False if not olThursday in weekmask else True,
                False if not olFriday in weekmask else True,
                False if not olSaturday in weekmask else True,
            ],
            "maxDate": maxDate, #EX: "2022-01-28"
            'interval': interval, #EX: 1
            'lessonId': '',
            'noteToClass': '',
            'noteToSubstitute': '',
            'eventId': '',
            'isPrivate': isPrivate,
            'resourceIds': [],
            'additionalLocationIds': [],
            'additionalResourceIds': [],
            'attachmentIds': []
        }

        response_calendar = session.post(url, params=params, json=data).json()
        #print(json.dumps(response_calendar, indent=4))

        if(response_calendar["status"]["message"] == "OK"):
            self.logger.info("Recuring event \"%s\" with start date %s was SUCCESSFULLY created" %(title,startDateTime))
        else:
            self.logger.warning("Recuring event \"%s\" with start date %s was UNSUCCESSFULLY created" %(title,startDateTime))

    def getProfile(self):
                # All API requests go to the below url
        # Each request has a number of parameters, of which method is always included
        # Data is returned in JSON

        session  = self.getSession()

        url = self.getAulaApiUrl()
        
        ### First example API request ###
        params = {
            'method': 'profiles.getProfilesByLogin'
            }
        # Perform request, convert to json and print on screen
        response_profile = session.get(url, params=params).json()
        print(json.dumps(response_profile, indent=4))

    def login(self, username, password):
        class LoginResponse:
            def __init__(self):
                self.status = False
                self.error_messages = []
                self.username = ""

        login_response = LoginResponse()

        # User info
        user = {
            'username': username,
            'password': password
            }
        
        # Start requests session
        session = self.getSession()
            
        # Get login page
        try:
            url = 'https://login.aula.dk/auth/login.php?type=unilogin'
            response = self.session.get(url)
        except requests.exceptions.ConnectionError as e:
            self.logger.critical("Det var ikke muligt, at oprette forbindelse til UNI-login dialogen")
            self.logger.critical(e)

        # Login is handled by a loop where each page is first parsed by BeautifulSoup.
        # Then the destination of the form is saved as the next url to post to and all
        # inputs are collected with special cases for the username and password input.
        # Once the loop reaches the Aula front page the loop is exited. The loop has a
        # maximum number of iterations to avoid an infinite loop if something changes
        # with the Aula login.
        counter = 0
        success = False
        self.logger.info("Attempting to login to AULA")
        while success == False and counter < 10:
            try:
                # Parse response using BeautifulSoup
                soup = BeautifulSoup(response.text, "html.parser")
                # Get destination of form element (assumes only one)
                url = soup.form['action']   
                
                # If form has a destination, inputs are collected and names and values
                # for posting to form destination are saved to a dictionary called data
                if url:
                    #Prints if any errors on login-dialog is present
                    login_errors = soup.find_all("span", {"class": "form-error-message"})

                    for login_error in login_errors:
                        self.logger.critical("UNI-Login error message: " + str(login_error.text))
                        login_response.error_messages.append("UNI-Login error message: " + str(login_error.text))
                        counter = 10 #Breaks the loop. TODO: MAKE Pythonic

                    # Get all inputs from page
                    inputs = soup.find_all('input')

                    # Check whether page has inputs
                    if inputs:
                        # Create empty dictionary 
                        data = {}
                        # Loop through inputs
                        for input in inputs:
                            # Some inputs may have no names or values so a try/except
                            # construction is used.
                            try:
                                # Save username if input is a username field
                                if input['name'] == 'username':
                                    data[input['name']] = user['username']
                                    self.logger.debug("Login-field: Username FOUND")
                                # Save password if input is a password field
                                elif input['name'] == 'password':
                                    data[input['name']] = user['password']
                                    self.logger.debug("Login-field: password FOUND")
                                #Selects login type, as employee this is "MEDARBEJDER_EKSTERN"
                                elif input['name'] == 'selected-aktoer':
                                    data[input['name']] = "MEDARBEJDER_EKSTERN"
                                    self.logger.debug("UNI-Login: role FOUND (Employee)")
                                # For all other inputs, save name and value of input
                                else:
                                    data[input['name']] = input['value']
                            # If input has no value, an error is caught but needs no handling
                            # since inputs without values do not need to be posted to next
                            # destination.
                            except:
                                pass

                    # If there's data in the dictionary, it is submitted to the destination url
                    if data:
                        response = session.post(url, data=data)
                    # If there's no data, just try to post to the destination without data
                    else:
                        response = session.post(url)
                    # If the url of the response is the Aula front page, loop is exited
                    if response.url == 'https://www.aula.dk:443/portal/':
                        success = True
            # If some error occurs, try to just ignore it
            except:
                pass
            # One is added to counter each time the loop runs independent of outcome
            counter += 1
        
        # Login succeeded without an HTTP error code and API requests can begin 
        if success == True and response.status_code == 200:
            self.logger.info("Login was successful")

            # All API requests go to the below url
            # Each request has a number of parameters, of which method is always included
            # Data is returned in JSON
            url = self.getAulaApiUrl()

            ### First API request. This request most be run to generate correct correct cookies for subsequent requests. ###
            params = {
                'method': 'profiles.getProfilesByLogin'
                }
            # Perform request, convert to json and print on screen
            response_profile = session.get(url, params=params).json()
            #print(json.dumps(response_profile, indent=4))

            self.setProfilesByLogin(response_profile)

            ### Second API request. This request most be run to generate correct correct cookies for subsequent requests. ###
            params = {
                'method': 'profiles.getProfileContext',
                'portalrole': 'employee', #Should be employee or guardian
            }
            # Perform request, convert to json and print on screen
            response_profile_context = session.get(url, params=params).json()
            #print(json.dumps(response_profile_context, indent=4))

            # Loop to get institutions and children associated with profile and save
            # them to lists
            institutions = []
            institution_profiles = []
            children = []

            for institution in response_profile_context['data']['institutions']:
                institutions.append(institution['institutionCode'])
                institution_profiles.append(institution['institutionProfileId'])
                for child in institution['children']:
                    children.append(child['id'])

            children_and_institution_profiles = institution_profiles + children

            ### Third example API request, uses data collected from second request ###
            params = {
                'method': 'notifications.getNotificationsForActiveProfile',
                'activeChildrenIds[]': children,
                'activeInstitutionCodes[]': institutions
            }

            # Perform request, convert to json and print on screen
            #notifications_response = session.get(url, params=params).json()
            #print(json.dumps(notifications_response, indent=4))

            ### Fourth example API request, only succeeds when the third has been run before ###
            params = {
                'method': 'messaging.getThreads',
                'sortOn': 'date',
                'orderDirection': 'desc',
                'page': '0'
            }

            # Perform request, convert to json and print on screen
            #response_threads = session.get(url, params=params).json()
            #print(json.dumps(response_threads, indent=4))

            ### Fifth example. getAllPosts uses a combination of children and instituion profiles. ###
            params = {
                'method': 'posts.getAllPosts',
                'parent': 'profile',
                'index': "0",
                'institutionProfileIds[]': children_and_institution_profiles,
                'limit': '10'
            }

            # Perform request, convert to json and print on screen
            #response_threads = session.get(url, params=params).json()
            #print(json.dumps(response_threads, indent=4))

            ### Sixth example. Posting a calender event. ###
            params = (
                ('method', 'calendar.createSimpleEvent'),
            )

            # Manually setting the cookie "profile_change". This probably has to do with posting as a parent.
            session.cookies['profile_change'] = '2'

            # Csrfp-tokenis manually added to session headers.
            session.headers['csrfp-token'] = session.cookies['Csrfp-Token']

            #Setting information for response
            login_response.status = True
            login_response.username = username

            return login_response

        # Login failed for some unknown reason
        else:
            self.logger.critical("Log in was unsuccessful")

            #Setting information for response
            login_response.status = False
            login_response.username = username
            return login_response

    #FROM: https://medium.com/@jorlugaqui/how-to-strip-html-tags-from-a-string-in-python-7cb81a2bbf44
    def __remove_html_tags(self,text):
        """Remove html tags from a string"""
        import re
        clean = re.compile('<.*?>')
        return re.sub(clean, '', text)

    def getEvents(self, startDatetime, endDatetime):
       
        #Calculates the diffence between the dates.
        monthsDiff = abs((endDatetime.year - startDatetime.year)) * 12 + abs(endDatetime.month - startDatetime.month)
        #print(monthsDiff)

        #Makes sure that even if only one event in same month, the loop will be run
        if monthsDiff <= 0:
            monthsDiff = 1

        events = []
        self.logger.info("Reading AULA calendar")
        self.logger.info("Locating events in calendars")
        step = 0
        for months in range(monthsDiff):
            lookUp_begin = startDatetime + relativedelta(months=months)
            lookUp_end = startDatetime + relativedelta(months=months+1)

            #End date can not be later than end date specified.
            if lookUp_end >= endDatetime:
                lookUp_end = endDatetime

            #outlookevents_from_aula = self.icalmanager.readAulaCalendarEvents()
            startTimeFormattet = lookUp_begin.strftime("%Y-%m-%dT%H:%M:%ST+02:00")
            endTimeFormattet = lookUp_end.strftime("%Y-%m-%dT%H:%M:%ST+02:00")

            step = step +1
            self.logger.info("  (%i of %i) Events from %s to %s"%(step,monthsDiff, startTimeFormattet,endTimeFormattet))

            #Includes institution
            self.logger.info("      In AULA institution calendar")
            events = events + self.getEventsForInstitutions(self.getProfileId(),self.getProfileinstitutionCode(),startTimeFormattet,endTimeFormattet)

            #Gets own events
            self.logger.info("      In AULA personal calendar")
            events = events + self.getEventsByProfileIdsAndResourceIds(self.getProfileId(), startTimeFormattet, endTimeFormattet)

            #Seems to be god with a simple cooldown time here. 
            time.sleep(0.1)

        class appointmentitem(object):
            pass

        aula_events = {}
        self.logger.info("Reading current AULA calendar events:")
        index = 1
        for event in events:
            response = self.getEventById(event["id"])
            #print(response["data"])

            try:
                self.logger.info("     (%s/%s) Event %s with start date %s" %(str(index),str(len(events)),response["data"]["title"],response["data"]["startDateTime"]))
            except TypeError as e:
                self.logger.warning("Skipping due to error: %s" %(e))

            mAppointmentitem = appointmentitem()
            mAppointmentitem.subject = response["data"]["title"]
            mAppointmentitem.body = response["data"]["description"]["html"]
            mAppointmentitem.aula_id = response["data"]["id"]
            mAppointmentitem.start = response["data"]["startDateTime"]
            mAppointmentitem.end = response["data"]["endDateTime"]
            mAppointmentitem.location = response["data"]["primaryResourceText"] 

            description = response["data"]["description"]["html"]
            #Finds AULA-Url for event
            #m = re.search('(?P<url>https?://[^\s]+)', description)
            #if m:
            #    aula_calendar_url = m.group("url").replace(",","")

            #Find GAID in description
            #m1 = re.search('o2a_outlook_GlobalAppointmentID=\S*', description)
            m1 = re.search('o2a_outlook_GlobalAppointmentID=\S*', description)
            if m1:
                outlook_GlobalAppointmentID = m1.group(0)
                outlook_GlobalAppointmentID = outlook_GlobalAppointmentID.split("=")[1].strip()
                outlook_GlobalAppointmentID = self.__remove_html_tags(outlook_GlobalAppointmentID).strip()

            #FINDS LMT in description
            m2 = re.search('o2a_outlook_LastModificationTime=\S* \S*\S\S:\S\S', description)
            if m2:
                outlook_LastModificationTime = m2.group(0)
                outlook_LastModificationTime = outlook_LastModificationTime.split("=")[1].strip()
                outlook_LastModificationTime = self.__remove_html_tags(outlook_LastModificationTime).strip()

            #if both GAID and LMT exists then add item to dict. 
            if m1 and m2:
                isDuplicate = False 
                if outlook_GlobalAppointmentID in aula_events.keys():
                    pass
                    #If event is added to institution calendar and also in own. Duples will occur. This prevents events from beeing removed. 
                    #if response["data"]["hideInOwnCalendar"] == True or response["data"]["hideInOwnCalendar"] == False and response["data"]["addedToInstitutionCalendar"] == False or  response["data"]["addedToInstitutionCalendar"] == False: 
                    #    outlook_GlobalAppointmentID = str(index) + "_" + outlook_GlobalAppointmentID
                    #    isDuplicate = True


                aula_events[outlook_GlobalAppointmentID]={
                    "appointmentitem":mAppointmentitem,
                    "isDuplicate" : isDuplicate,
                    "outlook_GlobalAppointmentID":outlook_GlobalAppointmentID,
                    "outlook_LastModificationTime":outlook_LastModificationTime
                }

            index = index +1
                #print( aula_events[outlook_GlobalAppointmentID]["outlook_GlobalAppointmentID"])



        return aula_events

    def test_run(self):
        from setupmanager import SetupManager
        #Gets AULA password and username from keyring
        self.setupmanager = SetupManager()
        aula_usr = self.setupmanager.get_aula_username()
        aula_pwd = self.setupmanager.get_aula_password()
        
        self.login(aula_usr,aula_pwd)

        self.updateEvent(287982478,"222NyTitel1","Min seje beskrivelse","2021-10-03T10:10:00.0000+02:00","2021-10-03T11:00:00.0000+02:00",[],False,False,True,False)

        #events = self.getEvents(None, None)

        #for event in events:
           # print(event["outlook_GlobalAppointmentID"])

        #events = self.getEventsByProfileIdsAndResourceIds(self.getProfileId())

        #for event in events:
        #    self.getEventById(event["id"])


    
        
        #invites = []
        #invites.append(aulagmr.findRecipient("Jesper Qvist"))

       # self.createEvent("TEST BEGIVENHED","BESKRIVELSEN AF BEGIVENHEDEN","2021-05-19T20:00:00+02:00","2021-05-19T23:00:00+02:00",invites,False,False)

       


#aulagmr = AulaManager()
#aulagmr.test_run()