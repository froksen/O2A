# Imports
from sys import getprofile
import requests                 # Perform http/https requests
from bs4 import BeautifulSoup   # Parse HTML pages
import json                     # Needed to print JSON API data
import logging

#THIS CODE IS LARGELY INSPIRED BY CODE FROM https://helmstedt.dk/2020/05/et-lille-kig-paa-aulas-api/

class AulaManager:
    def __init__(self):
        # Start requests session
        self.session = requests.Session()
        self.__profilesByLogin = ""
        
        #Sets logger
        self.logger = logging.getLogger('O2A')

    def setProfilesByLogin(self,profile):
        self.__profilesByLogin = profile

    def getProfilesByLogin(self):
        return self.__profilesByLogin

    def getProfileinstitutionCode(self):
        return self.getProfilesByLogin()['data']['profiles'][0]['institutionProfiles'][0]['institutionCode']

    def getProfileId(self):
        return self.getProfilesByLogin()['data']['profiles'][0]['institutionProfiles'][0]['id']

    def getSession(self):
        return self.session

    def getAulaApiUrl(self):
        return 'https://www.aula.dk/api/v11/'

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
        url = " https://www.aula.dk/api/v11/?method=search.findRecipients&text="+recipient+"&query="+recipient+"&id="+str(self.getProfileId())+"&typeahead=true&limit=100&scopeEmployeesToInstitution=false&fromModule=event&instCode="+str(self.getProfileinstitutionCode())+"&docTypes[]=Profile&docTypes[]=Group"
        
        response  = session.get(url, params=params).json()
        #response = session.get(url).json()
        print(json.dumps(response, indent=4))

        try:
            recipient_profileid = response["data"]["results"][0]["docId"] #Appearenly its docId and not profileId
            print(recipient_profileid)

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
        print(json.dumps(response, indent=4))

        if(response["status"]["message"] == "OK"):
            self.logger.info("Event was removed!")

        else:
            self.logger.warning("Event was not removed!")

    def createEvent(self, title, description, startDateTime, endDateTime, attendee_ids = [], allDay = False, isPrivate = False):
        #EventArray

        print("IDDD")
        print(attendee_ids)

        session = self.getSession()

        # All API requests go to the below url
        # Each request has a number of parameters, of which method is always included
        # Data is returned in JSON
        url = self.getAulaApiUrl()
        
        ### First example API request ###
        params = {
            'method': 'calendar.createSimpleEvent'
            }

        data = {
            'title': title,
            'description': description,
            'startDateTime': startDateTime, # 2021-05-18T14:30:00.0000+02:00
            'endDateTime': endDateTime, # '2021-05-18T15:00:00.0000+02:00'
            #'startDate': startDate, #'2021-05-17'
            #'endDate': endDate #'2021-05-17', #'2021-05-17'
            #'startTime': '12:00:19', 
            #'endTime': '12:30:19',
            'id': '',
            'institutionCode': self.getProfileinstitutionCode(),
            'creatorInstProfileId': self.getProfileId(),
            'type': 'event',
            'allDay': allDay,
            'private': isPrivate,
            'primaryResource': {},
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
            'addToInstitutionCalendar': False,
            'hideInOwnCalendar': False,
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
            self.logger.info("Event was created!")
        else:
            self.logger.warning("Event was not created!")

    def getProfile(self):
                # All API requests go to the below url
        # Each request has a number of parameters, of which method is always included
        # Data is returned in JSON

        session  = self.getSession()

        url = 'https://www.aula.dk/api/v11/'
        
        ### First example API request ###
        params = {
            'method': 'profiles.getProfilesByLogin'
            }
        # Perform request, convert to json and print on screen
        response_profile = session.get(url, params=params).json()
        print(json.dumps(response_profile, indent=4))

    def login(self, username, password):
        
        # User info
        user = {
            'username': username,
            'password': password
            }
        
        # Start requests session
        session = self.getSession()
            
        # Get login page
        url = 'https://login.aula.dk/auth/login.php?type=unilogin'
        response = self.session.get(url)
        
        # Login is handled by a loop where each page is first parsed by BeautifulSoup.
        # Then the destination of the form is saved as the next url to post to and all
        # inputs are collected with special cases for the username and password input.
        # Once the loop reaches the Aula front page the loop is exited. The loop has a
        # maximum number of iterations to avoid an infinite loop if something changes
        # with the Aula login.
        counter = 0
        success = False
        while success == False and counter < 10:
            try:
                # Parse response using BeautifulSoup
                soup = BeautifulSoup(response.text, "lxml")
                # Get destination of form element (assumes only one)
                url = soup.form['action']   
                
                # If form has a destination, inputs are collected and names and values
                # for posting to form destination are saved to a dictionary called data
                if url:
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
                                # Save password if input is a password field
                                elif input['name'] == 'password':
                                    data[input['name']] = user['password']
                                #Selects login type, as employee this is "MEDARBEJDER_EKSTERN"
                                elif input['name'] == 'selected-aktoer':
                                    data[input['name']] = "MEDARBEJDER_EKSTERN"
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
            self.logger.info("Log in was successful")


            # All API requests go to the below url
            # Each request has a number of parameters, of which method is always included
            # Data is returned in JSON
            url = 'https://www.aula.dk/api/v11/'

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

            return True

        # Login failed for some unknown reason
        else:
            self.logger.critical("Log in was unsuccessful")

            return False
