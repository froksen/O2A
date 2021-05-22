import aulamanager
import datetime as dt
import keyring

evm = aulamanager.AulaManager()

events = {}
evm.setBrowser(evm.createBrowser(headless=False))


#Gets AULA password and username from keyring
aula_usr = keyring.get_password("o2a", "aula_username")
aula_pwd = keyring.get_password("o2a", "aula_password")

#Login to AULA
if not evm.loginToAula_use_unilogin(aula_usr,aula_pwd) == 0:
    print("FEJLEDE")

attendees = []

evm.createCalendarEvent(event_title="Titel på begivenhed",start_date="28/04/2021",end_date="28/04/2021",start_time="12:00",end_time="14:15",place="Stedet",description="beskrivelse",attendees=attendees,all_day_event=False,sensitivity=0)

