#from aulamanager import AulaMananger
from setupmanager import SetupManager
from eventmanager import EventManager as eventmanager
import datetime as dt
from datetime import timedelta
import logging
import sys, getopt
from dateutil.relativedelta import relativedelta
from contactschecker import ContactsChecker
from outlookmanager import OutlookManager
import traceback

from databasemanager import DatabaseManager


#
# LOGGER
#
# create logger with 'spam_application'
logger = logging.getLogger('O2A')
logger.setLevel(logging.DEBUG)
# create file handler which logs even debug messages
fh = logging.FileHandler('o2a.log')
fh.setLevel(logging.DEBUG)
# create console handler with a higher log level
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)
# create formatter and add it to the handlers
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)
# add the handlers to the logger
logger.addHandler(fh)
logger.addHandler(ch)
logger.info('O2A startet')
today = dt.datetime.today()

outlookmanager=OutlookManager()

def run_script(force_update_existing_events = False):
      print ("***********************************************************")
      print("                   OUTLOOK TO AULA")
      print(" Jesper Qvist, Kløver-Skolen & Ole Frandsen, Dybbøl-Skolen")
      print ("***********************************************************")
      #Startdate is today, enddate is today next year - Tenical limit from AULA.
      try:
        eman = eventmanager()
        #comp = eman.compare_calendars(today,today+relativedelta(days=+4)) #Start dato er nu altid dags dato :) 
        comp = eman.compare_calendars(dt.datetime(today.year,today.month,today.day,1,00,00,00),dt.datetime(today.year+1,7,1,00,00,00,00),force_update_existing_events)
        eman.update_aula_calendar(comp)
      except Exception as err:
        logger.critical(traceback.format_exc())
        outlookmanager.send_a_mail_program(traceback.format_exc())
      finally:
        pass

#The main function
def main(argv):
  forceupdate = False

  #If no argument is passed
  if len(sys.argv) <= 1:
      run_script(forceupdate)

  #If any argument is passed
  try:
    opts, args = getopt.getopt(argv,"hsgardfc",["setup","setupgui","help","run","database","force","check","database_recipient_reset"])
  except getopt.GetoptError:
    print('OPTIONS')
    print(' without parameter  : same as -r')
    print(' -s    --setup                         : Run setup in terminal')
    print(' -g    --setupgui                      : Run setup with GUI')
    print(' -r    --run                           : To run script')
    print(' -a    --database_recipient_reset      : Reset local recipient database. ')
    print(' -f    --force                         : Force update all existing events')
    print(' -c    --check                         : Check if people in "contacts_to_check.csv" is present in AULA')
    print(' -h    --help                          : To show help')
    sys.exit(2)
  for opt, arg in opts:
    if opt in ("-h", "--help"): 
      print('OPTIONS')
      print(' without parameter  : same as -r')
      print(' -s    --setup                         : Run setup in terminal')
      print(' -g    --setupgui                      : Run setup with GUI')
      print(' -r    --run                           : To run script')
      print(' -a    --database_recipient_reset      : Reset local recipient database. ')
      print(' -f    --force                         : Force update all existing events')
      print(' -c    --check                         : Check if people in "contacts_to_check.csv" is present in AULA')
      print(' -h    --help                          : To show help')
    elif opt in ("-a", "--database_recipient_reset"):
      dbmanger = DatabaseManager()
      dbmanger.create_recipients_table(reset_table=True)
    elif opt in ("-d", "--database"): 
      print("database tjek")
      print(str(arg))
      dbmanger = DatabaseManager()
      rlts = dbmanger.update_record("040000008200E000744C5B7101A82E0080000000090DAC029B16AD801000000000000000010000000E3046EA24B8B6B4ABF3012BD878F09B9","2342343")
      print(rlts)

    elif opt in ("-f", "--force"): 
      forceupdate = True
      logger.warning("Force update is set to: " + str(forceupdate))
    elif opt in ("-r", "--run"): 
      #forceupdate = True #TODO: Da det retter/afhjælper, at nogle kolleger af en eller anden grund ikke tilføjes første gang. Finde fejlen, og ret det i stedet. 
      logger.warning("Force update is set to: " + str(forceupdate))
      run_script(forceupdate)
    elif opt in ("-g", "--setupgui"):
      setupmgr = SetupManager()
      setupmgr.setup_menu_gui()
    elif opt in ("-s", "--setup"):
      setupmgr = SetupManager()
      setupmgr.do_setup()
    elif opt in ("-c", "--check"):
      mChecker = ContactsChecker()
      mChecker.searchForPeople()
    else:
      print('OPTIONS')
      print(' without parameter  : same as -r')
      print(' -s    --setup                         : Run setup in terminal')
      print(' -g    --setupgui                      : Run setup with GUI')
      print(' -r    --run                           : To run script')
      print(' -a    --database_recipient_reset      : Reset local recipient database. ')
      print(' -f    --force                         : Force update all existing events')
      print(' -c    --check                         : Check if people in "contacts_to_check.csv" is present in AULA')
      print(' -h    --help                          : To show help')


if __name__ == "__main__":
    main(sys.argv[1:])