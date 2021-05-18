#from aulamanager import AulaMananger
from setupmanager import SetupManager
from eventmanager import EventManager as eventmanager
import datetime as dt
import logging
import sys, getopt

#
# LOGGER
#
# create logger with 'spam_application'
logger = logging.getLogger('O2A')
logger.setLevel(logging.INFO)
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

#The main function
def main(argv):

  #If no argument is passed
  if len(sys.argv) <= 1:
      eman = eventmanager()
      
      #Startdate is today, enddate is today next year - Tenical limit from AULA.
      comp = eman.compare_calendars(today,dt.datetime(today.year +1 ,today.month,today.day)) #Start dato er nu altid dags dato :) 
      eman.update_aula_calendar(comp)

  #If any argument is passed
  try:
    opts, args = getopt.getopt(argv,"hsr",["setup","help","run"])
  except getopt.GetoptError:
    print('OPTIONS')
    print(' without parameter  : same as -r')
    print(' -s --setup  : To setup script')
    print(' -r --run    : To run script')
    print(' -h --help   : To show help')
    sys.exit(2)
  for opt, arg in opts:
    if opt in ("-h", "--help"): 
      print('OPTIONS')
      print(' without parameter  : same as -r')
      print(' -s --setup  : To setup script')
      print(' -r --run    : To run script')
      print(' -h --help   : To show help')
    elif opt in ("-r", "--run"): 
      eman = eventmanager()
      
      #Startdate is today, enddate is today next year - Tenical limit from AULA.
      comp = eman.compare_calendars(today,dt.datetime(today.year +1 ,today.month,today.day)) #Start dato er nu altid dags dato :) 
      eman.update_aula_calendar(comp)
    elif opt in ("-s", "--setup"):
      setupmgr = SetupManager()
      setupmgr.do_setup()
    else:
      print('OPTIONS')
      print(' -s --setup  : To setup script')
      print(' -r --run    : To run script')
      print(' -h --help   : To show help')


if __name__ == "__main__":
    main(sys.argv[1:])