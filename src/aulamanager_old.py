from selenium import webdriver
from selenium.webdriver.chrome.options import Options 
from selenium.webdriver.firefox.options import Options
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
import sys
import os
from webdriver_manager.chrome import ChromeDriverManager


class AulaManager:
    def __init__(self):
        print("Aula Manager Initialized")
        self.browser = None

        self.logger = logging.getLogger('O2A')


    def setBrowser(self, browser):
        self.browser = browser

    def closeBrowser(self):
        self.browser.quit()

    def removeCalendarEvent(self,event_url):
        wait = WebDriverWait(self.browser, 60)

        try:
            #Goto event
            self.browser.get(event_url)
            time.sleep(5)
            #Click on "Rediger"
            self.logger.debug("Clicking 'Rediger'")
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="modal-1"]/div[1]/div/div[2]/div[2]/div/div/button[1]/a')))
            self.browser.find_element_by_xpath('//*[@id="modal-1"]/div[1]/div/div[2]/div[2]/div/div/button[1]/a').click()

            #Click on "Slet" i dialog
            time.sleep(5)
            self.logger.debug("Waiting for dialog/delete-button")
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="drawer"]/div[2]/div/form/footer/div[1]/button[1]')))
            self.browser.execute_script("arguments[0].scrollIntoView();", self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/footer/div[1]/button[1]')) #Scrools to element
            removeBtn = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/footer/div[1]/button[1]')
            self.logger.debug("Clicking dialog/delete-button")
            ActionChains(self.browser).move_to_element(removeBtn).click().perform()

            time.sleep(5)

            #Confirm
            self.logger.debug("Clicking 'OK'")
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="modal-1"]/div[1]/div/footer/div/button')))
            self.browser.find_element_by_xpath('//*[@id="modal-1"]/div[1]/div/footer/div/button').click()
            wait.until(EC.invisibility_of_element((By.XPATH,'//*[@id="modal-1"]/div[1]/div/footer/div/button')))
        except:
            e = sys.exc_info()[0]
            self.logger.warning("Unable to remove event. Failed with error: %s" %(e))
            return -1

        self.logger.info("Event succesfully REMOVED")
        return 0
    
    def createCalendarEvent(self, event_title,start_date,end_date,start_time,end_time, place, description, attendees = [], all_day_event = False, sensitivity = 0):
        self.logger.info("Attempting to CREATE event:\n - Title: %s\n - Start_date: %s\n - End_date: %s\n - Start_time:%s" %(event_title,start_date,end_date,start_time))

        try:
            self.logger.info("Entering to AULA Calendar")

            #self.browser.get("https://www.aula.dk/portal/#/kalender")
            self.browser.get("https://www.aula.dk/portal/#/kalender/opretbegivenhed?parent=profile")
            time.sleep(5)

            wait = WebDriverWait(self.browser, 20)

            time.sleep(2)

            #Fills out the form
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="eventTitle"]')))

            #description_field = self.browser.find_element_by_xpath('//*[@id="tinymce"]/p')

            #Event title
            self.logger.info("Setting title")
            event_title_field = self.browser.find_element_by_xpath('//*[@id="eventTitle"]')
            ActionChains(self.browser).move_to_element(event_title_field).click().click().send_keys(event_title).perform()

            #Attenedes
            self.logger.info("Setting attendees")
            attendees_field = self.browser.find_element_by_xpath('//*[@id="eventInvitees"]/div/div/div[1]/input')

            ActionChains(self.browser).move_to_element(attendees_field).click().perform()
            for attendee in attendees:
                self.logger.info("Attempting to add %s" %(attendee))
                ActionChains(self.browser).send_keys(attendee).perform()
                time.sleep(10)
                ActionChains(self.browser).send_keys(Keys.ENTER).perform()
                #time.sleep(5)

                def check_exists_by_css(csspath):
                    try:
                        self.browser.find_element_by_css_selector(csspath)
                    except:
                        #print("NOT EXISTS")
                        return False
                    #print("EXISTS")
                    return True

                if check_exists_by_css('p.el-select-dropdown__empty') == True:
                    self.logger.info("Attendee was not found in AULA")
                    self.logger.debug("Removing string from inputfield")
                    for x in range(len(attendee)):
                        self.logger.debug("Pressing backspace")
                        ActionChains(self.browser).send_keys(Keys.BACK_SPACE).perform()
                else:
                    self.logger.info("Attendee was found in AULA")


            #Start date
            self.logger.info("Setting start date")
            start_date_field = self.browser.find_element_by_xpath('//*[@id="startDate"]')
            start_date_field.clear()
            ActionChains(self.browser).move_to_element(self.browser.find_element_by_xpath('//*[@id="startDate"]')).click().send_keys(start_date).send_keys(Keys.ENTER).perform()
            start_date_field.clear()
            ActionChains(self.browser).move_to_element(self.browser.find_element_by_xpath('//*[@id="startDate"]')).click().send_keys(start_date).send_keys(Keys.ENTER).perform()

            #End Date
            self.logger.info("Setting end date")
            end_date_field = self.browser.find_element_by_xpath('//*[@id="endDate"]')
            end_date_field.clear()
            ActionChains(self.browser).move_to_element(end_date_field).click().send_keys(end_date).send_keys(Keys.ENTER).perform()
            end_date_field.clear()
            ActionChains(self.browser).move_to_element(end_date_field).click().send_keys(end_date).send_keys(Keys.ENTER).perform()

            #Start time
            self.logger.info("Setting start time")
            #print("START ENDTIME: ",end_time)
            start_time_field = self.browser.find_element_by_xpath('//*[@id="startTime"]')
            ActionChains(self.browser).move_to_element(start_time_field).click().perform()

            def time_selector_move(element,key,distance):
                if distance <= 0:
                    return

                i=0

                ActionChains(self.browser).move_to_element(element).click().perform()
                actions = ActionChains(self.browser).move_to_element(element)

                while i < distance-1:
                    #action.send_keys(key)
                    actions.send_keys(key)
                    i += 1
                actions.perform()

            start_time_hours_field = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/div[1]/div/div[3]/div[1]/div/div[2]/span/div/div/ul[1]')
            time_selector_move(start_time_hours_field,Keys.ARROW_UP,24)
            time_selector_move(start_time_hours_field,Keys.ARROW_DOWN,int(start_time.split(":")[0])-1)
            ActionChains(self.browser).send_keys(Keys.ENTER).perform()

            start_time_minuts_field = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/div[1]/div/div[3]/div[1]/div/div[2]/span/div/div/ul[2]')
            time_selector_move(start_time_minuts_field,Keys.ARROW_UP,24)
            five_minuts_steps =  round(float(int(start_time.split(":")[1])/5)) 
            time_selector_move(start_time_minuts_field,Keys.ARROW_DOWN,five_minuts_steps)
            ActionChains(self.browser).send_keys(Keys.ENTER).perform()


            time.sleep(2)
            ActionChains(self.browser).move_to_element(event_title_field).click().click().perform()
            time.sleep(2)

            #End time
            end_time_field = self.browser.find_element_by_xpath('//*[@id="endTime"]')

            
            self.browser.execute_script("arguments[0].scrollIntoView();", end_time_field) #Scrools to element


            ActionChains(self.browser).move_to_element(end_time_field).click().perform() #Sets focus

            #print("SETTING ENDTIME: ",end_time)
            self.logger.info("Setting end time")
            #Sets hours
            end_time_hours_field = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/div[1]/div/div[4]/div[1]/div/div[2]/span/div/div/ul[1]')
            time_selector_move(end_time_hours_field,Keys.ARROW_UP,24)
            time_selector_move(end_time_hours_field,Keys.ARROW_DOWN,int(end_time.split(":")[0])-1)
            #time.sleep(2)
            ActionChains(self.browser).send_keys(Keys.ENTER).perform()

            #Sets minuts
            end_time_minuts_field = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/div[1]/div/div[4]/div[1]/div/div[2]/span/div/div/ul[2]')
            time_selector_move(end_time_minuts_field,Keys.ARROW_UP,24)
            five_minuts_steps =  round(float(int(end_time.split(":")[1])/5)) 
            time_selector_move(end_time_minuts_field,Keys.ARROW_DOWN,five_minuts_steps)
            #time.sleep(2)
            ActionChains(self.browser).send_keys(Keys.ENTER).perform()

            time.sleep(2)
            ActionChains(self.browser).move_to_element(event_title_field).click().click().perform()
            time.sleep(2)

            #If AllDay event
            if all_day_event == True:
                allDay_field = self.browser.find_element_by_xpath('//*[@id="allDay"]')
                self.browser.execute_script("arguments[0].scrollIntoView();", allDay_field) #Scrools to element
                ActionChains(self.browser).move_to_element(allDay_field).click().perform()    


            #If event is private (2 = private)
            if sensitivity == 2:
                private_field = self.browser.find_element_by_xpath('//*[@id="isPrivate"]')
                self.browser.execute_script("arguments[0].scrollIntoView();", private_field) #Scrools to element
                ActionChains(self.browser).move_to_element(private_field).click().perform()  

            #Place
            place_field = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/div[1]/div/div[6]/div[1]/div/div[1]/div/div/div[1]/input')
            self.browser.execute_script("arguments[0].scrollIntoView();", place_field) #Scrools to element
            ActionChains(self.browser).move_to_element(place_field).click().send_keys(place).perform()

            #Descriptionsfield
            self.logger.info("Setting description")
            self.browser.switch_to.frame(self.browser.find_element_by_id("editorInCalendar_ifr"))
            elem = self.browser.find_element_by_xpath('//*[@id="tinymce"]/p')
            description_action = ActionChains(self.browser).move_to_element(elem).click()

            description_lines = description.split("\n")
            for line in description_lines:
                #description_action.send_keys_to_element(elem,line)
                elem.send_keys(line)
                elem.send_keys(Keys.ENTER)
                #description_action.send_keys_to_element(elem,Keys.ENTER)
                print(line)
            
            #description_action.perform()
            #elem.send_keys(description)
            self.browser.switch_to.default_content()
            
            #Event title
            self.logger.info("Pressing the OK button")
            self.logger.debug("Waiting")

            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="drawer"]/div[2]/div/form/footer/div[1]/button[2]')))
            self.logger.debug("finding element")
            create_btn = self.browser.find_element_by_xpath('//*[@id="drawer"]/div[2]/div/form/footer/div[1]/button[2]')
            self.logger.debug("Scolling to element")
            self.browser.execute_script("arguments[0].scrollIntoView();", create_btn) #Scrools to element

            #time.sleep(5)
            ActionChains(self.browser).move_to_element(create_btn).click().perform()

            # If attendees is otherwise occupied, alert box will show. Press OK to this. If alert box does not show, then continue. 
            try:
                time.sleep(2)
                attendees_alert_dialog_ok_button = self.browser.find_element_by_xpath('//*[@id="modal-1"]/div[1]/div/footer/div/button')
                ActionChains(self.browser).move_to_element(attendees_alert_dialog_ok_button).click().perform()
            except:
                pass #if element does not exist. Just pass

            self.logger.debug("Clicking Element")

            wait.until(EC.invisibility_of_element((By.XPATH,'//*[@id="drawer"]/div[2]/div/form/footer/div[1]/button[2]')))

        except:
            e = sys.exc_info()[0]
            self.logger.warning("Event was UNSUCCESSFULLY created. Failed with error: %s" %(e))
            return -1

        self.logger.info("Event was SUCCESSFULLY created")

    def loginToAula_use_unilogin(self, username, password):
        self.logger.info("Attempting to log into AULA")

        if not isinstance(username,str) or not isinstance(password,str):
            self.logger.warning("Username or Password is not a string. Attempting to continue. Most likely to fail. Missing running setup script?")

        browser = self.browser 

        self.logger.debug("Going to https://www.aula.dk/portal/#/login")
        browser.get('https://www.aula.dk/portal/#/login') 

        
        wait = WebDriverWait(browser, 20)

        try:
            #Button for "UNILOGIN options"
            self.logger.info("Login step 1 of 7")
            self.logger.debug("Clicking 'UNILOGIN' option")
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="main"]/div[1]/div[4]/div[1]/div')))
            browser.find_element_by_xpath('//*[@id="main"]/div[1]/div[4]/div[1]/div').click()
            #self.browser.execute_script("arguments[0].scrollIntoView();", otherOptionsBtn) #Scrools to element
            time.sleep(1)

            #INPUTfield for Username
            self.logger.info("Login step 2 of 7")
            self.logger.debug("Entering USERNAME")
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="username"]')))
            usrfield = browser.find_element_by_xpath('//*[@id="username"]')
            #self.browser.execute_script("arguments[0].scrollIntoView();", municipalidpBtn) #Scrools to element
            ActionChains(browser).move_to_element(usrfield).click().send_keys(username).perform()
            time.sleep(1)

            #Clicking Næste
            self.logger.info("Login step 3 of 7")
            self.logger.debug("Clicking 'Næste' option")
            wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/main/div/div/form/nav/button')))
            browser.find_element_by_xpath('/html/body/main/div/div/form/nav/button').click()
            time.sleep(1)

            #INPUTfield for Password
            self.logger.info("Login step 4 of 7")
            self.logger.debug("Entering PASSWORD")
            wait.until(EC.visibility_of_element_located((By.XPATH,'//*[@id="form-error"]')))
            pwd_input_field = browser.find_element_by_xpath('//*[@id="form-error"]')
            ActionChains(browser).move_to_element(pwd_input_field).click().send_keys(password).perform()
    
            #Clicking Login
            self.logger.info("Login step 5 of 7")
            self.logger.debug("Clicking 'Log ind' btn")
            wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/main/div/div/form/nav/div/div[1]/button')))
            browser.find_element_by_xpath('/html/body/main/div/div/form/nav/div/div[1]/button').click()

            #Clicking Employee-button
            self.logger.info("Login step 6 of 7")
            self.logger.debug("Selecting Employee-option")
            
            wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR,"a[onclick=\"selectAktoer('MEDARBEJDER_EKSTERN');return false;\"]")))
            browser.find_element_by_css_selector("a[onclick=\"selectAktoer('MEDARBEJDER_EKSTERN');return false;\"]").click()
            #wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/main/div/div/nav/a[2]')))
            #browser.find_element_by_xpath('/html/body/main/div/div/nav/a[2]').click()

            time.sleep(20)
            self.logger.info("Login step 7 of 7")
            self.logger.info("Clicking OK to Cookie-information")
            wait.until(EC.visibility_of_element_located((By.XPATH,'/html/body/div[1]/div[4]/div[1]/div[2]/button')))
            cookieBtn = self.browser.find_element_by_xpath('/html/body/div[1]/div[4]/div[1]/div[2]/button')
            self.browser.execute_script("arguments[0].scrollIntoView();", cookieBtn) #Scrools to element
            cookieBtn.click()
        except:
            e = sys.exc_info()[0]
            self.logger.critical("UNSUCCESFULLY logged into AULA. Error: %s" %(e))
            return -1

        self.logger.info("SUCCESFULLY logged into AULA")
        return 0

    def createBrowser(self, headless = False):
        chrome_options = webdriver.ChromeOptions()
        #chrome_options.add_argument("--start-maximized")
        firefox_options = Options()
        
        if(headless):
            chrome_options.add_argument("--headless");
            firefox_options.headless = True

        #chromedriver_autoinstaller.install()  # Check if the current version of chromedriver exists
                                            # and if it doesn't exist, download it automatically,
                                            # then add chromedriver to path


        #print("WORKING DIR")
        #print(os.getcwd())
        browser = webdriver.Chrome(executable_path=ChromeDriverManager().install(),options=chrome_options)
        #browser = webdriver.Chrome(executable_path=r"./webdrivers/chromedriver.exe",options=chrome_options)
        #browser.minimize_window()
        #browser = webdriver.Firefox(executable_path=r"geckodriver.exe",options=firefox_options)
        browser.set_window_size(1024, 768)
        return browser

