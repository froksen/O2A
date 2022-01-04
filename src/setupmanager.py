import getpass
import keyring
import configparser
import win32com.client
import time
import sys

class SetupManager:
    def __init__(self):
        self.config = configparser.ConfigParser()
        self.__read_config_file()


    def do_setup(self):
        self.__show_welcome_screen()
        self.__aula_setup()

    def __show_welcome_screen(self):
        print()
        print()
        print()
        print()
        print("..:: This is the initial setup for O2A ::..")
        print()
        print("WHY: This wizard will ask you for information that is needed to make the script work.")
        print("WHAT IF: If you misspell or make other mistakes. Run this wizard again.")
        print("SECURITY: All passwords and usernames are stored in the keyring for your operation system.")
        print("")
        input("Press <ENTER> to continue")

    def __aula_setup(self):
        print("..:: Information for AULA ::..")
        print("The following information is used to operate and login to AULA. Please enter information for UNI-login")


        usr = self.__ask_for_username()
        passwd = self.__ask_for_password()

        print("")
        print("")
        print("")
        print("")
        print("Is the following correct?")
        print("UNI-username: " + str(usr))
        print("UNI-password: " + str(passwd)) 
        print("(All passwords are stored in the keyring for your operation system.)")
        print("")
        print("")
        print("")
        
        should_save = False
        while not should_save:
            reply = str(input('Do you want to store these information? (Y)es or (N)o: ')).lower().strip()

            try:
                if reply[:1] == 'y':
                    #self.create_outlook_categories()
                    should_save = True
                if reply[:1] == 'n':
                    should_save = False
            except IndexError:
                should_save = False

            if should_save == False:
                print("Process abortet, nothing saved or changed! Rerun this setup if, you want to retype username or password.")
                sys.exit()

        try:
            self.config.add_section("AULA")
        except configparser.DuplicateSectionError:
            pass #If section already exists, then skip

        self.config['AULA']['username'] = usr
        keyring.set_password("o2a", "aula_password", passwd)

        self.__write_config_file()

        print("Username and password stored!")

        print()
        print()
        self.yes_or_no("Do you want to create necessary categories in Outlook?")

        print("AULA setup completed!")

    def store_yes_or_no(self, question):
        while "the answer is invalid":
            reply = str(input(question+' (y/n): ')).lower().strip()

            try:
                if reply[:1] == 'y':
                    #self.create_outlook_categories()
                    return True
                if reply[:1] == 'n':
                    return False
            except IndexError:
                return False

    def yes_or_no(self, question):
        while "the answer is invalid":
            reply = str(input(question+' (y/n): ')).lower().strip()

            try:
                if reply[:1] == 'y':
                    self.create_outlook_categories()
                    return True
                if reply[:1] == 'n':
                    return False
            except IndexError:
                return False

    def create_task(self):
        import os

        os.system("schtasks /CREATE /F /TN MINOPGAVE /XML task_template.xml")


    def create_outlook_categories(self):
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns = outlook.GetNamespace("MAPI")

        print("Checking if Outlook has necessary categories")
        hasAula = False
        hasAULAInstitutionskalender = False
        for category in ns.Categories:
            if(category.name == "AULA"):
                hasAula = True

            if category.name == "AULA Institutionskalender":
                hasAULAInstitutionskalender = True

        if not hasAula:
            print("Missing category 'AULA'. Will be created")
            ns.Categories.Add("AULA")
            time.sleep(1) #needed because otherwise outlook can keep up.

        if not hasAULAInstitutionskalender:
            print("Missing category 'AULA Institutionskalender'. Will be created")
            ns.Categories.Add("AULA Institutionskalender")
            time.sleep(1) #needed because otherwise outlook can keep up.

        if hasAULAInstitutionskalender and hasAula:
            print("All necessary categories was found.")
            #print(category.name)
            #print(category.CategoryID)

    def get_aula_username(self):
        return self.config['AULA']['username']

    def get_aula_password(self):
        return keyring.get_password("o2a","aula_password")

    def __read_config_file(self):
        try:
            self.config.read('configuration.ini')
        except Exception:
            pass

    def __write_config_file(self):
        with open('configuration.ini', 'w') as configfile:
            self.config.write(configfile)

    def __ask_for_password(self):
        pwd = ""
        try:
            pwd = getpass.getpass(prompt='Password: ', stream=None)
        except Exception as e:
            print(e)

        return pwd

    def __ask_for_username(self):
        usr = ""
        try:
            usr = input ("Username ["+getpass.getuser()+"]:")  or getpass.getuser()
        except Exception as e:
            print(e)

        return usr

