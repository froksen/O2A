import getpass
import keyring
import configparser

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
        
        try:
            self.config.add_section("AULA")
        except configparser.DuplicateSectionError:
            pass #If section already exists, then skip

        self.config['AULA']['username'] = usr
        keyring.set_password("o2a", "aula_password", passwd)

        self.__write_config_file()

        print("Username and password stored!")
        print("AULA setup completed!")

    def get_aula_username(self):
        return self.config['AULA']['username']

    def get_aula_password(self):
        return keyring.get_password("o2a",self.get_aula_username())

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

