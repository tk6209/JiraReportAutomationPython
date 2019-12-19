# Download the dailyDashboard info

import os
import glob
from selenium import webdriver

activeUser = os.getlogin()

#Files handling paths
finalPath = ("C:\\Users\\vteixeira\\Downloads\\")
initPath = glob.glob("C:/Users/vteixeira/Downloads/*.csv")

#Remove Old files
for f in initPath:
    os.remove(f)


#Creates FireFox Profile
pref = webdriver.FirefoxProfile()
pref.set_preference("browser.helperApps.neverAsk.saveToDisk","application/csv,text/csv")
pref.set_preference("browser.download.manager.showWenStarting", False)
pref.set_preference("browser.download.dir", '"C:/Users/" , activeUser , "/tempAutomation/"')
pref.set_preference("browser.download.folderList", 2)

#Invoke WebDriver
ie = webdriver.Firefox(firefox_profile=pref)

#Login Credentials activeUser + password
password = os.environ.get('cred_anyconn')

#Download routine
ie.set_page_load_timeout(20)
ie.get('https://clientsupport.worldlinesweden.com/sr/jira.issueviews:searchrequest-csv-current-fields/11177/SearchRequest-11177.csv')
username_button = ie.find_element_by_id("login-form-username")
username_button.send_keys(activeUser)
password_button = ie.find_element_by_id("login-form-password")
password_button.send_keys(password)
login_button = ie.find_element_by_id("login-form-submit")
login_button.click()
ie.quit()


#File Rename | r=root, d=directories, f = files
files = []
for r, d, f in os.walk(finalPath):
    for file in f:
        if '.csv' in file:
            files.append(os.path.join(r,file))

for f in files:
    os.rename(f,"C:\\Users\\vteixeira\\Downloads\\raw.csv")

#Call VBA script
os.system("C:/scripts/OnePage/run.bat")

#print ("Congratulations today's report has been successfully sent")
exit()


#Other donwload features
#"text/plain,application/pdf,application/msword,application/csv,text/csv,application/rtf,application/xml,text/xml,application/octet-stream,application/vnd.ms-excel,application/zip,text/txt,text/plain,application/pdf,application/x-pdf")
