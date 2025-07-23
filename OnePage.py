# Download the dailyDashboard info

import os
import glob
import time
from selenium import webdriver
from selenium.webdriver.common.by import By

# Usuário ativo do sistema
activeUser = os.getlogin()

# Caminhos de arquivos
download_dir = os.path.join("C:\\Users", activeUser, "Downloads")
temp_dir = os.path.join("C:\\Users", activeUser, "tempAutomation")

# Remove arquivos .csv antigos da pasta de downloads
for f in glob.glob(os.path.join(download_dir, "*.csv")):
    os.remove(f)

# Configurações do perfil do Firefox
profile = webdriver.FirefoxProfile()
profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/csv,text/csv")
profile.set_preference("browser.download.manager.showWhenStarting", False)
profile.set_preference("browser.download.dir", temp_dir)
profile.set_preference("browser.download.folderList", 2)

# Inicia o Firefox com o perfil configurado
driver = webdriver.Firefox(firefox_profile=profile)

# Captura senha da variável de ambiente
password = os.environ.get('cred_anyconn')
if not password:
    raise ValueError("A variável de ambiente 'cred_anyconn' não está definida.")

# Baixa o relatório do JIRA
try:
    driver.set_page_load_timeout(20)
    driver.get("https://clientsupport.worldlinesweden.com/sr/jira.issueviews:searchrequest-csv-current-fields/11177/SearchRequest-11177.csv")

    driver.find_element(By.ID, "login-form-username").send_keys(activeUser)
    driver.find_element(By.ID, "login-form-password").send_keys(password)
    driver.find_element(By.ID, "login-form-submit").click()

    # Aguarda download (ajuste conforme necessário)
    time.sleep(5)

finally:
    driver.quit()

# Renomeia o novo arquivo CSV
csv_files = glob.glob(os.path.join(download_dir, "*.csv"))
if csv_files:
    os.rename(csv_files[0], os.path.join(download_dir, "raw.csv"))
else:
    raise FileNotFoundError("Nenhum arquivo CSV foi encontrado para renomear.")

# Executa macro do Excel via .bat
os.system("C:/scripts/OnePage/run.bat")

exit()
