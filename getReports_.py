from pathlib import Path
from msedge.selenium_tools import Edge, EdgeOptions
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import os
import time
import shutil
import pymsteams


# -------- Directories --------
powerbi_dir = str(r"")
db_dir = str(r"")
hist_path = str(r"")
driver_path = str(r".\msedgedriver.exe")


# -------- Files URLs --------
loginPage = ""
urlDownload = [
]

# -------- File names --------
file_old = [
    "incident",
    "task",
    "tsp1_demand",
    "sc_task"
    
]

file_new = [
    "incident",
    "tasks_wsr",
    "SEs_InProgress",
    "general_sctasks"
]

# -------- Enable automatic download ---------
def enable_download(driver):
    driver.command_executor._commands["send_command"] = ("POST", '/session/$sessionId/chromium/send_command')
    params = {'cmd':'Page.setDownloadBehavior', 'params': {'behavior': 'allow', 'downloadPath': db_dir}}
    driver.execute("send_command", params)

# -------- Set Chrome services ---------
#def setting_edge_service():    
    #chrome_service = ChromeService('chromedriver')
    #edge_service = EdgeService('msedgedriver')
    #edge_service.creationflags = CREATE_NO_WINDOW
    #return edge_service;

# -------- Set Chrome options ---------
def setting_edge_options():
    edge_options = EdgeOptions()
    #edge_options.add_argument("--headless")
    #edge_options.add_argument('--no-sandbox')
    #edge_options.add_argument('--disable-gpu')
    #edge_options.add_argument("--window-size=1,1")
    #edge_options.add_argument("--user-data-dir=C:\\Users\\lucas.guedes\\AppData\\Local\\Microsoft\\Edge\\User Data\\Default")
    return edge_options;

# -------- Moves file to history directory and rename with the file datetime creation ---------
def moveToHist():

    for i in range(len(file_new)):
        pathDb = str(db_dir + "\\" + file_new[i] + ".xlsx")
        pathHistory = str(hist_path + "\\" + file_new[i] + ".xlsx")
        if os.path.exists(pathDb):
            shutil.move(pathDb, pathHistory)
            t = os.path.getctime(pathHistory)
            t_str = time.ctime(t)
            t_obj = time.strptime(t_str)
            form_t = time.strftime("%Y%m%d_%H%M%S", t_obj)
            os.rename(pathHistory, str(hist_path + "\\" + file_new[i] + "_" + form_t + ".xlsx"))


# -------- Copy file from DB folder to PowerBI/Metrics folder ---------
def copyToPbi():

    for i in range(len(file_new)):
        pathPbi = str(powerbi_dir + "\\" + file_new[i] + ".xlsx")
        pathDb = str(db_dir + "\\" + file_new[i] + ".xlsx")
        if os.path.exists(pathDb):
            shutil.copy2(pathDb, pathPbi)


# -------- Main ---------
if __name__ == '__main__':
    
    moveToHist()
    
    browser = Edge(executable_path=driver_path,options=setting_edge_options()) #service=setting_edge_service()) # Apply the Edge settings to session
    #browser.set_window_position(3840, 1080, windowHandle ='current')
    enable_download(browser)
    browser.get(loginPage)
    time.sleep(5)
    
    try:
        element = WebDriverWait(browser, 10).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#tilesHolder > div:nth-child(1) > div > div > div > div.table-cell.text-left.content")))
        element.click()
    except TimeoutException:
        print("No element found")

    for i in range(len(urlDownload)):

        browser.get(urlDownload[i])
        j = 1

        while not os.path.exists(str(Path(db_dir + "\\" + file_old[i] + ".xlsx"))):
            
            if j <= 10:
                time.sleep(j)
                j += 1
            
            else:
            
                # #Error message to Teams Channel AMS Automacao to notify about the fail
                webhook = ''
                myTeamsMessage = pymsteams.connectorcard(webhook) # Connectorcard object creation
                myTeamsMessage.color("#ed0000") 
                myTeamsMessage.title("Download de Arquivos PowerBI")
                myTeamsMessage.text(f"Arquivo<strong> {file_old[i]}.xlsx </strong>não encontrado.") # Message to be sent
                myTeamsMessage.send() # Send the message.
                break
        else:

            os.rename(str(Path(db_dir + "\\" + file_old[i] + ".xlsx")), str(Path(db_dir + "\\" + file_new[i] + ".xlsx")))

    copyToPbi()

    browser.quit()