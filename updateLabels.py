# Function updating QGIS labels. Copies old files, moves them into new folder with todays date, runs R Scripts that alter
# the downloaded files from iLandman.
# This makes R work

import shutil
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException
import os
from selenium.webdriver import ActionChains
os.environ['PYTHONHOME'] = r'C:\Users\Accounting\Anaconda3'
os.environ['PYTHONPATH'] = r'C:\Users\Accounting\Anaconda3\lib\site-packages'
os.environ['R_HOME'] = 'C:/Program Files/R/R-4.0.2'
os.environ['R_USER'] = r'C:\Users\Accounting\Anaconda3\lib\site-packages\rpy2'
import rpy2.robjects as ro
import time
import glob
from datetime import datetime

# Function that finds the element and waits for it to load before continuing


def updateLabels(webscrape, workingInterestFile, ownershipFile):
    if webscrape == True:

        # Set up
        user_email = r'pvankatwyk@gmail.com'
        user_password = r'A&8BYpG%&*xy#s!'
        path = r'\\WEM-MASTER\Working Projects\WEMU Leasing\Python Codes\Python Code\chromedriver.exe'
        driver = webdriver.Chrome(path)

        # Go to iLandman
        driver.get(r'https://www.p2central.com/ilandman/Project')

        def findElement(XPATH):
            # path = r'\\WEM-MASTER\Working Projects\WEMU Leasing\Python Codes\Python Code\chromedriver.exe'
            # driver = webdriver.Chrome(path)
            find = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH)))
            return find

        # Log in
        email = findElement(
            r'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div/div[1]/div/input')
        counter = 0
        while counter < 10:
            try:
                email.click()
                counter = 10
            except ElementNotInteractableException:
                counter += 0.1

        email.click()
        email.send_keys(user_email)
        password = findElement(
            r'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div/div[2]/div/div/input')
        password.click()
        password.send_keys(user_password)
        password.send_keys(Keys.RETURN)

        project = findElement(r'//*[@id="projectList"]/table/tbody/tr/td[1]/a')
        project.click()

        reportdrop = findElement(r'//*[@id="menu"]/ul/li[6]/span/a')
        reportdrop.click()
        tractreport = findElement(r'//*[@id="menu"]/ul/li[6]/div[1]/div/a[1]')
        tractreport.click()
        tractr = findElement(r'//*[@id="Tract_reports"]/a')
        tractr.click()
        time.sleep(1)
        tractowner = findElement(r'//*[@id="ui-id-2"]/div[contains(.,"Tract Ownership List")]/a')
        tractowner.click()
        # Get download ID
        target = findElement('//*[@id="filters"]/table/tbody/tr[11]/td[2]/div/a')
        target = str(target.get_attribute('id'))
        ids = target[len('data-picker-'):-len('-trigger')]
        selowner = findElement(f'//*[@id="data-picker-{ids}-trigger"]')
        selowner.click()
        bar = findElement(f'//*[@id="data-picker-{ids}-simple"]/div/input[2]')
        bar.click()
        fourx = findElement(r'//*[@id="4108947"]')
        fourx.click()
        bar.send_keys('WEM')
        wem = findElement(r'//*[@id="4081860"]')
        wem.click()
        bar.click()
        bar.clear()
        bar.send_keys('UT 1808')
        ut = findElement(r'//*[@id="4059790"]')
        ut.click()
        ok = findElement(r'/html/body/div[137]/div[3]/div/button[2]/span')
        ok.click()
        viewreport = findElement(r'//*[@id="save-form"]')
        viewreport.click()
        time.sleep(15)
        cls = findElement(r'/html/body/div[144]/div[3]/div/button')
        cls.click()

        # Rename last download to ownership
        todayDate = datetime.date(datetime.now())
        time.sleep(5)
        ownerDate = r'ownerLabels ' + str(todayDate) + '.xlsx'
        list_of_files = glob.glob(r'C:\Users\Accounting\Downloads\*')
        ownership = max(list_of_files, key=os.path.getctime)
        newownerName = r'C:\\Users\\Accounting\\Downloads\\' + str(ownerDate)
        os.rename(ownership, newownerName)

        # Contract Reports
        contreport = findElement(r'//*[@id="contract-reports-submenu-link"]')
        contreport.click()
        contreportbut = findElement(r'//*[@id="Contract_reports"]')
        contreportbut.click()
        time.sleep(4)
        workinginterest = findElement(r'//*[@id="ui-id-2"]/div[contains(.,"Working")]/a')
        workinginterest.click()
        target2 = findElement('//*[@id="filters"]/table/tbody/tr[1]/td[2]/div/a')
        target2 = str(target2.get_attribute('id'))
        ids2 = target2[len('data-picker-'):-len('-trigger')]
        showacres = findElement(f'//*[@id="data-picker-{ids2}-trigger"]')
        showacres.click()
        search = findElement(f'//*[@id="data-picker-{ids2}-simple"]/div/input[2]')
        search.click()
        search.send_keys('WEM')
        time.sleep(1)
        wemii = findElement(r'//*[@id="115153"]')
        wemii.click()
        wemi = findElement(r'//*[@id="90095"]')
        wemi.click()
        okay = findElement(r'/html/body/div[contains(.,"OK")]/div[3]/div/button[2]')
        okay.click()
        target3 = findElement('//*[@id="filters"]/table/tbody/tr[3]/td[2]/div/a')
        target3 = str(target3.get_attribute('id'))
        ids3 = target3[len('data-picker-'):-len('-trigger')]
        contowner = findElement(f'//*[@id="data-picker-{ids3}-trigger"]')
        contowner.click()
        search2 = findElement(f'//*[@id="data-picker-{ids3}-simple"]/div/input[2]')
        search2.click()
        search2.send_keys('WEM')
        time.sleep(1)
        #wem2 = findElement(r'//*[@id="115153"]')
        wem2 = findElement(f'//*[@id="data-picker-{ids3}-multiple-results"]/ul/li[1]')
        wem2.click()
        wem1 = findElement(f'//*[@id="data-picker-{ids3}-multiple-results"]/ul/li[2]')
        wem1.click()
        time.sleep(1)
        okkk = findElement(r'/html/body/div[97]/div[3]/div/button[2]')
        okkk.click()
        target4 = findElement('//*[@id="filters"]/table/tbody/tr[7]/td[2]/div/a')
        target4 = str(target4.get_attribute('id'))
        ids4 = target4[len('data-picker-'):-len('-trigger')]
        owntype = findElement(f'//*[@id="data-picker-{ids4}-trigger"]')
        owntype.click()
        own = findElement(r'//*[@id="Ownership"]')
        own.click()
        okk = findElement(r'/html/body/div[101]/div[3]/div/button[2]/span')
        okk.click()
        viewrep = findElement(r'//*[@id="save-form"]')
        viewrep.click()
        downloadclick = WebDriverWait(driver, 10000).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="ui-id-46"]/p/p/a')))
        time.sleep(10)

        # Rename last download to leasehold
        todayDate = datetime.date(datetime.now())
        time.sleep(2)
        leasingDate = r'leasingLabels ' + str(todayDate) + '.xlsx'
        list_of_files = glob.glob(r'C:\Users\Accounting\Downloads\*')
        leasehold = max(list_of_files, key=os.path.getctime)
        newleaseName = r'C:\\Users\\Accounting\\Downloads\\' + str(leasingDate)
        os.rename(leasehold, newleaseName)


        # MAKE A COPY OF THE PREVIOUS FILES AND STORE THEM IN A NEW FOLDER WITH THE DATE
        dir = r"//WEM-MASTER/Sensitive Data/WEM Uintah/Maps/QGIS/IMPORTANT EXCEL DOCS/QGIS/"
        todayDate = datetime.date(datetime.now())
        dateFolder = dir + str(todayDate) + ' Backup'
        # check if directory exists or not yet
        if not os.path.exists(dateFolder):
            os.makedirs(dateFolder)
        # move files into created directory
        file_pathLeasing = r"//WEM-MASTER/Sensitive Data/WEM Uintah/Maps/QGIS/IMPORTANT EXCEL DOCS/QGIS/" \
                           r"QGIS - WEM LEASE Crop List.csv"
        file_pathOwnership = r'//WEM-MASTER/Sensitive Data/WEM Uintah/Maps/QGIS/IMPORTANT EXCEL DOCS/QGIS/' \
                             r'QGIS - WEM, UT, 4X Import List.csv'
        shutil.copy(file_pathLeasing, dateFolder)
        shutil.copy(file_pathOwnership, dateFolder)
        # rename the files
        shutil.move(dateFolder + '/QGIS - WEM LEASE Crop List.csv',
                    dateFolder+'/QGIS - WEM LEASE Crop List (moved ' + str(todayDate) + ').csv')
        shutil.move(dateFolder + '/QGIS - WEM, UT, 4X Import List.csv',
                    dateFolder + '/QGIS - WEM, UT, 4X Import List (moved ' + str(todayDate) + ').csv')
        driver.quit()


    if webscrape == False:
        newleaseName = workingInterestFile
        newownerName = ownershipFile
    # RUN THE R SCRIPTS FIXING THE DATA FROM ILANDMAN TO QGIS
    test = r"\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\Code\leasingLabels.R"
    ro.r.source(test)
    ro.r["leasingLabels"](newleaseName)  # Have the input to the function be the file path
    fpOwnership = r"\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\Code\ownershipLabels.R"
    ro.r.source(fpOwnership)
    ro.r["ownershipLabels"](newownerName)


    print("Done!")


#updateLabels(True, '', '')