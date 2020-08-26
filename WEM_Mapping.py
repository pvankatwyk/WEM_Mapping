# Import Modules
import PySimpleGUI as sg
import pandas as pd
import numpy as np
import os
import re
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
from selenium.webdriver import ActionChains
from datetime import datetime
import shutil
import glob
from playsound import playsound
import threading
# from MBExtract import MBExtract
# from updateLabels import updateLabels
# from MBUltimate import MBUltimate
# from mbDegreesImport import mbDegreesImport
# from mbDegrees import mbDegrees
# from iLandmanTract import iLandmanTract
# from iLandmanMapTract import iLandmanTractMap
# from soundfxn import sound
os.environ['PYTHONHOME'] = r'C:\Users\Accounting\Anaconda3'
os.environ['PYTHONPATH'] = r'C:\Users\Accounting\Anaconda3\lib\site-packages'
os.environ['R_HOME'] = 'C:/Program Files/R/R-4.0.2'
os.environ['R_USER'] = r'C:\Users\Accounting\Anaconda3\lib\site-packages\rpy2'
import rpy2.robjects as ro
# TODO: Add Comments!!!
# TODO: Update the user manual

def sound(name):
    from playsound import playsound
    # FXN: Plays a sound
    # INPUT: Name of the file in Sounds folder (see Filepath) -- "beep"
    name = name+'.mp3'
    song = r"\\WEM-MASTER\Working Projects\WEMU Leasing\Python Codes\Python Code\Sounds" + "\\" + name
    playsound(song)

def iLandmanTract(Township, Range, Section, TractNumber, County, GrossAcres, Description, MB, Map):
    # format inputs to match naming convention in iLM (4S 3W to 04S-03W)
    if len(Township) == 3:
        Township = Township
    elif len(Township) == 2:
        Township = '0' + str(Township)

    if len(Range) == 3:
        Range = Range
    elif len(Range) == 2:
        Range = '0' + str(Range)

    if len(Section) == 2:
        Section = Section
    elif len(Section) == 1:
        Section = '0' + str(Section)

    if len(TractNumber) == 3:
        TractNumber = TractNumber
    elif len(TractNumber) == 2:
        TractNumber = '0' + str(TractNumber)
    elif len(TractNumber) == 1:
        TractNumber = '00' + str(TractNumber)


    # Set up
    user_email = r'pvankatwyk@gmail.com'
    user_password = r'A&8BYpG%&*xy#s!'
    # Establish path to the Chrome driver (allows you to access Google Chrome)
    path = r'\\WEM-MASTER\Working Projects\WEMU Leasing\Python Codes\Python Code\chromedriver.exe'
    #path = r'C:\Users\Peter\Downloads\chromedriver_win32 (1)\chromedriver.exe'
    driver = webdriver.Chrome(path)

    # Go to iLandman
    driver.get(r'https://www.p2central.com/ilandman/Project')

    # Function that finds the element and waits for it to load before continuing
    def findElement(XPATH):
        find = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH)))
        return find

    # Log in
    email = findElement(r'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div/div[1]/div/input')
    counter = 0
    while counter < 10:
        try:
            email.click()
            counter = 10
        except ElementNotInteractableException:
            counter += 0.1

    email.click()
    # Enter in USN and PASS
    email.send_keys(user_email)
    password = findElement(r'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div/div[2]/div/div/input')
    password.click()
    password.send_keys(user_password)
    password.send_keys(Keys.RETURN)


    # Get to Add Tract
    project = findElement(r'//*[@id="projectList"]/table/tbody/tr/td[1]/a')
    project.click()
    dropDown = findElement(r'//*[@id="menu"]/ul/li[3]/span/a')
    dropDown.click()
    addTract = findElement(r'//*[@id="menu"]/ul/li[3]/div[1]/div/a[2]')
    addTract.click()

    # Add Tract First Page
    TractName = 'U-' + Township + '-' + Range + '-' + Section + '.' + TractNumber
    number = findElement(r'//*[@id="Tract_Number"]')
    number.click()
    time.sleep(1)
    number.send_keys(TractName)

    ctyDropDown = findElement(r'//*[@id="Tract_LocationCountyParish"]')
    ctyDropDown.click()

    if County in ('Uintah', 'uintah', 'Uinta', 'uinta'):
        ctyDropDown.send_keys("u")
    if County in ('Duchesne', 'duchesne', 'duschene', 'Duschene'):
        ctyDropDown.send_keys("d")

    gross = findElement(r'//*[@id="Tract_Acres"]')
    gross.click()
    gross.send_keys(GrossAcres)

    shortDes = findElement(r'//*[@id="Tract_ShortDesc"]')
    shortDes.click()

    # Regex for description
    # Recognizes the actionable items within the legal description based on the pattern below
    pattern = re.compile(r'(?<=:)(.*)')
    matches = pattern.finditer(Description)
    matches1 = ''
    # compiles the matches into a list
    for match in matches:
        matches1 += match.group()

    # Reformat the info into a short description
    ShortDescription = ''
    if MB == False:
        ShortDescription = TractName[0:-4] + ':' + str(matches1) + ', ' + str(GrossAcres) + 'ac'
    if MB == True:
        ShortDescription = TractName[0:-4] + ':' + ' M&B' + ', ' + str(GrossAcres) + 'ac'

    shortDes.send_keys(ShortDescription)

    savefirst = findElement(r'//*[@id="save-form"]')
    savefirst.click()

    # Add Tract Second Page
    legal = findElement(r'//*[@id="Tract_Description"]')

    # Format Legal Description
    TownshipDir = ''
    RangeDir = ''
    if Township[2] == 'S':
        TownshipDir = 'South'
    elif Township[2] == 'N':
        TownshipDir = 'North'
    if Range[2] == 'E':
        RangeDir = 'East'
    elif Range[2] == 'W':
        RangeDir = 'West'

    legalTR = 'Township ' + Township[1] + ' ' + TownshipDir + ', ' + 'Range ' + Range[1] \
              + ' ' + RangeDir + ', U.S.M.' + '\n' + Description
    legal.click()
    legal.send_keys(legalTR)

    # Add location
    addlocation = findElement(r'//*[@id="locations-content"]/div[1]/span/a[1]')
    addlocation.click()

    sectionLocation = findElement(r'//*[@id="Section"]')
    sectionLocation.click()
    sectionLocation.send_keys(Section)

    townshipLocation = findElement(r'//*[@id="Township"]')
    townshipLocation.click()
    townshipLocation.send_keys(Township)

    rangeLocation = findElement(r'//*[@id="Range"]')
    rangeLocation.click()
    rangeLocation.send_keys(Range)

    # Tractulator capabilities for those that aren't M&B
    if MB == False:
        editLoc = findElement(r'//*[@id="Description"]/input[2]')
        editLoc.click()

        # Recognizes each item needed for tractulator input
        # For more information on what the Regular Expression code means and to practice, go to "Regex101.com"
        #regex for description elements
        pattern = re.compile(r'\s[NSEW][/1-9A-Z]{1,11}')
        matches = pattern.finditer(Description)
        matchlist = []
        for match in matches:
            matchlist.append(match.group())

        DesEl1 = matchlist[0]
        try:
            DesEl2 = matchlist[1]
        except IndexError:
            DesEl2 = ''
        try:
            DesEl3 = matchlist[2]
        except IndexError:
            DesEl3 = ''
        try:
            DesEl4 = matchlist[3]
        except IndexError:
            DesEl4 = ''
        try:
            DesEl5 = matchlist[4]
        except IndexError:
            DesEl5 = ''
        try:
            DesEl6 = matchlist[5]
        except IndexError:
            DesEl6 = ''

        north = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[1]/td[2]/div')
        south = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[3]/td[2]/div')
        east = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[2]/td[3]/div')
        west = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[2]/td[1]/div')
        quarter = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[1]/td[3]/div')
        half = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[1]/td[1]/div')
        andbutton = findElement(r'//*[@id="tractCalculatorContainer"]/div/table/tbody/tr[3]/td[3]/div')

        # Press tractulator buttons based on the list above (when you see North, click North)
        for i in range(len(DesEl1)):
            if DesEl1[i] == 'N':
                north.click()
            elif DesEl1[i] == 'S':
                south.click()
            elif DesEl1[i] == 'E':
                east.click()
            elif DesEl1[i] == 'W':
                west.click()
            elif DesEl1[i] == '/':
                continue
            elif DesEl1[i] == '2':
                half.click()
            elif DesEl1[i] == '4':
                quarter.click()
        try:
            if DesEl2 != '':
                andbutton.click()
                for i in range(len(DesEl2)):
                    if DesEl2[i] == 'N':
                        north.click()
                    elif DesEl2[i] == 'S':
                        south.click()
                    elif DesEl2[i] == 'E':
                        east.click()
                    elif DesEl2[i] == 'W':
                        west.click()
                    elif DesEl2[i] == '/':
                        continue
                    elif DesEl2[i] == '2':
                        half.click()
                    elif DesEl2[i] == '4':
                        quarter.click()
        except IndexError:
            pass

        try:
            if DesEl3 != '':
                andbutton.click()
                for i in range(len(DesEl3)):
                    if DesEl3[i] == 'N':
                        north.click()
                    elif DesEl3[i] == 'S':
                        south.click()
                    elif DesEl3[i] == 'E':
                        east.click()
                    elif DesEl3[i] == 'W':
                        west.click()
                    elif DesEl3[i] == '/':
                        continue
                    elif DesEl3[i] == '2':
                        half.click()
                    elif DesEl3[i] == '4':
                        quarter.click()
        except IndexError:
            pass

        try:
            if DesEl4 != '':
                andbutton.click()
                for i in range(len(DesEl4)):
                    if DesEl4[i] == 'N':
                        north.click()
                    elif DesEl4[i] == 'S':
                        south.click()
                    elif DesEl4[i] == 'E':
                        east.click()
                    elif DesEl4[i] == 'W':
                        west.click()
                    elif DesEl4[i] == '/':
                        continue
                    elif DesEl4[i] == '2':
                        half.click()
                    elif DesEl4[i] == '4':
                        quarter.click()
        except IndexError:
            pass

        try:
            if DesEl5 != '':
                andbutton.click()
                for i in range(len(DesEl5)):
                    if DesEl5[i] == 'N':
                        north.click()
                    elif DesEl5[i] == 'S':
                        south.click()
                    elif DesEl5[i] == 'E':
                        east.click()
                    elif DesEl5[i] == 'W':
                        west.click()
                    elif DesEl5[i] == '/':
                        continue
                    elif DesEl5[i] == '2':
                        half.click()
                    elif DesEl5[i] == '4':
                        quarter.click()
        except IndexError:
            pass
        try:
            if DesEl6 != '':
                andbutton.click()
                for i in range(len(DesEl6)):
                    if DesEl6[i] == 'N':
                        north.click()
                    elif DesEl6[i] == 'S':
                        south.click()
                    elif DesEl6[i] == 'E':
                        east.click()
                    elif DesEl6[i] == 'W':
                        west.click()
                    elif DesEl6[i] == '/':
                        continue
                    elif DesEl6[i] == '2':
                        half.click()
                    elif DesEl6[i] == '4':
                        quarter.click()
        except IndexError:
            pass

        time.sleep(2)
        setdes = findElement(r'/html/body/div[contains(.,"Set Description")]/div[3]/div/button[1]/span')
        setdes.click()

    time.sleep(1)
    saveloc = findElement(r'/html/body/div[contains(.,"Save")]/div[3]/div/button[2]')
    saveloc.click()

    # Scroll down
    driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")

    saveTract = findElement(r'//*[@id="save-form"]')
    saveTract.click()

    if MB == False and Map == True:
        addTractMap = findElement(r'//*[@id="location-addtract"]')
        addTractMap.click()
        add = add = findElement(r'/html/body/div[72]/div[3]/div/button[1]/span')
        add.click()

    time.sleep(5)
    driver.quit()

    output = "Done!"
    return output

def iLandmanTractMap(Township, Range, Section, TractNumber, County, Corner, Description):

    # Set up
    user_email = r'pvankatwyk@gmail.com'
    user_password = r'A&8BYpG%&*xy#s!'
    path = r'\\WEM-MASTER\Working Projects\WEMU Leasing\Python Codes\Python Code\chromedriver.exe'
    driver = webdriver.Chrome(path)

    # Go to iLandman
    driver.get(r'https://www.p2central.com/ilandman/Project')

    # Function that finds the element and waits for it to load before continuing
    def findElement(XPATH):
        find = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, XPATH)))
        return find

    # Log in
    email = findElement(r'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div/div[1]/div/input')
    counter = 0
    while counter < 10:
        try:
            email.click()
            counter = 10
        except ElementNotInteractableException:
            counter += 0.1

    email.click()
    email.send_keys(user_email)
    password = findElement(r'//*[@id="auth0-lock-container-1"]/div/div[2]/form/div/div/div[2]/span/div/div/div/div/div/div/div/div/div/div[2]/div/div/input')
    password.click()
    password.send_keys(user_password)
    password.send_keys(Keys.RETURN)

    # Get to map
    project = findElement(r'//*[@id="projectList"]/table/tbody/tr/td[1]/a')
    project.click()

    mapDrop = findElement(r'//*[@id="menu"]/ul/li[7]/span/a')
    mapDrop.click()
    projmap = findElement(r'//*[@id="menu"]/ul/li[7]/div[1]/div/a[1]')
    projmap.click()
    driver.switch_to.window(driver.window_handles[1])
    closeIntro = findElement(r'/html/body/div[7]/div[1]/button/span[1]')
    time.sleep(2)
    closeIntro.click()

    actionChains = ActionChains(driver)
    fullmap = findElement(r'//*[@id="container"]')
    actionChains.context_click(fullmap).perform()

    editor = findElement(r'//*[@id="piemenuCanvas"]')
    action = webdriver.common.action_chains.ActionChains(driver)
    action.move_to_element_with_offset(editor, 5, 100).perform()
    action.click()
    action.perform()


    sound('beep')
    dropdownTract = WebDriverWait(driver, 3600).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="layerEditorSelectLayer"]/i')))
    dropdownTract.click()

    tractMB = findElement(r'//*[@id="ui-id-6"]')
    tractMB.click()

    time.sleep(3)
    MBdraw = findElement(r'//*[@id="layerEditorMetesBounds"]/i')
    MBdraw.click()

    state = findElement(r'//*[@id="metesBoundsStates"]')
    state.click()
    state.send_keys('u')
    state.send_keys(Keys.RETURN)
    time.sleep(3)

    county = findElement(r'//*[@id="metesBoundsCounties"]')
    county.click()
    if County in ('Uintah', 'uintah', 'Uinta', 'uinta'):
        county.send_keys("u")
    if County in ('Duchesne', 'duchesne', 'duschene', 'Duschene'):
        county.send_keys("duch")


    county.send_keys(Keys.RETURN)

    tshp = findElement(r'//*[@id="metesBoundsBoundaryTierCountainer"]/span[1]/select')
    tshp.click()
    time.sleep(0.75)
    tshp.send_keys(Township)

    rng = findElement(r'//*[@id="metesBoundsBoundaryTierCountainer"]/span[2]/select')
    rng.click()
    time.sleep(0.75)
    rng.send_keys(Range)

    sec = findElement(r'//*[@id="metesBoundsBoundaryTierCountainer"]/span[3]/select')
    sec.click()
    time.sleep(0.75)
    sec.send_keys(Section)

    crn = findElement(r'//*[@id="metesBoundsCorner"]')
    crn.click()
    time.sleep(0.75)
    crn.send_keys(Corner)
    crn.send_keys(Keys.RETURN)

    sound('beep')
    firstcoord = WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="metesBoundsCoordinatesContainer"]/div')))
    firstcoord.click()

    # Get actions
    pattern = re.compile(
        r'.hence.(North|South|East|West|N|S|E|W|north|south|east|west).\d{1,4}.{0,2}\d{0,2}.{0,2}\d{0,2}.{0,3}(North|South|East|West|N|S|E|W|north|south|east|west){0,5}.{0,1}\d{0,4}.{0,1}\d{0,2}.(feet|rods|chains|Feet|ft|Rods)')
    matches = pattern.finditer(Description) # TODO: Is this reading the degrees minutes AND seconds? -- I THINK SO.... TEST
    matches1 = ''
    matchlist = []
    for match in matches:
        matches1 += '\n' + match.group()
        matchlist.append(match.group())

    match_split = []
    for item in matchlist:
        # match_split.append((re.split(r"[\s°º’′'\"][\s]{0,1}", item))) # THIS INCLUDES " - don't know how to get it to add space after
        match_split.append((re.split(r"[\s°º’′'\"]", item)))
        # match_split.append((re.split(r"[\s]", item)))

    for i in range(len(match_split)):
        if len(match_split[i]) < 5 and (
                match_split[i][1] == 'North' or match_split[i][1] == 'N' or match_split[i][1] == 'n' or match_split[i][
            1] == 'north' or match_split[i][1] == 'South' or match_split[i][1] == 's' or match_split[i][1] == 'south' or
                match_split[i][1] == 'S'):
            match_split[i].insert(2, '0')
            match_split[i].insert(3, '0')
            match_split[i].insert(4, '0')
            match_split[i].insert(5, '0')
        elif len(match_split[i]) < 5 and (
                match_split[i][1] == 'East' or match_split[i][1] == 'E' or match_split[i][1] == 'e' or match_split[i][
            1] == 'east' or match_split[i][1] == 'West' or match_split[i][1] == 'W' or match_split[i][1] == 'west' or
                match_split[i][1] == 'w'):
            match_split[i].insert(1, '0')
            match_split[i].insert(2, '0')
            match_split[i].insert(3, '0')
            match_split[i].insert(4, '0')

    for list in match_split:
        if len(list) == 9:
            del list[5]
        if len(list) == 7:
            list.insert(4, '')

    for list in match_split:
        if list[7] in ('rods', 'Rods', 'rod', 'Rod'):
            list[6] = str(float(list[6]) * 16.5)
            list[7] = 'feet'

    for list in match_split:
        try:
            list.remove('thence')
        except ValueError:
            list.remove('Thence')
        try:
            list.remove('feet')
        except ValueError:
            list.remove('ft')

    northing = findElement(r'//*[@id="metesBoundsVectorUpDownDirection"]')
    deg = findElement(r'//*[@id="metesBoundsVectorDegrees"]')
    easting = findElement(r'//*[@id="metesBoundsVectorLeftRightDirection"]')
    ft = findElement(r'//*[@id="metesBoundsVectorMagnitude"]')
    add = findElement(r'//*[@id="metesBoundsAddVector"]')

    for list in match_split:
        if list[0] in ('North', 'north', 'N', 'n'):
            northing.click()
            northing.send_keys('n')
        elif list[0] in ('South', 's', 'south', 'S'):
            northing.click()
            northing.send_keys('s')

        if list[0] in (0, '0'):
            deg.click()
            deg.send_keys('90')
        else:
            deg.click()
            deg.send_keys(list[1], ' ', list[2], ' ', list[3])

        if list[4] in ('East', 'east', 'E', 'e'):
            easting.click()
            easting.send_keys('e')
        elif list[4] in ('West', 'w', 'west', 'W'):
            easting.click()
            easting.send_keys('w')
        ft.click()
        ft.send_keys(list[5])
        add.click()

    addpoly = findElement(r'/html/body/div[contains(.,"Add Polygon")]/div[3]/div/button[1]/span')
    addpoly.click()
    close = findElement(r'/html/body/div[13]/div[1]/button/span[1]')
    close.click()

    time.sleep(2)
    select = findElement(r'//*[@id="layerEditorSelect"]/i')
    select.click()
    time.sleep(3)
    # attributes = findElement(r'//*[@id="layerEditorAttributes"]')

    # From stack exchange when ElementClickInterceptedException.. idk why this works
    sound('beep')
    time.sleep(10)
    button = driver.find_element_by_xpath(r'//*[@id="layerEditorAttributes"]')
    driver.execute_script("arguments[0].click();", button)


    # format inputs
    if len(Township) == 3:
        Township = Township
    elif len(Township) == 2:
        Township = '0' + str(Township)

    if len(Range) == 3:
        Range = Range
    elif len(Range) == 2:
        Range = '0' + str(Range)

    if len(Section) == 2:
        Section = Section
    elif len(Section) == 1:
        Section = '0' + str(Section)

    if len(TractNumber) == 3:
        TractNumber = TractNumber
    elif len(TractNumber) == 2:
        TractNumber = '0' + str(TractNumber)
    elif len(TractNumber) == 1:
        TractNumber = '00' + str(TractNumber)


    TractName = 'U-' + Township + '-' + Range + '-' + Section + '.' + TractNumber

    time.sleep(3)
    att = findElement(r'//*[@id="AttributeEditorTractNumber"]')
    att.click()
    att.send_keys(TractName)

    savename = findElement(r'/html/body/div[contains(.,"Save")]/div[3]/div/button[1]/span')
    savename.click()
    time.sleep(1)
    savetomap = findElement(r'//*[@id="layerEditorCommit"]/i')
    savetomap.click()
    savechanges = findElement(r'/html/body/div[contains(.,"Save all changes")]/div[3]/div/button[1]/span')
    savechanges.click()

    time.sleep(100)

    output = "Done!"
    return output

def updateLabels(webscrape, workingInterestFile, ownershipFile):
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
    time.sleep(2)
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
    workinginterest = findElement(r'//*[@id="ui-id-2"]/div[27]/a')
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

    driver.quit()
    print("Done!")

def MBUltimate(MBstring):
    # FXN: Takes a legal description string, separates it into actionable steps, and exports the results into CSV to be uploaded into QGIS
    # INPUT: Legal Description (string)

    # Identify pattern that will find actionable steps
    pattern = re.compile(r'.hence.(North|South|East|West|N|S|E|W|north|south|east|west).\d{1,4}.{0,2}\d{0,2}.{0,2}\d{0,2}.{0,2}(North|South|East|West|N|S|E|W|north|south|east|west){0,5}.{0,1}\d{0,4}.{0,1}\d{0,2}.(feet|rods|chains|Feet|ft|Rods)')
    matches = pattern.finditer(MBstring)
    matches1 = ''
    matchlist = []
    # Compile the matches into a list
    for match in matches:
        matches1 += '\n' + match.group()
        matchlist.append(match.group())

    match_split = []

    # Split the matches on white space (\s), or degrees minutes seconds notation
    for item in matchlist:
        match_split.append((re.split(r"[\s°’']", item)))

    # For actions which only have 2 components (such as West 100 feet), created empty strings to show there are no degrees minutes seconds
    for i in range(len(match_split)):
        if len(match_split[i]) < 5 and (match_split[i][1] == 'North' or match_split[i][1] =='N' or match_split[i][1] =='n' or match_split[i][1] =='north' or match_split[i][1] =='South' or match_split[i][1] =='s' or match_split[i][1] =='south' or match_split[i][1] =='S'):
            match_split[i].insert(2,'')
            match_split[i].insert(3, '')
            match_split[i].insert(4, '')
            match_split[i].insert(5, '')
        elif len(match_split[i]) < 5 and (match_split[i][1] == 'East' or match_split[i][1] =='E' or match_split[i][1] =='e' or match_split[i][1] =='east' or match_split[i][1] =='West' or match_split[i][1] =='W' or match_split[i][1] =='west' or match_split[i][1] =='w'):
            match_split[i].insert(1, '')
            match_split[i].insert(2, '')
            match_split[i].insert(3, '')
            match_split[i].insert(4, '')

    # Make the list of lists a data frame (makes it easier to work with)
    df_matches = pd.DataFrame(match_split)
    # Name the columns
    df_matches.columns = ['thence', 'northing', 'degree', 'minute', 'second', 'easting', 'distance', 'unit']
    # Drop the column with all of the 'thence' in it
    df_final = df_matches.drop(columns = ['thence'])

    # Convert all of the numbers into numeric values in order to do calculations
    df_final['degree'] = pd.to_numeric(df_final['degree'], errors = 'coerce')
    df_final['minute'] = pd.to_numeric(df_final['minute'], errors = 'coerce')
    df_final['second'] = pd.to_numeric(df_final['second'], errors = 'coerce')
    df_final['distance'] = pd.to_numeric(df_final['distance'], errors = 'coerce')

    # Fill in 0 in all of the spaces with "NA"
    df_final.fillna(0, inplace = True)

    # Conversion calculation into new 'degrees' column (not working with directionality)
    df_final['degrees'] = df_final['degree'] + (df_final['minute'] / 60) + (df_final['second'] / 3600)

    # Get new columns ready (pre-allocate them and make sure they're floats)
    df_final['decimalDegrees'] = 0
    df_final['decimalDegrees'] = pd.to_numeric(df_final['decimalDegrees'], errors = 'coerce')
    df_final['decimalDegrees'] = df_final['decimalDegrees'].astype(float)
    df_final['distanceConverted'] = 0
    df_final['distanceConverted'] = df_final['distanceConverted'].astype(float)

    # Encorporate directionality into conversion above
    #   - basically, if North, add 0 / South, add 180 to the calculation. Then, if East then add degrees (clockwise), west subtract (counterclockwise)
    for i in range(len(df_final)):
        if ((df_final['northing'][i] == "N") or (df_final['northing'][i] == 'n') or (df_final['northing'][i] == 'North') or (df_final['northing'][i] == 'north')) and ((df_final['easting'][i] == 'East') or (df_final['easting'][i] == 'east') or (df_final['easting'][i] == 'E') or (df_final['easting'][i] == 'e')):
            df_final['decimalDegrees'][i] = 0 + df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'N') or (df_final['northing'][i] == 'n') or (df_final['northing'][i] == 'North') or (df_final['northing'][i] == 'north')) and ((df_final['easting'][i] == 'W') or (df_final['easting'][i] == 'w') or (df_final['easting'][i] == 'West') or (df_final['easting'][i] == 'west')):
            df_final['decimalDegrees'][i] = 360 - df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'S') or (df_final['northing'][i] == 's') or (df_final['northing'][i] == 'South') or (df_final['northing'][i] == 'south')) and ((df_final['easting'][i] == 'E') or (df_final['easting'][i] == 'e') or (df_final['easting'][i] == 'East') or (df_final['easting'][i] == 'east')):
            df_final['decimalDegrees'][i] = 180 - df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'S') or (df_final['northing'][i] == 's') or (df_final['northing'][i] == 'South') or (df_final['northing'][i] == 'south')) and ((df_final['easting'][i] == 'W') or (df_final['easting'][i] == 'w') or (df_final['easting'][i] == 'West') or (df_final['easting'][i] == 'west')):
            df_final['decimalDegrees'][i] = 180 + df_final['degrees'][i]
        elif ((df_final['northing'][i] == 'S') or (df_final['northing'][i] == 's') or (df_final['northing'][i] == 'South') or (df_final['northing'][i] == 'south')) and df_final['easting'][i] == '':
            df_final['decimalDegrees'][i] = 180
        elif (df_final['northing'][i] == '') and ((df_final['easting'][i] == 'W') or (df_final['easting'][i] == 'w') or (df_final['easting'][i] == 'West') or (df_final['easting'][i] == 'west')):
            df_final['decimalDegrees'][i] = 270
        elif (df_final['northing'][i] == '') and ((df_final['easting'][i] == 'E') or (df_final['easting'][i] == 'e') or (df_final['easting'][i] == 'East') or (df_final['easting'][i] == 'east')):
            df_final['decimalDegrees'][i] = 90

    # Convert units into feet
    for i in range(len(df_final)):
        if df_final['unit'][i] in ('feet', 'ft', 'Feet', 'ft.', 'Ft'):
            df_final['distanceConverted'][i] = df_final['distance'][i]
        elif df_final['unit'][i] in ('rods', 'Rods'):
            df_final['distanceConverted'][i] = df_final['distance'][i] * 16.5
        elif df_final['unit'][i] == 'chains':
            df_final['distanceConverted'][i] = df_final['distance'][i] * 66

    # Copy to new dataframe the columns we need
    df_decimal = df_final[['decimalDegrees','distanceConverted']].copy()

    # Export as a CSV in a format that the M&B mapper on QGIS will recognize it
    pd.DataFrame.to_csv(self = df_decimal, path_or_buf=r'C:\Users\Accounting\Downloads\MBUltimate.csv',
                                                encoding='utf-8', index=False, sep=';', header = False)
    return matches1

def MBExtract(string):
    pattern = re.compile(r'.hence.(North|South|East|West|N|S|E|W|north|south|east|west).\d{1,4}.{0,2}\d{0,2}.{0,2}\d{0,2}.{0,2}(North|South|East|West|N|S|E|W|north|south|east|west){0,5}.{0,1}\d{0,4}.{0,1}\d{0,2}.(feet|rods|chains|Feet|ft|Rods)')
    matches = pattern.finditer(string)
    matches1 = ''
    matchlist = []
    for match in matches:
        matches1 += '\n' + match.group()
        matchlist.append(match.group())
    return matches1

def mbDegrees(cDirection1, dDegrees, dMinutes, dSeconds, cDirection2, feet, unit):
    # FXN: Convert input to decimal degrees
    # INPUTS: Legal description actionable item

    # Make sure inputs are floats
    dDegrees = float(dDegrees)
    dMinutes = float(dMinutes)
    dSeconds = float(dSeconds)
    feet = float(feet)

    # Convert degrees
    degrees = dDegrees + (dMinutes / 60) + (dSeconds / 3600)
    decimalDegrees = 1
    # Configure the decimal degrees according to directionality
    if (cDirection1 == 'N' or cDirection1 == 'n') and (cDirection2 == 'E' or cDirection2 == 'e'):
        decimalDegrees = 0 + degrees
    elif (cDirection1 == 'N' or cDirection1 == 'n') and (cDirection2 == 'W' or cDirection2 == 'w'):
        decimalDegrees = 360 - degrees
    elif (cDirection1 == 'S' or cDirection1 == 's') and (cDirection2 == 'E' or cDirection2 == 'e'):
        decimalDegrees = 180 - degrees
    elif (cDirection1 == 'S' or cDirection1 == 's') and (cDirection2 == 'W' or cDirection2 == 'w'):
        decimalDegrees = 180 + degrees

    # Convert units into feet
    if unit == 'feet':
        feet = feet
    elif unit == 'rods':
        feet = feet * 16.5
    elif unit == 'chains':
        feet = feet * 66

    output = str(decimalDegrees) + ' degrees ' + str(feet) + ' feet'

    # Format some error messages (to catch wrong inputs)
    if (cDirection1 != "N") and (cDirection1 != "n") and (cDirection1 != "S") and (cDirection1 != "S"):
        output = "The Northing value has to be either 'N' or 'S'."
    if (cDirection2 != "E") and (cDirection2 != "e") and (cDirection2 != "W") and (cDirection2 != "w"):
        output = "The Easting value has to be either 'E' or 'W'."

    return output

def mbDegreesImport(cDirection1, dDegrees, dMinutes, dSeconds, cDirection2, feet):
    # FXN: Several singular conversions of Degrees/Minutes/Seconds to Decimal Degrees with export
    # INPUTS: Legal description actionable item

    # Format the inputs and outputs
    if (cDirection1 == '') and (dDegrees == '') and (dMinutes == '') and (dSeconds == '') and (cDirection2 == '') and (feet == ''):
        output = [0,0]
    elif (cDirection1 != "N") and (cDirection1 != "n") and (cDirection1 != "S") and (cDirection1 != "S"):
        output = "The Northing value has to be either North or South."
    elif (cDirection2 != "E") and (cDirection2 != "e") and (cDirection2 != "W") and (cDirection2 != "w"):
        output = "The Easting value has to be either East or West."
    elif (cDirection1 == '') and (dDegrees == '') and (dMinutes == '') and (dSeconds == '') and (cDirection2 == '') and (feet == ''):
        output = np.array((0,0))
    else:
        # Convert input to decimal degrees
        dDegrees = float(dDegrees)
        dMinutes = float(dMinutes)
        dSeconds = float(dSeconds)
        feet = float(feet)

        degrees = dDegrees + (dMinutes / 60) + (dSeconds / 3600)
        decimalDegrees = 1
        # Configure the decimal degrees according to directionality
        if (cDirection1 == 'N' or cDirection1 == 'n') and (cDirection2 == 'E' or cDirection2 == 'e'):
            decimalDegrees = 0 + degrees
        elif (cDirection1 == 'N' or cDirection1 == 'n') and (cDirection2 == 'W' or cDirection2 == 'w'):
            decimalDegrees = 360 - degrees
        elif (cDirection1 == 'S' or cDirection1 == 's') and (cDirection2 == 'E' or cDirection2 == 'e'):
            decimalDegrees = 180 - degrees
        elif (cDirection1 == 'S' or cDirection1 == 's') and (cDirection2 == 'W' or cDirection2 == 'w'):
            decimalDegrees = 180 + degrees

        output = np.array((decimalDegrees, feet))


    return output









sg.theme('Dark2')

# Main Window
layout_Main = [[sg.Image(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\Code\WEM (2).png')],
               [sg.Text('What would you like to view?')],
               [sg.Button('Programs'), sg.Button('Files')]]

win_Main = sg.Window("WEM Mapping Application", element_justification = 'c').Layout(layout_Main)
winActive_Programs = False
winActive_Files = False
winActive_Labels = False
winActive_MB = False
winActive_Layer = False
winActive_MBImport = False
winActive_MBExtract = False
winActive_MBUltimate = False
winActive_AddTract = False
threadMessage = "Program Started"
while True:
    ev_Main, vals_Main = win_Main.Read()
    if ev_Main is None or ev_Main == 'Exit':
        break
    # Programs Window
    if not winActive_Programs and ev_Main == 'Programs':
        winActive_Programs = True
        win_Main.Hide()
        layout_Programs = [[sg.Text('Which Program?')],
                           [sg.Text('Map Updates')],
                           [sg.Button('QGIS Layer Updates')],
                           [sg.Button('QGIS Label Updates')],
                           [sg.Text('Metes and Bounds Calculators')],
                           [sg.Button('Decimal Degrees M&B Calculator')],
                           [sg.Button('M&B String Extractor')],
                           [sg.Button('QGIS M&B Import/Calculator')],
                           [sg.Button('The Ultimate M&B Program')],
                           [sg.Text('iLandman')],
                           [sg.Button('Add Tract')]]
        win_Programs = sg.Window('Programs', element_justification = 'l').Layout(layout_Programs)
        while True:
            ev_Programs, vals_Programs = win_Programs.Read()
            if ev_Programs is None or ev_Programs == 'Exit':
                winActive_Programs = False
                win_Programs.Close()
                win_Main.UnHide()
                break

            # QGIS Label Updates Window
            if not winActive_Labels and ev_Programs == 'QGIS Label Updates':
                winActive_Labels = True
                win_Programs.Hide()
                layout_Labels = [[sg.Text('Please enter the Working Interest Detail and Ownership Labels')],
                                 [sg.Checkbox("Webscrape these files?", key = '-scrape-')],
                                 [sg.Text("Working Interest Detail: "),sg.FileBrowse('Browse', target = '-wi-'), sg.InputText(key='-wi-', size=(65,1))],
                                 [sg.Text("Ownership Labels: "),sg.FileBrowse('Browse', target = '-or-'), sg.InputText(key='-or-', size=(65,1))],
                                 [sg.Text('', key='-output-', size=(10, 1))],
                                 [sg.Submit()]]
                win_Labels = sg.Window('QGIS Label Updates').Layout(layout_Labels)
                while True:
                    ev_Labels, vals_Labels = win_Labels.Read()
                    if ev_Labels is None or ev_Labels == 'Exit':
                        winActive_Labels = False
                        win_Labels.Close()
                        win_Programs.UnHide()
                        break
                    if ev_Labels == 'Submit':
                        labelupdate = threading.Thread(target = updateLabels, args = (vals_Labels['-scrape-'], vals_Labels['-wi-'], vals_Labels['-or-']), daemon=True)
                        labelupdate.start()
                        win_Labels.FindElement('-output-').Update(threadMessage)

            if not winActive_Layer and ev_Programs == 'QGIS Layer Updates':
                winActive_Layer = True
                win_Programs.Hide()
                layout_Layer = [[sg.Text('The program to update QGIS Ownership and Leasehold \n'
                                         'has to be run within QGIS. Open up the Python Console \n'
                                         'and type in the command "updateMapLayers()". If this \n'
                                         'doesn\'t work, you can run the script, save it, and try \n'
                                         'again. Then, copy and paste the styles from the old \n'
                                         'layers to the new layers. \n'
                                         'More info in the manual below:')],
                                [sg.Button('Open Job Manual')]]
                win_Layer = sg.Window('QGIS Layer Updates').Layout(layout_Layer)
                while True:
                    ev_Layer, vals_Layer = win_Layer.Read()
                    if ev_Layer is None or ev_Layer == 'Exit':
                        winActive_Layer = False
                        win_Layer.Close()
                        win_Programs.UnHide()
                        break
                    if ev_Layer == "Open Job Manual":
                        os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\MapsJOB MANUAL.docx')


            # Metes and Bounds Calculator
            if not winActive_MB and ev_Programs == 'Decimal Degrees M&B Calculator':
                winActive_MB = True
                win_Programs.Hide()
                layout_MB = [[sg.Text('Please enter the Metes and Bounds from the description: ')],
                             [sg.Text("Enter Here: "), sg.Combo(['N','S'], key = 'northing', size = [3,1]), sg.InputText(key = 'degrees', size = [5,1]),
                              sg.InputText(key = 'minutes', size = [5,1]),sg.InputText(key = 'seconds', size = [5,1]),sg.Combo(['E','W'], key = 'easting', size = [3,1]),
                              sg.InputText(key = 'feet', size = [6,1]), sg.Combo(['feet', 'rods', 'chains'], key = 'unit')],
                             [sg.Text('', key = '-output-', size = (50,1))],
                             [sg.Submit()]]
                win_MB = sg.Window('Decimal Degrees M&B Calculator').Layout(layout_MB)
                while True:
                    ev_MB, vals_MB = win_MB.Read()
                    if ev_MB is None or ev_Programs == 'Exit':
                        winActive_MB = False
                        win_MB.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MB == 'Submit':
                        output = mbDegrees(vals_MB['northing'], vals_MB['degrees'], vals_MB['minutes'], vals_MB['seconds'], vals_MB['easting'], vals_MB['feet'], vals_MB['unit'])
                        win_MB.FindElement('-output-').Update(output)

            # QGIS M&B Import/Calculator
            if not winActive_MBImport and ev_Programs == 'QGIS M&B Import/Calculator':
                winActive_MBImport = True
                win_Programs.Hide()
                layout_MBImport = [[sg.Text('Please enter the Metes and Bounds from the description: \n'
                                            'Make sure you enter in the right Northing/Eastings (if not, it will crash)')],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing1', size=[3, 1]), sg.InputText(key='degrees1', size=[5, 1]), sg.InputText(key='minutes1', size=[5, 1]),
                                    sg.InputText(key='seconds1', size=[5, 1]),sg.Combo(['E', 'W'], key='easting1', size=[3, 1]), sg.InputText(key='feet1', size=[6, 1]),sg.Text('', key='-output-1', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing2', size=[3, 1]), sg.InputText(key='degrees2', size=[5, 1]), sg.InputText(key='minutes2', size=[5, 1]),
                                    sg.InputText(key='seconds2', size=[5, 1]), sg.Combo(['E', 'W'], key='easting2', size=[3, 1]),sg.InputText(key='feet2', size=[6, 1]), sg.Text('', key='-output-2', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing3', size=[3, 1]),sg.InputText(key='degrees3', size=[5, 1]), sg.InputText(key='minutes3', size=[5, 1]),
                                    sg.InputText(key='seconds3', size=[5, 1]),sg.Combo(['E', 'W'], key='easting3', size=[3, 1]),sg.InputText(key='feet3', size=[6, 1]), sg.Text('', key='-output-3', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing4', size=[3, 1]),sg.InputText(key='degrees4', size=[5, 1]), sg.InputText(key='minutes4', size=[5, 1]),
                                    sg.InputText(key='seconds4', size=[5, 1]),sg.Combo(['E', 'W'], key='easting4', size=[3, 1]),sg.InputText(key='feet4', size=[6, 1]), sg.Text('', key='-output-4', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing5', size=[3, 1]),sg.InputText(key='degrees5', size=[5, 1]), sg.InputText(key='minutes5', size=[5, 1]),
                                    sg.InputText(key='seconds5', size=[5, 1]),sg.Combo(['E', 'W'], key='easting5', size=[3, 1]),sg.InputText(key='feet5', size=[6, 1]), sg.Text('', key='-output-5', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing6', size=[3, 1]),sg.InputText(key='degrees6', size=[5, 1]), sg.InputText(key='minutes6', size=[5, 1]),
                                    sg.InputText(key='seconds6', size=[5, 1]),sg.Combo(['E', 'W'], key='easting6', size=[3, 1]), sg.InputText(key='feet6', size=[6, 1]), sg.Text('', key='-output-6', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing7', size=[3, 1]),sg.InputText(key='degrees7', size=[5, 1]), sg.InputText(key='minutes7', size=[5, 1]),
                                    sg.InputText(key='seconds7', size=[5, 1]),sg.Combo(['E', 'W'], key='easting7', size=[3, 1]),sg.InputText(key='feet7', size=[6, 1]), sg.Text('', key='-output-7', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing8', size=[3, 1]),sg.InputText(key='degrees8', size=[5, 1]), sg.InputText(key='minutes8', size=[5, 1]),
                                    sg.InputText(key='seconds8', size=[5, 1]),sg.Combo(['E', 'W'], key='easting8', size=[3, 1]),sg.InputText(key='feet8', size=[6, 1]), sg.Text('', key='-output-8', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing9', size=[3, 1]),sg.InputText(key='degrees9', size=[5, 1]), sg.InputText(key='minutes9', size=[5, 1]),
                                    sg.InputText(key='seconds9', size=[5, 1]),sg.Combo(['E', 'W'], key='easting9', size=[3, 1]),sg.InputText(key='feet9', size=[6, 1]), sg.Text('', key='-output-9', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing10', size=[3, 1]),sg.InputText(key='degrees10', size=[5, 1]), sg.InputText(key='minutes10', size=[5, 1]),
                                    sg.InputText(key='seconds10', size=[5, 1]),sg.Combo(['E', 'W'], key='easting10', size=[3, 1]),sg.InputText(key='feet10', size=[6, 1]), sg.Text('', key='-output-10', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing11', size=[3, 1]),sg.InputText(key='degrees11', size=[5, 1]), sg.InputText(key='minutes11', size=[5, 1]),
                                    sg.InputText(key='seconds11', size=[5, 1]),sg.Combo(['E', 'W'], key='easting11', size=[3, 1]),sg.InputText(key='feet11', size=[6, 1]), sg.Text('', key='-output-11', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing12', size=[3, 1]),sg.InputText(key='degrees12', size=[5, 1]), sg.InputText(key='minutes12', size=[5, 1]),
                                    sg.InputText(key='seconds12', size=[5, 1]),sg.Combo(['E', 'W'], key='easting12', size=[3, 1]),sg.InputText(key='feet12', size=[6, 1]), sg.Text('', key='-output-12', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing13', size=[3, 1]),sg.InputText(key='degrees13', size=[5, 1]), sg.InputText(key='minutes13', size=[5, 1]),
                                    sg.InputText(key='seconds13', size=[5, 1]),sg.Combo(['E', 'W'], key='easting13', size=[3, 1]),sg.InputText(key='feet13', size=[6, 1]), sg.Text('', key='-output-13', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing14', size=[3, 1]),sg.InputText(key='degrees14', size=[5, 1]), sg.InputText(key='minutes14', size=[5, 1]),
                                    sg.InputText(key='seconds14', size=[5, 1]),sg.Combo(['E', 'W'], key='easting14', size=[3, 1]),sg.InputText(key='feet14', size=[6, 1]), sg.Text('', key='-output-14', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing15', size=[3, 1]), sg.InputText(key='degrees15', size=[5, 1]), sg.InputText(key='minutes15', size=[5, 1]),
                                    sg.InputText(key='seconds15', size=[5, 1]),sg.Combo(['E', 'W'], key='easting15', size=[3, 1]),sg.InputText(key='feet15', size=[6, 1]), sg.Text('', key='-output-15', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing16', size=[3, 1]),sg.InputText(key='degrees16', size=[5, 1]), sg.InputText(key='minutes16', size=[5, 1]),
                                    sg.InputText(key='seconds16', size=[5, 1]),sg.Combo(['E', 'W'], key='easting16', size=[3, 1]),sg.InputText(key='feet16', size=[6, 1]), sg.Text('', key='-output-16', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing17', size=[3, 1]),sg.InputText(key='degrees17', size=[5, 1]), sg.InputText(key='minutes17', size=[5, 1]),
                                    sg.InputText(key='seconds17', size=[5, 1]),sg.Combo(['E', 'W'], key='easting17', size=[3, 1]),sg.InputText(key='feet17', size=[6, 1]), sg.Text('', key='-output-17', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing18', size=[3, 1]),sg.InputText(key='degrees18', size=[5, 1]), sg.InputText(key='minutes18', size=[5, 1]),
                                    sg.InputText(key='seconds18', size=[5, 1]),sg.Combo(['E', 'W'], key='easting18', size=[3, 1]),sg.InputText(key='feet18', size=[6, 1]), sg.Text('', key='-output-18', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing19', size=[3, 1]),sg.InputText(key='degrees19', size=[5, 1]), sg.InputText(key='minutes19', size=[5, 1]),
                                    sg.InputText(key='seconds19', size=[5, 1]),sg.Combo(['E', 'W'], key='easting19', size=[3, 1]),sg.InputText(key='feet19', size=[6, 1]), sg.Text('', key='-output-19', size=(30, 1))],
                                   [sg.Text("Enter Here: "), sg.Combo(['N', 'S'], key='northing20', size=[3, 1]),sg.InputText(key='degrees20', size=[5, 1]), sg.InputText(key='minutes20', size=[5, 1]),
                                    sg.InputText(key='seconds20', size=[5, 1]),sg.Combo(['E', 'W'], key='easting20', size=[3, 1]),sg.InputText(key='feet20', size=[6, 1]), sg.Text('', key='-output-20', size=(30, 1))],

                                   [sg.Submit()]]
                win_MBImport = sg.Window('QGIS M&B Import/Calculator').Layout(layout_MBImport)
                while True:
                    ev_MBImport, vals_MBImport = win_MBImport.Read()
                    if ev_MBImport is None or ev_MBImport == 'Exit':
                        winActive_MBImport = False
                        win_MBImport.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MBImport == 'Submit':
                        output1 = mbDegreesImport(vals_MBImport['northing1'], vals_MBImport['degrees1'], vals_MBImport['minutes1'],
                                                  vals_MBImport['seconds1'], vals_MBImport['easting1'], vals_MBImport['feet1'])
                        output2 = mbDegreesImport(vals_MBImport['northing2'], vals_MBImport['degrees2'], vals_MBImport['minutes2'],
                                                  vals_MBImport['seconds2'], vals_MBImport['easting2'],vals_MBImport['feet2'])
                        output3 = mbDegreesImport(vals_MBImport['northing3'], vals_MBImport['degrees3'],vals_MBImport['minutes3'],
                                                  vals_MBImport['seconds3'], vals_MBImport['easting3'],vals_MBImport['feet3'])
                        output4 = mbDegreesImport(vals_MBImport['northing4'], vals_MBImport['degrees4'], vals_MBImport['minutes4'],
                                                  vals_MBImport['seconds4'], vals_MBImport['easting4'],  vals_MBImport['feet4'])
                        output5 = mbDegreesImport(vals_MBImport['northing5'], vals_MBImport['degrees5'], vals_MBImport['minutes5'],
                                                  vals_MBImport['seconds5'], vals_MBImport['easting5'],vals_MBImport['feet5'])
                        output6 = mbDegreesImport(vals_MBImport['northing6'], vals_MBImport['degrees6'],vals_MBImport['minutes6'],
                                                  vals_MBImport['seconds6'], vals_MBImport['easting6'],vals_MBImport['feet6'])
                        output7 = mbDegreesImport(vals_MBImport['northing7'], vals_MBImport['degrees7'],vals_MBImport['minutes7'],
                                                  vals_MBImport['seconds7'], vals_MBImport['easting7'],vals_MBImport['feet7'])
                        output8 = mbDegreesImport(vals_MBImport['northing8'], vals_MBImport['degrees8'],vals_MBImport['minutes8'],
                                                  vals_MBImport['seconds8'], vals_MBImport['easting8'],vals_MBImport['feet8'])
                        output9 = mbDegreesImport(vals_MBImport['northing9'], vals_MBImport['degrees9'],vals_MBImport['minutes9'],
                                                  vals_MBImport['seconds9'], vals_MBImport['easting9'],vals_MBImport['feet9'])
                        output10 = mbDegreesImport(vals_MBImport['northing10'], vals_MBImport['degrees10'],vals_MBImport['minutes10'],
                                                  vals_MBImport['seconds10'], vals_MBImport['easting10'],vals_MBImport['feet10'])
                        output11 = mbDegreesImport(vals_MBImport['northing11'], vals_MBImport['degrees11'],vals_MBImport['minutes11'],
                                                  vals_MBImport['seconds11'], vals_MBImport['easting11'], vals_MBImport['feet11'])
                        output12 = mbDegreesImport(vals_MBImport['northing12'], vals_MBImport['degrees12'],vals_MBImport['minutes12'],
                                                  vals_MBImport['seconds12'], vals_MBImport['easting12'],vals_MBImport['feet12'])
                        output13 = mbDegreesImport(vals_MBImport['northing13'], vals_MBImport['degrees13'],vals_MBImport['minutes13'],
                                                  vals_MBImport['seconds13'], vals_MBImport['easting13'],vals_MBImport['feet13'])
                        output14 = mbDegreesImport(vals_MBImport['northing14'], vals_MBImport['degrees14'],vals_MBImport['minutes14'],
                                                  vals_MBImport['seconds14'], vals_MBImport['easting14'],vals_MBImport['feet14'])
                        output15 = mbDegreesImport(vals_MBImport['northing15'], vals_MBImport['degrees15'],vals_MBImport['minutes15'],
                                                  vals_MBImport['seconds15'], vals_MBImport['easting15'], vals_MBImport['feet15'])
                        output16 = mbDegreesImport(vals_MBImport['northing16'], vals_MBImport['degrees16'], vals_MBImport['minutes16'],
                                                  vals_MBImport['seconds16'], vals_MBImport['easting16'],vals_MBImport['feet16'])
                        output17 = mbDegreesImport(vals_MBImport['northing17'], vals_MBImport['degrees17'],vals_MBImport['minutes17'],
                                                  vals_MBImport['seconds17'], vals_MBImport['easting17'],vals_MBImport['feet17'])
                        output18 = mbDegreesImport(vals_MBImport['northing18'], vals_MBImport['degrees18'], vals_MBImport['minutes18'],
                                                  vals_MBImport['seconds18'], vals_MBImport['easting18'], vals_MBImport['feet18'])
                        output19 = mbDegreesImport(vals_MBImport['northing19'], vals_MBImport['degrees19'],vals_MBImport['minutes19'],
                                                  vals_MBImport['seconds19'], vals_MBImport['easting19'],vals_MBImport['feet19'])
                        output20 = mbDegreesImport(vals_MBImport['northing20'], vals_MBImport['degrees20'],vals_MBImport['minutes20'],
                                                   vals_MBImport['seconds20'], vals_MBImport['easting20'],vals_MBImport['feet20'])
                        # Update Outputs ---------
                        win_MBImport.FindElement('-output-1').Update(output1)
                        win_MBImport.FindElement('-output-2').Update(output2)
                        win_MBImport.FindElement('-output-3').Update(output3)
                        win_MBImport.FindElement('-output-4').Update(output4)
                        win_MBImport.FindElement('-output-5').Update(output5)
                        win_MBImport.FindElement('-output-6').Update(output6)
                        win_MBImport.FindElement('-output-7').Update(output7)
                        win_MBImport.FindElement('-output-8').Update(output8)
                        win_MBImport.FindElement('-output-9').Update(output9)
                        win_MBImport.FindElement('-output-10').Update(output10)
                        win_MBImport.FindElement('-output-11').Update(output11)
                        win_MBImport.FindElement('-output-12').Update(output12)
                        win_MBImport.FindElement('-output-13').Update(output13)
                        win_MBImport.FindElement('-output-14').Update(output14)
                        win_MBImport.FindElement('-output-15').Update(output15)
                        win_MBImport.FindElement('-output-16').Update(output16)
                        win_MBImport.FindElement('-output-17').Update(output17)
                        win_MBImport.FindElement('-output-18').Update(output18)
                        win_MBImport.FindElement('-output-19').Update(output19)
                        win_MBImport.FindElement('-output-20').Update(output20)

                        output_total = np.vstack((output1,output2,output3,output4,output5,output6,output7,output8,output9,
                                                     output10,output11,output12,output13,output14,output15,output16,output17,
                                                     output18,output19,output20))

                        output_df = pd.DataFrame(data=output_total, columns = ['Azimuth', 'Distance'])
                        pd.DataFrame.to_csv(self=output_df, path_or_buf=r'C:\Users\Accounting\Downloads\MBImport.csv',
                                            encoding='utf-8', index=False, sep=';')

            if not winActive_MBExtract and ev_Programs == 'M&B String Extractor':
                winActive_MBExtract = True
                win_Programs.Hide()
                layout_MBExtract = [[sg.Text("Enter M&B Description: ")],
                                    [sg.InputText('', key = '-string-', size = [50,30])],
                                    [sg.Text('', key='-output-', size=[50, 30])],
                                    [sg.Text(r"Remember: Don't trust computers.. check your output")],
                                    [sg.Submit()]]
                win_MBExtract = sg.Window('M&B String Extractor').Layout(layout_MBExtract)
                while True:
                    ev_MBExtract, vals_MBExtract = win_MBExtract.Read()
                    if ev_MBExtract is None or ev_Programs == 'Exit':
                        winActive_MBExtract = False
                        win_MBExtract.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MBExtract == 'Submit':
                        output = MBExtract(vals_MBExtract['-string-'])
                        win_MBExtract.FindElement('-output-').Update(output)

            if not winActive_MBUltimate and ev_Programs == 'The Ultimate M&B Program':
                winActive_MBUltimate = True
                win_Programs.Hide()
                layout_MBUltimate = [[sg.Text("""This program does the following:
                                                1. Separates a raw M&B Description and ouputs here
                                                2. Calculates Decimal Degrees from Northing/Easting
                                                3. Exports the conversion as MBUltimate in the Downloads
                                                
                                                IMPORTANT NOTE:
                                                Make sure to quality check the description inputted.
                                                Make sure the northing/easting isn't separated by
                                                spaces, make sure that there aren't any weird
                                                characters, etc. If you don't, the program won't work.""")],
                                     [sg.Text('Enter the description here: '), sg.InputText('', key = '-input-', size = [50, 10])],
                                     [sg.Text('', key='-output-', size = (50,30))],
                                     [sg.Text(r"Remember: Don't trust computers.. check your output")],
                                     [sg.Text('Next Step: Import MBUltimate into the M&B Plugin which will draw it for you.')],
                                     [sg.Submit()]
                                     ]


                win_MBUltimate = sg.Window('The Ultimate M&B Program').Layout(layout_MBUltimate)
                while True:
                    ev_MBUltimate, vals_MBUltimate = win_MBUltimate.Read()
                    if ev_MBUltimate is None or ev_Programs == 'Exit':
                        winActive_MBUltimate = False
                        win_MBUltimate.Close()
                        win_Programs.UnHide()
                        break
                    if ev_MBUltimate == 'Submit':
                        output = MBUltimate(vals_MBUltimate['-input-'])
                        win_MBUltimate.FindElement('-output-').Update(output)

            if not winActive_AddTract and ev_Programs == 'Add Tract':
                winActive_AddTract = True
                win_Programs.Hide()
                layout_AddTract = [[sg.Text("""This program does the following:
                1. Formats tract description
                2. Opens iLandman and the Add Tract page
                3. Fills in all the information based on user input

                IMPORTANT NOTE:
                Make sure to quality check the description inputted.
                """)],
                                   [sg.Text('Enter the tract info here: ')],
                                   [sg.Text('Township: '), sg.InputText('', key='-township-', size=[5, 10]), sg.Text('(4S)')],
                                   [sg.Text('Range: '), sg.InputText('', key='-range-', size=[5, 10]), sg.Text('(3W)')],
                                   [sg.Text('Section: '), sg.InputText('', key='-section-', size=[5, 10]), sg.Text('(29)')],
                                   [sg.Text('Tract Number: '), sg.InputText('', key='-tractnumber-', size=[5, 10]), sg.Text('(5)')],
                                   [sg.Text('County: '), sg.Combo(['Duchesne', 'Uintah'], key='-county-', size=[10, 10]), sg.Text('(Duchesne)')],
                                   [sg.Text('Gross Acreage: '), sg.InputText('', key='-acres-', size=[7, 10]), sg.Text('(100)')],
                                   [sg.Text('Description: '), sg.InputText('', key='-description-', size=[30, 10]), sg.Text('(Section 29:..)')],
                                   [sg.Checkbox('Add to iLandman?', key = '-iLm-'), sg.Checkbox('M&B?', key='-MB-'), sg.Checkbox('Add to Map?', key='-Map-')],
                                   [sg.Text('', key='-output-', size=(15, 1))],
                                   [sg.Text(r"Remember: Don't trust computers.. check your output")],
                                   [sg.Text(
                                         'Note: If the tract is M&B there are a few places\n'
                                         'where user input are required. Check the Application\n'
                                         'Manual for details.')],
                                   [sg.Submit()]
                                    ]

                win_AddTract = sg.Window('Add Tract').Layout(layout_AddTract)
                while True:
                    ev_AddTract, vals_AddTract = win_AddTract.Read()
                    if ev_AddTract is None or ev_Programs == 'Exit':
                        winActive_AddTract = False
                        win_AddTract.Close()
                        win_Programs.UnHide()
                        break
                    if ev_AddTract == 'Submit':
                        if vals_AddTract['-iLm-']:
                            addtract = threading.Thread(target = iLandmanTract, args = (vals_AddTract['-township-'], vals_AddTract['-range-'],
                                                   vals_AddTract['-section-'], vals_AddTract['-tractnumber-'],
                                                   vals_AddTract['-county-'], vals_AddTract['-acres-'],
                                                   vals_AddTract['-description-'], vals_AddTract['-MB-'],
                                                   vals_AddTract['-Map-']), daemon = True)
                            addtract.start()
                            win_AddTract.FindElement('-output-').Update(threadMessage)
                        if vals_AddTract['-MB-'] and vals_AddTract['-Map-']:
                            mapMB = threading.Thread(target = iLandmanTractMap, args = (vals_AddTract['-township-'], vals_AddTract['-range-'],
                                               vals_AddTract['-section-'], vals_AddTract['-tractnumber-'],
                                               vals_AddTract['-county-'], 'NE',
                                               vals_AddTract['-description-']), daemon = True)
                            mapMB.start()
                            win_AddTract.FindElement('-output-').Update(threadMessage)

    # Files Window
    if not winActive_Files and ev_Main == 'Files':
        winActive_Files = True
        win_Main.Hide()
        layout_Files = [[sg.Text('Which File?')],
                        [sg.Text('Useful Folders/Documents')],
                        [sg.Button('Job Manual')],
                        [sg.Button('Code')],
                        [sg.Button('Important Excel Documents')],
                        [sg.Text('Maps')],
                        [sg.Button('QGIS Map')],
                        [sg.Button('Other Map Layers')],
                        [sg.Text('Financial Database')],
                        [sg.Button('WEM Financial Database 2020')],
                        [sg.Button('Database Spreadsheet')]
                        ]
        win_Files = sg.Window('Files').Layout(layout_Files)

        while True:
            ev_Files, vals_Files = win_Files.Read()
            if ev_Files is None or ev_Files == 'Exit':
                winActive_Files = False
                win_Files.Close()
                win_Main.UnHide()
                break
            if ev_Files == 'Job Manual':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\JOB MANUAL.docx')
            if ev_Files == 'Important Excel Documents':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\QGIS\IMPORTANT EXCEL DOCS\QGIS')
            if ev_Files == 'QGIS Map':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\QGIS\MAIN MAP.qgz')
            if ev_Files == 'Other Map Layers':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\QGIS\Other Map Layers')
            if ev_Files == 'Code':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\Code')
            if ev_Files == 'WEM Financial Database 2020':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\WEM Financial\WEM Financial Database 2020.accdb')
            if ev_Files == 'Database Spreadsheet':
                os.startfile(r'\\WEM-MASTER\Sensitive Data\WEM Uintah\WEM Financial\WEM 2020 Finances.xlsx')




# EXPORTING TO APPLICATION -
# cd //WEM-MASTER/Sensitive Data/WEM Uintah/Maps/WEM_Mapping Application
# cd \\WEM-MASTER\Sensitive Data\WEM Uintah\Maps\WEM_Mapping Application
# C:\Users\Accounting\Documents\Peter\Application>pyinstaller WEM_Mapping v2.py --windowed --hidden-import time --hidden-import PySimpleGUI --hidden-import os --hidden-import rpy2.robjects --hidden-import datetime --hidden-import shutil --onedir
# First: cd *directory* -- wherever you are putting the application
# Second: See above. Make sure to include:
#           --windowed (can't remember what this does)
#           -i icon.ico (reads an icon file, make sure its in your cd)
#           --hidden-import *package name* (imports the packages you used, include all of them)
#           --onefile (compresses several files into one application)
