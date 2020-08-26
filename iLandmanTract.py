from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException
import time
import re


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




#iLandmanTract('3S', '4W', '17', '1', 'Duchesne', '240', 'Section 1: NE/4, N/2SE/4', False, False)
#iLandmanTract('2S', '2W', '31', '10', 'Duchesne', '152.74', 'Section 31: W/2NW/4, E/2NW/4', False, True)
#iLandmanTract(Township, Range, Section, TractNumber, County, GrossAcres, Description, MB, Map)
