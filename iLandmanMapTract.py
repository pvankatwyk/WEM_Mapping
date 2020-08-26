from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException
from selenium.webdriver import ActionChains
import time
import re
from soundfxn import sound

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

#iLandmanTractMap('4S', '3W', '29', '5', 'Duchesne', 'NE', r"Section 29: SW/4NW1/4; W1/2NE1/4; also the tract beginning at the NW corner NE1/4NE1/4 of said Section 29 and running thence South 1320 feet; thence E 470 feet; thence N 18°35' West 408 feet; thence South 59°30' West 56 feet; thence North 9°20' East 193 feet; thence North 45°00' East 113 feet; thence North 4°50' East 213 feet; thence North 69°15' East 124 feet; thence North 64°20' East 175 feet; thence North 71°20' E. 73 feet; thence North 18°40' West 33 feet; thence South 71°20' West 73 feet; thence North 18°40' West 214 feet; thence North 12°05' East 126.6 feet; thence West 643 feet to the place of beginning.", True)
# iLandmanTractMap('4S', '2W', '1', '10', 'Duchesne', 'NE', r"TOWNSHIP 4 SOUTH, RANGE 2 WEST, UINTAH SPECIAL BASE AND MERIDIAN SECTION 1: Beginning at East quarter corner of said Section and running thence South 00°50'15" East 1432.59 feet along the East line of said Section; Thence North 89°12'40" West 543.61 feet; Thence South 00°50' 15" East 399.99 feet to an existing fence; Thence North 87°29'32" East 106.54 feet along said existing fence to an existing fence corner; Thence South 87°47'36" East 164.83 feet along an existing fence to an existing fence corner; Thence South 71°53'21" East 24.97 feet along an existing fence to an existing fence corner; Thence South 02°15'06" East 292.70 feet; Thence South 00°09'10" East 277.72 feet; Thence North 89°13'06" East 244.78 feet to said East line; Thence South 00°50'15" East 243.87 feet along said East line to the Southeast corner of said Section; Thence South 88°43'41" West 1317.70 feet along the South line of said Section to an existing fence; Thence North 01°11'13" West 1323.96 feet along said existing fence to an existing three way fence corner; Thence North 00°57'31" West 1322.67 feet along an existing fence to an existing three way fence corner; Thence North 88°47'12" East 1328.57 feet to the point of beginning. Subject to that portion being used for County Road right-of-way.
# ")



