from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import ElementNotInteractableException, ElementClickInterceptedException
from selenium.webdriver import ActionChains
import time
import glob
import os
from datetime import datetime

# Set up
user_email = r'pvankatwyk@gmail.com'
user_password = r'A&8BYpG%&*xy#s!'
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

analytics = findElement(r'//*[@id="piemenuCanvas"]')
action = webdriver.common.action_chains.ActionChains(driver)
action.move_to_element_with_offset(analytics, 150, 200).perform()
action.click()
action.perform()

owner = findElement(r'//*[@id="geoanalytical-ownership-by-owner"]/span')
owner.click()
time.sleep(1)

# Get session ID to be able to click
target = findElement(r'//*[@id="analytic-criteria-form"]/div/table/tbody/tr[4]/td[2]/div/a')
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
ok = findElement(r'/html/body/div[11]/div[3]/div/button[2]/span')
ok.click()
addanalytic = findElement(r'//*[@id="save-form"]')
addanalytic.click()
time.sleep(3)

# Get download ID
target2 = findElement('//*[@id="layerListContainer"]/div[contains(., "Ownership By Owner")]')
target2 = str(target2.get_attribute('id'))
ids2 = target2[len('layerContainer'):len(target2)]
hamburger = findElement(f'//*[@id="layerImage{ids2}"]')
hamburger.click()

options = findElement(r'//*[@id="ui-id-10"]')
options.click()
dwnld = findElement(f'//*[@id="legendOptionDownload{ids2}"]')
dwnld.click()
time.sleep(5)
cls = findElement(r'/html/body/div[13]/div[1]/button/span[1]')
time.sleep(2)
cls.click()
x = findElement(r'/html/body/div[12]/div[1]/button/span[1]')
x.click()

# Rename last download to ownership
todayDate = datetime.date(datetime.now())
time.sleep(5)
ownerDate = r'ownership ' + str(todayDate) + '.zip'
list_of_files = glob.glob(r'C:\Users\Accounting\Downloads\*')
#list_of_files = glob.glob(r'C:\Users\Peter\Downloads\*')
ownership = max(list_of_files, key=os.path.getctime)
newownerName = r'C:\\Users\\Accounting\\Downloads\\' + str(ownerDate)
#newownerName = r'C:\\Users\\Peter\\Downloads\\' + str(ownerDate)
os.rename(ownership, newownerName)

# Ownership by Lessee
actionChains = ActionChains(driver)
fullmap = findElement(r'//*[@id="container"]')
actionChains.context_click(fullmap).perform()

analytics = findElement(r'//*[@id="piemenuCanvas"]')
action = webdriver.common.action_chains.ActionChains(driver)
action.move_to_element_with_offset(analytics, 150, 200).perform()
action.click()
action.perform()

lessee = findElement(r'//*[@id="geoanalytical-ownership-by-lessee"]/span')
lessee.click()

target3 = findElement(r'//*[@id="analytic-criteria-form"]/div/table/tbody/tr[12]/td[2]/div/a')
target3 = str(target3.get_attribute('id'))
ids3 = target3[len('data-picker-'):-len('-trigger')]
contowner = findElement(f'//*[@id="data-picker-{ids3}-trigger"]')
contowner.click()
time.sleep(1)
bar2 = findElement(f'//*[@id="data-picker-{ids3}-simple"]/div/input[2]')
bar2.click()
bar2.send_keys('WEM')
#wem2 = findElement(r'//*[@id="115153"]')
wem1 = findElement(r'//*[@id="90095"]')
#wem2.click()
wem1.click()
okay = findElement(r'/html/body/div[20]/div[3]/div/button[2]/span')
time.sleep(2)
okay.click()
conType = findElement(r'//*[@id="analytic-criteria-form"]/div/table/tbody/tr[14]/td[2]/div/a')
conType.click()
own = findElement(r'//*[@id="Ownership"]')
own.click()
okk = findElement(r'/html/body/div[22]/div[3]/div/button[2]/span')
okk.click()
addA = findElement(r'//*[@id="save-form"]')
addA.click()

# Get download ID
target4 = WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="layerListContainer"]/div[contains(., "Ownership By Lessee")]')))
#target4 = findElement(r'//*[@id="layerListContainer"]/div[contains(., "Ownership By Lessee")]')
target4 = str(target4.get_attribute('id'))
ids4 = target4[len('layerContainer'):len(target4)]
hamburger = findElement(f'//*[@id="layerImage{ids4}"]')
hamburger.click()
time.sleep(3)
options = findElement(r'//*[@id="ui-id-90"]')
options.click()
dwnld = findElement(f'//*[@id="legendOptionDownload{ids4}"]')
dwnld.click()
time.sleep(5)
cls = findElement(r'/html/body/div[35]/div[1]/button/span[1]')
time.sleep(2)
cls.click()
x = findElement(r'/html/body/div[34]/div[1]/button/span[1]')
x.click()

# Rename last download to leasehold
time.sleep(5)
leaseholdDate = r'leasehold ' + "WEM I "+ str(todayDate) + '.zip'
list_of_files = glob.glob(r'C:\Users\Accounting\Downloads\*')
#list_of_files = glob.glob(r'C:\Users\Peter\Downloads\*')
leasehold = max(list_of_files, key=os.path.getctime)
newleaseName = r'C:\\Users\\Accounting\\Downloads\\'+str(leaseholdDate)
#newleaseName = r'C:\\Users\\Peter\\Downloads\\'+str(leaseholdDate)
os.rename(leasehold, newleaseName)

# Ownership by Lessee WEM II
actionChains = ActionChains(driver)
fullmap = findElement(r'//*[@id="container"]')
actionChains.context_click(fullmap).perform()

analytics = findElement(r'//*[@id="piemenuCanvas"]')
action = webdriver.common.action_chains.ActionChains(driver)
action.move_to_element_with_offset(analytics, 150, 200).perform()
action.click()
action.perform()

lessee = findElement(r'//*[@id="geoanalytical-ownership-by-lessee"]/span')
lessee.click()

target3 = findElement(r'//*[@id="analytic-criteria-form"]/div/table/tbody/tr[12]/td[2]/div/a')
target3 = str(target3.get_attribute('id'))
ids3 = target3[len('data-picker-'):-len('-trigger')]
contowner = findElement(f'//*[@id="data-picker-{ids3}-trigger"]')
contowner.click()
time.sleep(1)
bar3 = findElement(f'//*[@id="data-picker-{ids3}-simple"]/div/input[2]')
bar3.click()
bar3.send_keys('WEM')
wem2 = findElement(f'//*[@id="data-picker-{ids3}-multiple-results"]/ul/li[contains(.,"WEM Uintah II")]')
#wem1 = findElement(r'//*[@id="90095"]')
wem2.click()
#wem1.click()
okay = findElement(r'/html/body/div[42]/div[3]/div/button[2]/span')
time.sleep(2)
okay.click()
conType = findElement(r'//*[@id="analytic-criteria-form"]/div/table/tbody/tr[14]/td[2]/div/a')
conType.click()
target5 = findElement(r'/html/body/div[44]/div[2]')
target5 = str(target5.get_attribute('id'))
ids5 = target5[len('data-picker-'):]
owned = findElement(f'//*[@id="data-picker-{ids5}-multiple-results"]/ul/li[1]')
owned.click()
#owned = findElement(r'//*[@id="Ownership"]/b')
#owned.click()
okkkk = findElement(r'/html/body/div[44]/div[3]/div/button[2]/span')
okkkk.click()
addA = findElement(r'//*[@id="save-form"]')
addA.click()


time.sleep(10)
# Get download ID
target6 = WebDriverWait(driver, 1000).until(EC.presence_of_element_located((By.XPATH, r'//*[@id="layerListContainer"]/div[contains(., "Ownership By Lessee")]')))
target6 = str(target6.get_attribute('id'))
ids6 = target6[len('layerContainer'):len(target6)]

hamburger = findElement(f'//*[@id="layerImage{ids6}"]')
hamburger.click()
time.sleep(3)
options = findElement(r'//*[@id="ui-id-170"]')
options.click()
dwnld = findElement(f'//*[@id="legendOptionDownload{ids6}"]')
dwnld.click()
time.sleep(8)
cls = findElement(r'/html/body/div[56]/div[3]/div/button/span')
time.sleep(2)
cls.click()
x = findElement(r'/html/body/div[55]/div[1]/button/span[1]')
x.click()

# Rename last download to leasehold
time.sleep(5)
leaseholdDate = r'leasehold ' + "WEM II " + str(todayDate) + '.zip'
list_of_files = glob.glob(r'C:\Users\Accounting\Downloads\*')
#list_of_files = glob.glob(r'C:\Users\Peter\Downloads\*')
leasehold = max(list_of_files, key=os.path.getctime)
newleaseName = r'C:\\Users\\Accounting\\Downloads\\'+str(leaseholdDate)
#newleaseName = r'C:\\Users\\Peter\\Downloads\\'+str(leaseholdDate)
os.rename(leasehold, newleaseName)

driver.quit()

print('Done!')
