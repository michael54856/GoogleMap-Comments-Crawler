from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
import winsound
chrome_driver_path = 'chromedriver.exe'

options = webdriver.ChromeOptions()
driver = webdriver.Chrome(executable_path=chrome_driver_path,options=options)


url='https://www.google.com/maps/place/%E5%8C%97%E6%8A%95%E6%BA%AB%E6%B3%89%E5%8D%9A%E7%89%A9%E9%A4%A8/@25.1365857,121.5071437,17z/data=!3m2!4b1!5s0x3442ae44c2a6f701:0x748f2436572c1f88!4m5!3m4!1s0x3442ae50f43af1e9:0xadf18f29697c0a3c!8m2!3d25.1365694!4d121.50715'
fileName = '北投溫泉博物館.xlsx'

driver.get(url)
time.sleep(1.5)

commentButtonStr1 = str(driver.find_element(By.CLASS_NAME, 'DkEaL').get_attribute('aria-label'))
commentButtonStr2 = ''
for i in range(0, len(commentButtonStr1)):
    if commentButtonStr1[i] >= '0' and commentButtonStr1[i] <= '9':
        commentButtonStr2 += commentButtonStr1[i]

commentCount = int(commentButtonStr2)

print(commentCount)

driver.find_element(By.CLASS_NAME, 'DkEaL').click()

time.sleep(0.5)


# 排序的button ID 
sortButtonID = 'DVeyrd.LCTIRd'
buttonList = driver.find_elements(By.CLASS_NAME, sortButtonID)
buttonList[2].click()

time.sleep(0.2)

sortTypeID = 'fxNQSd'
sortTypeList = driver.find_elements(By.CLASS_NAME, sortTypeID)
sortTypeList[1].click()

time.sleep(0.5)

windowElement = driver.find_element(By.CLASS_NAME,'m6QErb.DxyBCb.kA9KIf.dS8AEf')

# Get scroll height
last_height = driver.execute_script("return document.body.scrollHeight",windowElement)
SCROLL_PAUSE_TIME = 0.1

scrollCounter = 0
while True:
    time.sleep(SCROLL_PAUSE_TIME)
    new_height = driver.execute_script("return arguments[0].scrollHeight",windowElement)
    if new_height != last_height:
        scrollCounter += 1
    last_height = new_height
    # Scroll down to bottom
    driver.execute_script('arguments[0].scrollBy(0, 50000);',windowElement)
    
    if(scrollCounter*10 >= commentCount):
        break
    if scrollCounter == 90:
        break

    


expandButtons = driver.find_elements(By.CLASS_NAME, "w8nwRe.kyuRq")
for expand in expandButtons:
    expand.click()


reviewerList_Element = driver.find_elements(By.CSS_SELECTOR, ".d4r55 > span")
#for reviewer in reviewerList:
#    print(reviewer .text)

revieweContent_Element = driver.find_elements(By.CLASS_NAME, "wiI7pd")
#for content in revieweContent:
#    print(content.text)

stars_Element = driver.find_elements(By.CLASS_NAME, "kvMYJc")
#for star in stars:
#    print(star.get_attribute("aria-label"))

commentCount = len(reviewerList_Element)


col1 = "評論者"
col2 = "評分"
col3 = "評論內容"

reviewerList = []
reviewContent = []
stars = []

for i in reviewerList_Element:
    reviewerList.append(i.text)

for i in revieweContent_Element:
    parent = i.find_element(By.XPATH,"./..")
    if parent.get_attribute("id") != "":
        reviewContent.append(i.text)

for i in stars_Element:
    stars.append(i.get_attribute("aria-label").split()[0])

print(len(reviewerList))
print(len(reviewContent))
print(len(stars))

data = pd.DataFrame({col3:reviewContent})
data.to_excel(fileName, sheet_name='sheet1', index=False)



frequency = 2500  # Set Frequency To 2500 Hertz
duration = 200  # Set Duration To 1000 ms == 1 second
winsound.Beep(frequency, duration)