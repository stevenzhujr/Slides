from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
import time

# Gather data

chrome_options = Options()
chrome_options.add_extension("./Adblock.crx")
driver = webdriver.Chrome(chrome_options)
driver.maximize_window()
t = time.time()
driver.set_page_load_timeout(20)
time.sleep(5)
links = open('links.txt', 'r')
lunk = links.readline()
curr = open('curr.txt', 'w')

# slide 7

try:
    driver.get(lunk)
    shot = driver.find_element(By.XPATH, "//div[@class='page-title symbol-header-info   ng-scope']")
    shot.screenshot("SevenTop.png")

    shot = driver.find_element(By.XPATH, "//div[@class='bc-quote-overview row ng-scope']")
    shot.screenshot("SevenMiddle.png")

    shot = driver.find_element(By.XPATH, "//div[@class='barchart-content-block commodity-profile']")
    driver.execute_script('arguments[0].scrollIntoView({block: "center"});', shot)
    shot.screenshot("SevenBottom.png")
except TimeoutException:
    driver.execute_script("window.stop();")

# slide 8

lunk = links.readline()

try:
    driver.get(lunk)
    shot = driver.find_element(By.XPATH, "//span[@class= 'last-change ng-binding']")
    curr.write(shot.text[0:6])

    shot = driver.find_element(By.XPATH, "//div[@class='page-title symbol-header-info   ng-scope']")
    shot.screenshot("EightTop.png")

    shot = driver.find_element(By.XPATH, "//div[@class='block-content table-wrapper clearfix']")
    driver.execute_script('arguments[0].scrollIntoView({block: "center"});', shot)
    shot.screenshot("EightMiddle.png")

    shot = driver.find_element(By.XPATH, "//div[@class='barchart-content-block symbol-price-performance']")
    driver.execute_script('arguments[0].scrollIntoView({block: "center"});', driver.find_element(By.XPATH, "//div[@class='block-title joined clearfix']"))
    shot.screenshot("EightBottom.png")

except TimeoutException:
    driver.execute_script("window.stop();")

driver.close()