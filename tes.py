import csv

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

DEPT_ID = '7654321'
PASSWD = '7654321'
HOST = '192.168.1.65'
CSV_FILE = 'contacts.csv'
ADDRESS_LIST_XPATH = '/html/body/div/div[2]/div[2]/div/div[1]/div/div[3]/div/div/table/tbody/tr[2]/td[1]/a[2]'
HEADLESS = False

# Browser setup
options = webdriver.Chrome('./chromedriver')
options.get("https://www.python.org")
# options = webdriver.FirefoxOptions()
# options.headless = HEADLESS
# browser=webdriver.Firefox(options=options)

# Login
browser.get(f'https://{HOST}:8443/rps/')
browser.find_element_by_xpath('/html/body/form/div/div/div[2]/div/div/div[1]/table/tbody[1]/tr[2]/td/input').send_keys(DEPT_ID)
browser.find_element_by_xpath('/html/body/form/div/div/div[2]/div/div/div[1]/table/tbody[1]/tr[3]/td/input').send_keys(PASSWD)
browser.find_element_by_xpath('/html/body/form/div/div/div[2]/div/div/div[2]/fieldset/input').click()

# Navigate to address book
WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.XPATH, '/html/body/form[3]/div/div[2]/div[2]/div/div[2]/div/div[2]/ul/li[3]/a/span'))).click()

# Navigate to selected address book
WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.XPATH, ADDRESS_LIST_XPATH))).click()

# Import contacts
with open(CSV_FILE, 'r') as fh:
    reader = csv.DictReader(fh)

    for row in reader:
        print(row['DisplayName'], row['UserPrincipalName'])

        # Register New Destination
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.XPATH, '/html/body/form[5]/div/div[2]/div[2]/div/div[1]/div/div[4]/div/div[1]/fieldset[1]/input[1]'))).click()

        # Enter user data
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, '#ANAME'))).send_keys(row['DisplayName'])
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, '#AAD1'))).send_keys(row['UserPrincipalName'])
        WebDriverWait(browser, 10).until(ec.presence_of_element_located((By.CSS_SELECTOR, '#OK_Button'))).click()