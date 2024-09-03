from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup
import pandas as pd

# Set up the WebDriver using WebDriverManager
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

try:
    # Navigate to the printer's web interface
    driver.get('http://192.168.1.100')

    # Log in
    WebDriverWait(driver, 40).until(
        EC.presence_of_element_located((By.NAME, 'user_name'))
    ).send_keys('7654321')
    
    driver.find_element(By.NAME, 'pwd').send_keys('7654321')
    driver.find_element(By.XPATH, '//a[@onclick="login();return false;"]').click()

    # Navigate to the report page
    driver.get('http://192.168.1.100/_dept.html?dn=1')
     
    # Switch to the 'body' frame
    WebDriverWait(driver, 40).until(
        EC.frame_to_be_available_and_switch_to_it((By.NAME, 'body'))
    )

    # Extract the page's HTML content
    page_source = driver.page_source

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(page_source, 'html.parser')

    # Initialize lists to hold the data
    data = []

    # Find all the rows in the table
    rows = soup.find_all('tr')

    # Iterate through the rows and extract the data
    for row in rows[3:]:  # Skip the first two rows (headers)
        cols = row.find_all('td')
        if len(cols) >= 8:
            department_id = cols[2].text.strip()
            total_prints = cols[3].text.strip().replace('<br>', '')
            copy = cols[4].text.strip().replace('<br>', '')
            # bw_scan = cols[5].text.strip().replace('<br>', '')
            # color_scan = cols[6].text.strip().replace('<br>', '')
            print_ = cols[7].text.strip().replace('<br>', '')
            data.append([department_id, total_prints, copy, print_])
            # data.append([department_id, total_prints, copy, bw_scan, color_scan, print_])

    # Create a pandas DataFrame from the data
    df = pd.DataFrame(data, columns=['ID', 'Total', 'Copy', 'Print'])

    # Save the DataFrame to an Excel file
    df.to_excel('output.xlsx', index=False)

    print("Excel file has been generated successfully!")

finally:
    # Close the browser
    driver.quit()
