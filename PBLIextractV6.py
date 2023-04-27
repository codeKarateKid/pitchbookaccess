from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import time
import pandas as pd


# Read the Excel file into a pandas dataframe
df = pd.read_excel('C:/DST Dropbox/Siddharth Prabhu/LinkedIn Navigator/R-Develop/Generative AI/inputPBurl.xlsx', header=0)

df.columns = ['PBid', 'PBurl']

# Extract the URLs from column A of the dataframe
urls = df['PBurl']

options = Options()
options.add_argument(r"--user-data-dir=C:\\Users\\Siddharth Prabhu\\AppData\\Local\\Google\\Chrome\\User Data")  # e.g. C:\Users\You\AppData\Local\Google\Chrome\User Data
options.add_argument(r'--profile-directory=Profile 3')  # e.g. Profile 3
options.add_argument('--no-sandbox')

service = Service(r'C:\Users\Siddharth Prabhu\Downloads\chromedriver_win32\chromedriver.exe')

driver = webdriver.Chrome(service=service, options=options)

# Navigate to Google
driver.get("https://www.google.com/")

# Wait for 30 seconds
time.sleep(5)

# Navigate to Pitchbook
driver.get("https://my.pitchbook.com/dashboard/")

# Wait for 30 seconds
time.sleep(30)

# Initialize an empty list to store the extracted data
data = []
link_urls = []  # Initialize an empty list to store the link URLs

# Loop through each URL in the list
for url in urls:
    # Navigate to the LinkedIn user's profile page
    driver.get(url)

    # Wait for 30 seconds
    time.sleep(15)

    # Find the table element
    table = driver.find_element(By.XPATH,
                                 "//table[@class='table table_Qq5Cks2RWJAoXrDy-K-lCg table_fixed table_fixed_Qq5Cks2RWJAoXrDy-K-lCg table_v-align-middle table_v-align-middle_Qq5Cks2RWJAoXrDy-K-lCg']")

    # Find all rows of the table body
    rows = table.find_elements(By.XPATH, ".//tbody/tr")

    # Iterate over each row and extract the data
    for row in rows:
        name = row.find_element(By.XPATH, ".//td[1]/span/a").text
        title = row.find_element(By.XPATH, ".//td[2]/span").text
        board_seats = row.find_elements(By.XPATH, ".//td[3]/span/a")

        if board_seats:
            board_seats = board_seats[0].text
        else:
            board_seats = ""
        office = row.find_element(By.XPATH, ".//td[4]/span").text
        linkedin_button = driver.find_element(By.CSS_SELECTOR,
                                              "a.icon-button__hole_qvE3jhy20O6hSDUPug-SqA > i.icon-linkedin")
        linkedin_url = linkedin_button.get_attribute("href")

        data.append([name, title, board_seats, office, linkedin_url, url])

    # Find all links with the specified CSS classes
    # Define the CSS selector for the links
    css_classes = (
        'a.icon-button.icon-button_qvE3jhy20O6hSDUPug-SqA.'
        'icon-button_button.icon-button_button_qvE3jhy20O6hSDUPug-SqA.'
        'icon-button_button_qvE3jhy20O6hSDUPug-SqA.'
        'icon-button_size-l.icon-button_size-l_qvE3jhy20O6hSDUPug-SqA.'
        'icon-button_size-l_qvE3jhy20O6hSDUPug-SqA.'
        'icon-button_flat.icon-button_flat_qvE3jhy20O6hSDUPug-SqA.'
        'icon-button_flat_qvE3jhy20O6hSDUPug-SqA'
    )

    # Find all links with the specified CSS classes
    links = driver.find_elements(By.CSS_SELECTOR, css_classes)

    # Extract the href attribute value of each link
    for link in links:
        url = link.get_attribute('href')
        link_urls.append(url)


# Create a DataFrame from the extracted data
df = pd.DataFrame(data, columns=["Name", "Title", "Board Seats", "Office", "URL"])

# Save the output to an Excel file
output_path = "C:/DST Dropbox/Siddharth Prabhu/LinkedIn Navigator/R-Develop/pitchbookAccess/output.xlsx"
df.to_excel(output_path, index=False)

df = pd.DataFrame(link_urls, columns=["LinkedIn"])
output_path2 = "C:/DST Dropbox/Siddharth Prabhu/LinkedIn Navigator/R-Develop/pitchbookAccess/output2.xlsx"
df.to_excel(output_path2, index=False)



# Close the browser
driver.quit()
