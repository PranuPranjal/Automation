import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

df = pd.read_excel("C:/Users/Lenovo/Desktop/Web/Kumbh/output2.xlsx") #Add path of survey.xlsx in your device

# Correct the path of Chrome driver in your device (use forward slashes or escape backslashes)
driver_path = 'C:/Users/Lenovo/Downloads/chromedriver-win64 (1)/chromedriver-win64/chromedriver.exe'

# Create a Service object for Chromedriver
service = Service(driver_path)

# Initialize the Webdriver with the Service object
driver = webdriver.Chrome(service=service)

# Open the login page
driver.get('https://www.uptisurvey.in/mahakumbh/mahakumbh-survey')  # Replace with the form URL

# Login process
# Wait for the username field to be visible and enter the username
username_input = WebDriverWait(driver, 10).until(
    EC.visibility_of_element_located((By.NAME, 'username'))  # Replace with actual field name
)
username_input.send_keys('7007202027')  # Replace with actual username

# Wait for the password field to be visible and enter the password
password_input = driver.find_element(By.NAME, 'password')  # Replace with actual field name
password_input.send_keys('1234567')  # Replace with actual password

# Find the login button and click it
login_button = driver.find_element(By.NAME, 'login')  # Replace with correct name or selector
login_button.click()

driver.get('https://www.uptisurvey.in/mahakumbh/mahakumbh-survey/')  # If not already at the form
a=0
# Iterate through each row in the dataframe
for index, row in df.iterrows():
    driver.get('https://www.uptisurvey.in/mahakumbh/mahakumbh-survey/')
    # Fill the Name field
    name_input = driver.find_element(By.NAME, 'input_22')
    name_input.send_keys(row['Name'])

    # Fill the Mobile Number field
    mobile_input = driver.find_element(By.NAME, 'input_5')
    mobile_input.send_keys(row['Mobile Number'])

    # Select the State from the dropdown
    state_dropdown = Select(driver.find_element(By.NAME, 'input_6'))  # Replace with actual dropdown name
    state_dropdown.select_by_visible_text(row['State'])

    # Fill the City field
    city_input = driver.find_element(By.NAME, 'input_24')
    city_input.send_keys(row['City'])

    # Select the Age group
    age_dropdown = Select(driver.find_element(By.NAME, 'input_7'))
    age_dropdown.select_by_visible_text(row['Age'])
   
    # Select the Gender
    gender_dropdown = Select(driver.find_element(By.NAME, 'input_8'))
    gender_dropdown.select_by_visible_text(row['Gender'])

    # Fill the 'Last Visit' field
    last_visited = driver.find_element(By.NAME, 'input_10')
    last_visited.send_keys(row['LastVisit'])

    # Fill the 'Next Visit' field
    next_visit = driver.find_element(By.NAME, 'input_11')
    next_visit.send_keys(row['NextVisit'])

    # Fill the 'Favourite Visit' field
    favourite_attraction = driver.find_element(By.NAME, 'input_12')
    favourite_attraction.send_keys(row['FavVisit'])

    # Select 'Yes' for the Plan field
    planned_dropdown = Select(driver.find_element(By.NAME, 'input_14'))
    planned_dropdown.select_by_visible_text(row['Plan'])

    # Select Mode of Travel
    mode_dropdown = Select(driver.find_element(By.NAME, 'input_15'))
    mode_dropdown.select_by_visible_text(row['Mode'])

    # Select the Group option
    group_dropdown = Select(driver.find_element(By.NAME, 'input_16'))
    group_dropdown.select_by_visible_text(row['Group'])

    # Fill the expenditure fields
    driver.find_element(By.NAME, 'input_27').send_keys(row['Travel'])
    driver.find_element(By.NAME, 'input_28').send_keys(row['Food'])
    driver.find_element(By.NAME, 'input_29').send_keys(row['ReligiousItems'])
    driver.find_element(By.NAME, 'input_30').send_keys(row['Recreation'])
    driver.find_element(By.NAME, 'input_31').send_keys(row['Shopping'])
    driver.find_element(By.NAME, 'input_32').send_keys(row['Others'])

    # Optionally handle image upload if required
    try:
        upload_image = driver.find_element(By.XPATH, "//div[@class='moxie-shim moxie-shim-html5']//input")
        upload_image.send_keys(row['Image'])
    
        time.sleep(5)
    except:
        pass



    # WebdriverWait(driver, 60).until(
    # EC.presence_of_element_located((By.XPATH, "//div[@id='o_1ijjueun7j8vq3a19av1ho31e0df']//span[@class='gfield_fileupload_percent']"))
    # )
    # Submit the form
    try:
        submit_button = driver.find_element(By.ID, 'gform_submit_button_15')
        submit_button.click()
    except:
        pass

    # Wait for the form submission to complete (you may need to adjust this based on the site's behavior)
    # time.sleep(5)
    # WebdriverWait(driver, 15).until(EC.url_changes(driver.current_url))

    # Optionally handle navigation or refreshing to fill the next form entry
    a+=1
    print(a)
# Close the browser
driver.quit()
