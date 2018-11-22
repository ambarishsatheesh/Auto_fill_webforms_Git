# This program opens a specific web page using a webdriver and then logs in to an account using a given username and password. It then opens up a form page
# to input user data into some fields.
# The data consists of names, emails and contact details, all of which are unique for each user.
# This data is read from a table in a spreadsheet and stored in a pandas DataFrame before being converted into lists.
# The program loops through the lists and inputs what was originally a row in the spreadsheet into the fields on the web page.
# The program clicks checkboxes and submits the entry. This entry process loops through the lists until the complete range specified has been entered.

 selenium import webdriver
import config
import time
import pandas as pd

# This method of username/password storage allows for hiding of sensitive data from anyone who views/edits this code
username = config.DATACOUP_USERNAME
password = config.DATACOUP_PASSWORD

# Instantiate chrome driver
driver = webdriver.Chrome(executable_path=r"file_path_here")

# Open relevant web page
driver.get("target_url_here")

# Wait (logic here is to attempt to circumvent auto bot-flagging mechanisms but not sure how effective this is)
time.sleep(2)

# Enter username (inspect element in browser to find ID)
elem = driver.find_element_by_id("ID_here")
elem.send_keys(username)

# Enter password
elem = driver.find_element_by_id("ID_here")
elem.send_keys(password)

# Wait
time.sleep(2)

# CLick "Keep me logged in" (inspect element in browser to find xpath)
elem = driver.find_element_by_xpath("xpath_here").click()

# Wait
time.sleep(1)

# CLick "Log in"
elem = driver.find_element_by_xpath("xpath_here").click()

# Wait
time.sleep(2)

# Open Excel and read cells
df = pd.read_excel('file_name.xlsx', sheet_name='List', usecols="A:G")

# Input column data into lists
email = df['Email Address'].tolist()
f_name = df['First Name'].tolist()
l_name = df['Last Name'].tolist()
c_name = df['Company Name'].tolist()
t_no = df['Telephone No'].tolist()
county = df['County'].tolist()
town = df['Town'].tolist()

# CLick "Add Subscriber"
elem = driver.find_element_by_xpath("xpath_here").click()

# Wait (some of these are used to account for browser slowdown)
time.sleep(4)

# CLick checkbox
elem = driver.find_element_by_xpath("xpath_here").click()

# Wait
time.sleep(1)

# CLick checkbox
elem = driver.find_element_by_xpath("xpath_here").click()

# Wait
time.sleep(4)

# Loop through stored lists and input to form
for i in range(start, end):

    # Type Email Address
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(email[i])

    # Type First Name
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(f_name[i])

    # Type Last Name
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(l_name[i])

    # Type Company Name
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(c_name[i])

    # Type Telephone Number
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(t_no[i])

    # Type County
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(county[i])

    # Type Town
    elem = driver.find_element_by_id("ID_here")
    elem.send_keys(town[i])

    # Uncheck checkbox (to account for any auto-uncheck while loop is running)
    elem = driver.find_element_by_xpath("xpath_here").click()

    # Wait
    time.sleep(1)

    # CLick checkbox again (to account any for auto-uncheck while loop is running)
    elem = driver.find_element_by_xpath("xpath_here").click()

    # Wait
    time.sleep(1)

    # Uncheck checkbox (to account for any auto-uncheck while loop is running)
    elem = driver.find_element_by_xpath("xpath_here").click()

    # Wait
    time.sleep(1)

    # CLick checkbox again (to account any for auto-uncheck while loop is running)
    elem = driver.find_element_by_xpath("xpath_here").click()

    # Wait
    time.sleep(1)

    # CLick "Submit"
    elem = driver.find_element_by_xpath("xpath_here").click()

    # Wait
    time.sleep(2)

# Close Driver
driver.close()


