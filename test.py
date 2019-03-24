# -*- coding: utf-8 -*-
"""
Created on Sun Mar 24 09:21:45 2019

@author: jamesbriggs
"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time

sleeptime = 1

driver = webdriver.Chrome('drivers/chromedriver.exe')

driver.set_page_load_timeout(10)

driver.get('https://www.conceptabs.co.uk/default.aspx')

"""
Logging in
"""
# so we can track the number of login attempts and stop process if too high
attempt = 0
# putting into a while statement to repeat until we are allowed in
while True:
    attempt += 1
    print('Login attempt {}'.format(attempt))

    # giving page a moment to load
    time.sleep(sleeptime)
    # username
    driver.find_element_by_name('ctl00$CPH1$LoginView1$Login1$UserName').send_keys('DELTE5')
    
    # password
    driver.find_element_by_name('ctl00$CPH1$LoginView1$Login1$Password').send_keys('jamesr_5')
    
    # press login button
    driver.find_element_by_name('ctl00$CPH1$LoginView1$Login1$LoginImageButton').click()
    
    """
    Navigating to CDOs
    """
    # give the page a moment to load
    time.sleep(sleeptime)
    # at this point we may need to login again, so will use try-except statement
    try:
        # click the search database and graphing option
        driver.find_element_by_name('ctl00$CPH1$Heads$PicLink2').click()
        # if successful, break from the while-loop
        break
    except:
        # if we need to login again...
        driver.find_element_by_name('ctl00$CPH1$LoginView1$btnLoginRed').click()
        # stop process if attempt number too high
        if attempt > 9:
            driver.close()
            raise ValueError('Too many refused login attempts. Driver closed.')
        
# give page time to load
time.sleep(sleeptime)

# click the Type tab
driver.find_element_by_id('__tab_ctl00_CPH1_TC3_tpType').click()

# give page time to load
time.sleep(sleeptime)

# click the down button
dropdown = Select(driver.find_element_by_id('ctl00_CPH1_TC3_tpType_ddlType_ddlType_Table'))

dropdown.select_by_visible_text('CDO')

"""
otherwise
dropdown button id = ctl00_CPH1_TC3_tpType_ddlType_ddlType_Button
ul id = ctl00_CPH1_TC3_tpType_ddlType_ddlType_OptionList
inside = CDO
"""
input('continue?\n>>> ')
# finish
driver.close()