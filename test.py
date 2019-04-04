# -*- coding: utf-8 -*-
"""
Created on Sun Mar 24 09:21:45 2019

@author: jamesbriggs
"""

from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
from datetime import datetime, timedelta
import os

"""
Setup
"""
# how long the script will wait when loading pages
sleeptime = 1

# getting the current date in the required format
# eg = dd MMM yyyy (ex 01 Mar 2019)

current_date = datetime.strftime(datetime.today(), '%d %b %Y')

# and for the previous week
prior_date = datetime.strftime(datetime.today() - timedelta(days=7), '%d %b %Y')
"""
current_date = '27 Mar 2019'
prior_date = '01 Mar 2019'
"""
# show user what date range we will be using
print('Pulling data from {} to {}.'.format(prior_date, current_date))

# setting up Selenium web driver
driver = webdriver.Chrome('drivers/chromedriver.exe')

# opening conceptABS webpage
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

# click the Date tab
driver.find_element_by_id('__tab_ctl00_CPH1_TC3_tpDate').click()

time.sleep(sleeptime)
    
# select first box (select current.date - 1 week)
driver.find_element_by_name('ctl00$CPH1$TC3$tpDate$Date1').send_keys(prior_date)

# select second box (select current.date)
driver.find_element_by_id('ctl00_CPH1_TC3_tpDate_Date2').click()
time.sleep(sleeptime)
driver.find_element_by_id('ctl00_CPH1_TC3_tpDate_Date2').send_keys(current_date)

attempt = 0

while True:
    attempt += 1
    print('Pulling data attempt {}'.format(attempt))
    time.sleep(sleeptime)

    # select + add filter
    driver.find_element_by_name('ctl00$CPH1$TC3$tpDate$btnDateAdd').click()
    time.sleep(sleeptime)
    
    # press search
    driver.find_element_by_name('ctl00$CPH1$btnSearch').click()
    time.sleep(sleeptime)
    
    # let page load
    time.sleep(sleeptime)
    
    # get list of downloads folder before download
    pre_dl = os.listdir(r'C:\Users\jamesbriggs\Downloads')
    
    try:
        # select tranches
        driver.find_element_by_name('ctl00$CPH1$btnExportTranches').click()
        break
    except:
        if attempt > 9:
            driver.close()
            raise ValueError('Too many failed data pull attempts. Driver closed.')
        pass

# wait 5 seconds to allow download
time.sleep(5)

# get list of downloads folder before download
post_dl = os.listdir(r'C:\Users\jamesbriggs\Downloads')

# find the difference between pre and post DL lists
new_file = list(set(post_dl) - set(pre_dl))

if len(new_file) == 1:
    print('New files:\n{}'.format(new_file[0]))
    
    # moving file
    os.rename(r'C:\Users\jamesbriggs\Downloads\{}'.format(new_file[0]),
              r'C:\Users\jamesbriggs\Documents\Web Scraping\ConceptABS\data\{}'.format(new_file[0]))
    
else:
    raise IOError('No new files downloaded to Downloads directory.')

# finish
driver.close()