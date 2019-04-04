# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 10:22:25 2019

@author: jamesbriggs
"""
import os
import json
import time
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import Select


class webscrape:
    
    def __init__(self):
        """
        Here will find when last web scrape was conducted.
        """
        detail_loc = r'properties\details.json'
        if not os.path.isfile(detail_loc):
            # if no details JSON object exists, create new using earliest
            # start date
            self.details = {
                    'previous': '27 Mar 2019'
                    }
            # saving to details.json
            with open(detail_loc, 'w') as fp:
                json.dump(self.details, fp)

        else:
            # if details JSON object does exist, just load
            with open(detail_loc, 'r') as fp:
                self.details = json.load(fp)

    def download(self):
        """
        Here we will pull all data from the previous web scrape to the current
        date.
        """
        # set sleeptime here (time required to let webpages load)
        sleeptime = 1

        # getting the current date in the required format
        # eg = dd MMM yyyy (ex 01 Mar 2019)
        current_date = datetime.strftime(datetime.today(), '%d %b %Y')
        
        # and for the previous week
        prior_date = self.details['previous']

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

    def scrape_all(self):
        """
        Here will scrape all data from the set starting date to now. This
        requires the data directory to contain no xlsx files.
        """
        pass

    def save_scrape_info(self):
        """
        Here we will save the scrape info to the detail.json file within
        the properties directory.
        """
        pass


class results:
    
    def __init__(self):
        pass

    def add(self, filename):
        """
        Clean the given filename CABS search results file and append to the
        combined datatable in the data\full directory.
        """
        # set the target file
        target = r'data\{}.xlsx'.format(filename)
        # first we will remove the top 5 rows in the CABS file so we are left
        # with only headings and data
        print(r'Loading .\{}'.format(target))
        # loading the workbook
        wb = load_workbook(filename=target)
        # selecting the worksheet
        ws = wb.active
        # before deleting the top 5 rows, we must check that they are empty
        formatted = False
        for i in range(1, 6):
            print(i)
            if ws['C{}'.format(i)].value != None:
                print('Sheet already formatted.')
                formatted = True
                break

        # if the sheet is not formatted, we will format
        if not formatted:
            print('Formatting sheet.')
            # deleting top 5 rows
            ws.delete_rows(1, 5)
            # saving modified workbook
            wb.save(target)

        # now we can pull the formatted data into Pandas
        new = pd.read_excel(target)


    def import_all(self):
        """
        Will use this to pull in all data within the data directory.
        """
        pass