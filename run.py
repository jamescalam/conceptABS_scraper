# -*- coding: utf-8 -*-
"""
Created on Thu Apr  4 11:13:11 2019

@author: jamesbriggs
"""

import conceptabs
"""
# initialise the web scraping object
scrape = conceptabs.webscrape()
# normal download of data (from last run-date to current date)
filename = scrape.download(1)

# initialise the data object
d = conceptabs.data()
# add the most recently downloaded data to the larger datafile (data.csv)
d.add(filename)
"""
# initialise the visualisation object
v = conceptabs.visualise()
# generate a basic HTML report
v.html()