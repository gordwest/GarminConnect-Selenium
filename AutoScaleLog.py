#!/usr/bin/env python
# coding: utf-8
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
import time, openpyxl, os

#################
### Variables ###
#################

# Garmin Connect - Creds
webAddress = "https://connect.garmin.com/"
email = "xXXXXXXXXXXXXXXXXXXXXXx"
password = "xXXXXXXXXXXXXXXXXXXXx"

# Column Names and corresponding fields
# C - Weight
# D - Body Mass Index
# E - Body Fat %
# F - Skeletal Muscle Mass
# G - Water %
fields = ('Weight', 'BMI', 'BF', 'SMM', 'Water')
columns = ('C', 'D', 'E', 'F', 'G')

# Dictionary to hold scale values from garmin connect
weighIn = {'Weight' : '',
            'BMI' : '',
            'BF' : '',
            'SMM' : '',
            'Water' : ''
          }

# Dict of xpaths to text elements
GarminConXpaths = {'Weight' : """//*[@id="directWeightClickedTableRow"]""",
                    'BMI' : """//*[@id="directBMIClickedTableRow"]""",
                    'BF' : """//*[@id="directBodyFatClickedTableRow"]""",
                    'SMM' : """//*[@id="directMuscleMassClickedTableRow"]""",
                    'Water' : """//*[@id="directBodyWaterClickedTableRow"]"""
                  }

# Open excel workbook and load specific sheet
filename = 'xXXXXXXXXXXXXXx.xlsx'
wb = openpyxl.load_workbook(filename)
allSheetNames = wb.sheetnames
weightLog = wb[allSheetNames[0]]

# Save backfile before update as precaution
wb.save('xXXXXXXXXXXXXXXXXXx.xlsx')

#################
### Functions ###
#################

def getFirstEmptyRow(columnLetter):
    """ Iterate over rows in a column and return the row number of the first empty cell
    params
    ------
    columnLetter: string
        Name of the column in excel to use

    returns
    -------
    row: int
        Row number of first empty cell
    """
    for row in range(1, 1000):
        cell_name = "{}{}".format(columnLetter, row)
        if weightLog[cell_name].value is None:
            return row
            break
     

def insertValue(columnLetter, fieldValue):
    """ Insert value in specific cell in excel spreadsheet
    params
    ------
    columnLetter: string
        Name of the column in excel to use
    fieldValue: int
        Value to be entered into spreadsheet from 'weighIn' dict
    """ 
    cell_name = "{}{}".format(columnLetter, getFirstEmptyRow(columnLetter))
    print('Value: {} - Inserted into cell: {}{}'.format(fieldValue, columnLetter, getFirstEmptyRow(columnLetter)))
    weightLog[cell_name].value = fieldValue

    
def clickItem(xpathstr):
    """ Click on a web element using selenium
    params
    ------
    xpathstr: string
        xPath of the web element
    """ 
    item = driver.find_element_by_xpath(xpathstr)
    webdriver.ActionChains(driver).click(item).perform()
    
    
def pressKey(keyName):
    """ Press a key in the browser
    params
    ------
    keyName: string
        name of key to press (All caps - ex: TAB, ENTER)
    """   
    webdriver.ActionChains(driver).key_down(keyName).perform()
    webdriver.ActionChains(driver).key_up(keyName).perform()

    
def getText(xpathstr):
    """ Copy text web element 
    params
    ------
    xpathstr: string
        xPath of the web element

    returns
    -------
    text object from web page
    """ 
    return driver.find_element_by_xpath(xpathstr).text


################
### Seleinum ###
################

# Initialize browser instance
driver = webdriver.Chrome()
driver.get(webAddress)
time.sleep(15)

# Click login button
clickItem("""//*[@id="___gatsby"]/div/div/header/nav/ul/li[4]/a/button""")
time.sleep(10)

# Enter email - press tab
webdriver.ActionChains(driver).send_keys(email).key_down(Keys.TAB).key_up(Keys.TAB).perform()
time.sleep(1)

# Enter password - press enter
webdriver.ActionChains(driver).send_keys(password).key_down(Keys.ENTER).key_up(Keys.ENTER).perform()
time.sleep(10)

# Click ' Health Stats dropdown'
clickItem("""/html/body/div[1]/nav/div/ul[2]/li[2]/a/span""")
time.sleep(2)

# Click ' Weight tab'
clickItem("""/html/body/div[1]/nav/div/ul[2]/li[2]/ul/li[2]/a""")
time.sleep(10)

# Get text values using xPaths, convert to float and store in 'weighIn' dict
for f in fields:
    weighIn[f] = float(getText(GarminConXpaths[f]))


################
### Openpyxl ###
################

# Insert values into worksheet 
for i in range(len(columns)):
        insertValue(columns[i], weighIn[fields[i]])
        
# Update workbook
wb.save(filename)
print("Excel file has been successfully updated with today's weigh in!")