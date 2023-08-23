import pandas as pd
import numpy as np
import datetime as dt
import glob
import xlsxwriter
import sys


## Used throughout script
na_vals = ["NA", "Missing"] ###############
benefitCode = {            # "BenefitCode" : "Coverage(Lookup)Code"
    "ADDL1": "AD+D",
    "ADDL2": "AD+D",
    "ADDL3": "AD+D",
    "ADDL4": "AD+D",
    "ADDLE": "V AD+D",
    "LIFLC": "VC LIFE",
    "LIFLE": "V LIFE",
    "LIFLS": "VS LIFE",
    "LIFL1": "LIFE",
    "LIFL2": "LIFE",
    "LIFL3": "LIFE",
    "LIFL4": "LIFE",
    "LTDL": "LTD",
    "ADDLS": "VS AD+D"
}
# glEEcode = "2404"      // not sure where to find these, are they relevant for lincoln???
# glERCobra = "1405"
glERcode = "7102"        ##############
glCompany = "2003"       ##############
benefit = 'Lincoln'      ##############
monthOVRD = False
monthYearOVRD = '07-2023'

## Used to get current Month+Year
if monthOVRD == False:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)
    strMonthYear = str(dtToday.month).rjust(2, "0") + "-" + str(dtToday.year)
else:
    dtToday = dt.datetime.today()
    strToday = str(dtToday.month) + "/" + str(dtToday.day) + "/" + str(dtToday.year)    
    strMonthYear = monthYearOVRD