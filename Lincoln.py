import pandas as pd
import numpy as np
import datetime as dt
import glob
import xlsxwriter
import sys


## Used throughout script
na_vals = ["NA", "Missing"] ###############
benefitCode = {             ###############
    "ADDL1": None,
    "ADDL2": None,
    "ADDL3": None,
    "ADDL4": None,
    "ADDLE": None,
    "LIFLC": None,
    "LIFLE": None,
    "LIFLS": None,
    "LIFL1": None,
    "LIFL2": None,
    "LIFL3": None,
    "LIFL4": None,
    "LTDL": None,
}
# glEEcode = "2404"      // not sure where to find these, are they relevant for lincoln???
# glERCobra = "1405"
glERcode = "7102"        ##############
glCompany = "2003"       ##############
benefit = 'Lincoln'      ##############
monthOVRD = False
monthYearOVRD = '07-2023'