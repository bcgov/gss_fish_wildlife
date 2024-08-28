""""
This script creates a new folder with year on it in the WILDLIFE Authorization folder
Paired with jenkins runs yearly
2024-08-27
"""
#importing libraries
import os
from datetime import date
#path of the folder 
base_dir = r"\\spatialfiles.bcgov\work\lwbc\nsr\Workarea\fcbc_fsj\WILDLIFE"

#year and path variables
today = date.today()
year= str(today.year)
path_year = os.path.join(base_dir,year)
#does the folder exist
if not os.path.exists(path_year): 
    #folder doesn't exist, create folder
    os.makedirs(path_year)
#folder exist, tell user
else: 
    print("this year's folder already exist")