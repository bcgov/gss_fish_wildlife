"""

Author: Amanda Lu
Ministry, Division, Branch: WLRS- GEOBC
Created Date: 2024-08-01
Updated Date: 2024-08-01

Description: This script is designed to delete transitory data in the Wildlife Authorization Process. 
Specifically data that is 6 months or older on today's date

Dependencies:
base_dir which is the file directory that the transitory data sits on. 
the base_dir having year as the folder name 

"""
#%%
# Import library
import os
import shutil
from datetime import date
from dateutil import relativedelta
from datetime import datetime

#location of the base directory
base_dir = r'\\spatialfiles.bcgov\Work\lwbc\nsr\Workarea\fcbc_fsj\WILDLIFE'

#date variables
today = date.today()
this_year = str(today.year)
this_month = today.month
path_year = os.path.join(base_dir,this_year) 

#have python walk through the files in the directory created 
for root,dir,files in os.walk(path_year):
        #thisPath = os.path.join(root,name)
        #only pick gdb and "mapx files"
        if root.endswith(".gdb") or 'mapx_files' in root:
#           we are going to isolate the selected files base on time 
            m_time = os.path.getmtime(root)
            dt_m = datetime.fromtimestamp(m_time)
            #getting late of last modification 
            r = relativedelta.relativedelta(today,dt_m)
            months_diff = (r.years*12)+r.months
            #pick only files that are 6 months or older
            if months_diff > 6: 
                  print(root,'has been removed')
                  #remove the file
                  shutil.rmtree(root)


              