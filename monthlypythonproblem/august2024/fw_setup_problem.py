'''
NET_FW_Setup.py 
Description: 
This script is a hybrid of
NETSetup_Pro.py
Description:  Automate creation of folders and shapefiles for Northeast
              tenure processing (Crown Lands, Water, Wildlife, Forests).
Author:  Mike Eastwood
         Ministry of Forests, Lands, Natural Resource Operations
           and Rural Development
         250-774-5502
         michael.eastwood@gov.bc.ca
Usage Note:  This script will only work on Arc 10.2 or later.
Version 2.2
  1.0 - original version
  1.1 - added Mines tenures
  1.2 - added Compatible Use Crown Lands tenures
  2.0 - ArcGIS Pro version
  2.1 - Update to match new org and folder structures
  2.2 - Added user options for various tasks
  
A break from the original happened in 2024-Apr as apart of the re-og Geospatial Services. 
As a result of the break we are calling this Version 3. 
'''
# ===========================================================================
# Preamble stuff - modules, environment, parameters and variables
# ===========================================================================
# Import python modules
import arcpy
import datetime
import os
import os.path
import sys
import arcpy.management
# Script arguments
file = arcpy.GetParameterAsText(0)          #Tenure Number #keep
append_aoi = arcpy.GetParameterAsText(1)    #Area of Interest to be Appended
arcpy.env.workspace = r"\\spatialfiles\work\lwbc\nsr\Workarea\fcbc_fsj\Wildlife"
arcpy.env.overwriteOutput = False
# Calculate date variables
date = datetime.date.today()
year = str(date.year)
mon = date.month
if mon == 1:
    month = "Jan"
elif mon == 2:
    month = "Feb"
elif mon == 3:
    month = "Mar"
elif mon == 4:
    month = "Apr"
elif mon == 5:
    month = "May"
elif mon == 6:
    month = "Jun"
elif mon == 7:
    month = "Jul"
elif mon == 8:
    month = "Aug"
elif mon == 9:
    month = "Sep"
elif mon == 10:
    month = "Oct"
elif mon == 11:
    month = "Nov"
elif mon == 12:
    month = "Dec"
# Set variables
base = arcpy.env.workspace
baseYear = os.path.join(base, year)
outName = file.upper()
geometry = "POLYGON"
template = r"\\spatialfiles.bcgov\Work\lwbc\nsr\Workarea\fcbc_fsj\Templates\BLANK_polygon.shp"
m = "SAME_AS_TEMPLATE"
z = "SAME_AS_TEMPLATE"
spatialReference = arcpy.Describe(template).spatialReference
# ===========================================================================
# Create Folders
# ===========================================================================
arcpy.AddMessage("  ")
arcpy.AddMessage("Creating folders . . .")
outName = file.upper()
#path to folder location
fileFolder = os.path.join(baseYear, outName)
shapeFolder = fileFolder
outPath = shapeFolder
if os.path.exists(fileFolder):
    arcpy.AddMessage(outName + " folder already exists.")
else:
    os.mkdir(fileFolder)
# ===========================================================================
# ===========================================================================
# Create Shapefile(s) and add them to the current map
# ===========================================================================
arcpy.AddMessage("  ")
arcpy.AddMessage("Creating Shapefiles . . .")
if os.path.isfile(os.path.join(outPath, outName + ".shp")) == True:
    arcpy.AddMessage(os.path.join(outPath, outName + ".shp") + " already exists")
    arcpy.AddMessage("Exiting without creating files")
    sys.exit()
else:
   create_shp = arcpy.management.CreateFeatureclass(outPath, outName, geometry, template, m, z, spatialReference)
   k_layer = arcpy.management.MakeFeatureLayer(create_shp,"area_of_interest")
   create_kml = os.path.join( outPath, outName + ".kml")
   arcpy.conversion.LayerToKML(k_layer,create_kml)
   arcpy.AddMessage("kml created: " + create_kml)
   arcpy.management.Append(append_aoi,create_shp,"NO_TEST")
   arcpy.AddMessage("Append Successful")
#adding data to the map
aprx = arcpy.mp.ArcGISProject("CURRENT")
aprxMap = aprx.activeMap
if aprxMap == None:
    arcpy.AddMessage("No map to add the data to.")
else:
    arcpy.AddMessage("Adding data to the active map.")
    output = os.path.join(outPath, outName + ".shp")
    aprxMap.addDataFromPath(output)
# ===========================================================================
# End script politely
# ===========================================================================
arcpy.AddMessage("  ")
arcpy.AddMessage("  ")
arcpy.AddMessage("===========================================================================")
arcpy.AddMessage(fileFolder + ", is ready for processing.")
arcpy.AddMessage("===========================================================================")
arcpy.AddMessage("  ")
arcpy.AddMessage("  ")
# ===========================================================================
