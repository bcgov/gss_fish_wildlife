** Press Ctrl+Shift+V to view Markdown Text

# Problem #1 - "Clipping" Data in Related Tables
You have a shapefile of lakes in the Omineca Region called Management Objectives. It was created by using the Omineca Wildlife Habitat Areas clipping the freshwater atlas layer. (For this exercies it is "Join Table")
This shapefile has 7 tables related to it. The tables are all related to the Management Objectives with Waterbody ID as the primary key. 
Unfortunately, all of the tables contain data not just for Omineca, but the whole Province of BC. The client wants you to remove all data that is not in the Omineca region so that the related tables only contain the data for Omineca. If all the data was spatialized, you could use the Omineca Polygon to clip the related table data. However, the data isn't spatialized so you will need to use python to clean the data.

Hint** The Join Table (aka our shapefile) is already clipped to Omineca. Therefore, every waterbody Id in the join table is an Omineca Waterbody

Tables are here: (Make a copy in your working drive)
 \\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\csostad\FishDataWebmap\ExcelSheets\Cleaned_Sheets_Master - Copy
