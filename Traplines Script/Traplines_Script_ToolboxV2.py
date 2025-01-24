'''
Trapline_Script.py 
Description: 
This script creates a trapline boundary feature layer and KML by way of user input as "feature_name', along with any associated Crown Land File(s) with 
Trapline Cabin features  WHSE_TANTALIS.TA_CROWN_TENURES_SVW Definition Query = TENURE_SUBPURPOSE = 'TRAPLINE CABIN'
Storing these Features as .shp in the data_dir. An APRX file is stored in the APRX_dir and a PDF is created in the pdf_dir. 
Description:  Automate creation of Trapline boundary data, aprx, kml and pdf folders.
Author:  Ozra (Sunny) Rahimi, Evan Breton
         Ministry of Forests, Lands, Natural Resource Operations
           and Rural Development
           
Edited by: Chris Sostad Dec 16th 2024
      
Usage Note:  This script will only work with Temp_Trapline_Master template found here:
\\spatialfiles.bcgov\Work\srm\nel\Local\Geomatics\Workarea\SharedWork\Trapline_Territories\aprx\Temp_Trapline_Master.aprx

Version 1.0
'''

import arcpy
import os
from dotenv import load_dotenv

# addded .shp to trapline boundaries export
# moved the assignment of layout and maps to objects to the top of the script
# changed variable names so it is easier for others (namely me) to read

# Create or get the feature_layer object from ArcGIS Pro's content pane or the appropriate source
aprx = arcpy.mp.ArcGISProject("CURRENT")

arcpy.env.overwriteOutput = False

# Define paths for the directories 
# arcpy.env.workspace = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\SharedWork\Trapline_Territories'
# workspace = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\SharedWork\Trapline_Territories'

# Get the hidden file path from the .env file
workspace = os.getenv('WORKSPACE_PATH') # File path 

# Set the workspace to the path
arcpy.env.workspace = workspace

kml_dir = os.path.join(workspace, 'Kml')
aprx_dir = os.path.join(workspace, 'Aprx')
data_dir = os.path.join(workspace, 'Data')
pdf_dir = os.path.join(workspace, 'pdf')

# Convert main maps and layers to objects early in the script so when others are working in the script, they can scroll to the top and confirm it has been assigned rather
# than having to search through the script to find where it was assigned.

#NOTE: 

map_obj = aprx.listMaps('Map')[0]
layout = aprx.listLayouts("Layout")[0] 
all_trapline_cabins_obj = map_obj.listLayers("All Trapline Cabins")[0]
all_trapline_boundaries_obj = map_obj.listLayers("All Trapline Boundaries")[0]


# Function to create a directory if it doesn't exist
def create_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)
        arcpy.AddMessage(f"Directory created: {directory}")
    else:
        arcpy.AddMessage(f"Directory already exists: {directory}")

# Create the main directories
create_directory(aprx_dir)
create_directory(data_dir)
create_directory(pdf_dir)

#NOTE - change to user input

# Define the feature name (extracted from the query)
#feature_name = 'TR0440T001'

#set the parameter 'feature_name' - User will be required to input the trapline boundary #
feature_name = arcpy.GetParameterAsText(0)


# Create subdirectories named after the feature in each main directory
aprx_subdir = os.path.join(aprx_dir, str(feature_name))
data_subdir = os.path.join(data_dir, str(feature_name))
pdf_subdir = os.path.join(pdf_dir, str(feature_name))

create_directory(aprx_subdir)
create_directory(data_subdir)
#create_directory(pdf_subdir)


# Set up workspace and environment
arcpy.env.workspace = workspace
arcpy.env.overwriteOutput = True

################################################################################################################################
#
# Step 2 - Create the Application Polygon
#
#############################################################################################################################

arcpy.AddMessage("Step 2 - Creating Application Polygon")

# Set the definition query and assign it to a variable

expression = f"{arcpy.AddFieldDelimiters(arcpy.env.workspace, 'TRAPLINE_1')} = '{feature_name}'"

# Apply the expression to the layer "Application" in the Site Map (Defined as a global variable at the top of the script)
all_trapline_boundaries_obj.definitionQuery= expression
arcpy.AddMessage(f"         Definition query {expression} set for all trapline boundaries")

# Error Handling -  Use a SearchCursor to count the number of records returned by the definition query
count = 0
with arcpy.da.SearchCursor(all_trapline_boundaries_obj, "*") as cursor:
    for row in cursor:
        count += 1
        del cursor
# Check if the count is 0 and display a message if the definition query was successful or not.
if count == 0:
    
    arcpy.AddMessage(f"         No records returned by definition query: {expression}, please check the file number and try again")
else:
    arcpy.AddMessage(f"         Definition query {expression} set for layer: {all_trapline_boundaries_obj}. Records returned: {count}")


# Specify the trapline boundary layer you want to export features from
application_trapline_boundary_path = os.path.join(data_dir, 'Scripting Data', 'Trapline_Boundaries_Export.shp')
arcpy.AddMessage(f"Trapline boundary path: {application_trapline_boundary_path}")

# Specify the output file path for the exported feature
application_trapline_boundary = os.path.join(data_subdir, f'{feature_name}.shp')
arcpy.AddMessage(f"Output feature path: {application_trapline_boundary}")

###########################################################################################################################################
#
# Step 3 - Export the Feature to a Shapefile
#
###########################################################################################################################################

# Create a feature layer to select the specific feature
try:
   arcpy.management.MakeFeatureLayer(all_trapline_boundaries_obj, "temp_layer") 
   # Make feature layer always creates a temporary layer. Its held in memory and will be deleted upon exit.  
   arcpy.AddMessage("Feature layer created.")
except arcpy.ExecuteError as e:
    arcpy.AddMessage("MakeFeatureLayer_management error:", e)
    exit()  # Exit the script if there's an error

# Check if the temporary layer exists
if arcpy.Exists("temp_layer"):
    try:
        
        arcpy.management.CopyFeatures("temp_layer", application_trapline_boundary) # Layer has been created (No longer just a path)
        arcpy.AddMessage("Export process complete.")
    except arcpy.ExecuteError as e:
        arcpy.AddMessage("CopyFeatures_management error:", e)
else:
    arcpy.AddMessage("Temporary layer 'temp_layer' does not exist.")
    exit()

# Add a new field for the area in hectares if it doesn't already exist
area_field = "Area_ha"
if area_field not in [f.name for f in arcpy.ListFields(application_trapline_boundary)]:
    arcpy.management.AddField(application_trapline_boundary, area_field, "DOUBLE")
    arcpy.AddMessage(f"Field '{area_field}' added to the feature class.")

# Define the function to calculate area in hectares
def calculate_area_in_hectares(geometry):
    arcpy.AddMessage("Defining calculate_area_in_hectares function...")
    
    area_sq_meters = geometry.area  # Area in square meters
    return area_sq_meters / 10000  # Convert to hectares

# Calculate the area for each polygon and update the new field
arcpy.AddMessage("Calculating area in hectares...")
with arcpy.da.UpdateCursor(application_trapline_boundary, ["SHAPE@", area_field]) as cursor:
    for row in cursor:
        area_hectares = calculate_area_in_hectares(row[0])
        row[1] = area_hectares  # Update the field with the area in hectares
        cursor.updateRow(row)
        del cursor  # Delete the cursor after it has completed its task
        arcpy.AddMessage("Deleted cursor")
        # Format the area to two decimal points
        formatted_area = f"{area_hectares:.2f} ha."
        arcpy.AddMessage(f"Formatted area: {formatted_area}")

arcpy.AddMessage(f"Area in hectares has been added to the field '{area_field}'.")

##############################################################################################################
#
# Step 4 - Clip the "Trapline Cabins" layer based on the feature layer
#
##############################################################################################################

# Specify the "Trapline Cabins" layer name
#NOTE used the all cabins layer in contents pane instead of the path

# all_trapline_cabins_layer = os.path.join(data_dir, 'TRAPLINE_CABINS_Polygon.shp')
clipped_cabins_output = os.path.join(data_subdir, f'{feature_name}_Cabins.shp')
arcpy.AddMessage
arcpy.AddMessage(f"Output feature path: {clipped_cabins_output}")
# Clip the "Trapline Cabins" layer based on the feature layer\

arcpy.analysis.Clip("All Trapline Cabins", application_trapline_boundary, clipped_cabins_output)  # After the clip, the clipped cabins output path has now become a layer

arcpy.AddMessage(f"Clipping of trapline boundary to Crown Lands layer completed.") 

#working on this portion to deal with files that do not have a crown land file number associated, single and multi Cabnin Crown Land File #'s also need to be managed.

crown_land_field = "CROWN_LAND"  # Assuming this field exists

# Check if the field exists in the attribute table of the clipped cabins output
if crown_land_field not in [f.name for f in arcpy.ListFields(clipped_cabins_output)]:
    arcpy.AddMessage(f"{crown_land_field} not found in the attribute table.")
else:
    # Initialize an empty list to store Crown Land values
    Crown_Num_Values = []

    # Open an UpdateCursor to iterate over rows in the clipped cabins feature class
    with arcpy.da.UpdateCursor(clipped_cabins_output, [crown_land_field]) as cursor:
        for row in cursor:
            # Capture the Crown Land value from the crown_land_field
            crown_land_value = row[0]

            # Only append non-null and non-empty values to the list
            if crown_land_value is not None and crown_land_value != "":
                Crown_Num_Values.append(crown_land_value)

    # Check if any valid Crown Land values were found
    if not Crown_Num_Values:
        arcpy.AddMessage("No valid Crown Land values found. Setting default layer name.")
        # Set a default name if no values found
        new_layer_name = "Trapline_Cabin_No_Values"
        all_trapline_cabins_obj.name = new_layer_name
        arcpy.AddMessage(f"Layer renamed to: {new_layer_name}")
    else:
        # Convert the list of Crown Land values to a single string, joined by underscores
        Crown_Num_Values_String = "_".join(map(str, Crown_Num_Values))
        
        # Update the name of the trapline cabin feature layer based on the joined Crown Land values
        new_layer_name = f"Trapline_Cabin_{Crown_Num_Values_String}"
        all_trapline_cabins_obj.name = new_layer_name
        arcpy.AddMessage(f"Layer renamed to: {new_layer_name}")
    
    # Apply definition query based on the number of Crown Land values found
    if len(Crown_Num_Values) == 1:
        # If there's only one Crown Land value, apply an equality query
        expression1 = f"{arcpy.AddFieldDelimiters(arcpy.env.workspace, 'CROWN_LAND')} = '{Crown_Num_Values[0]}'"
        all_trapline_cabins_obj.definitionQuery = expression1
        arcpy.AddMessage(f"Definition query applied to Layer1: {expression1}")
    else:
        # If there are multiple Crown Land values, apply an IN query
        values_string = ', '.join([f"'{val}'" for val in Crown_Num_Values])  # Prepare values for IN clause
        expression1 = f"{arcpy.AddFieldDelimiters(arcpy.env.workspace, 'CROWN_LAND')} IN ({values_string})"
        all_trapline_cabins_obj.definitionQuery = expression1
        arcpy.AddMessage(f"Definition query applied to Layer1: {expression1}")

# Rename the boundaries layer
all_trapline_boundaries_obj.name = f"{feature_name} ({formatted_area})"
arcpy.AddMessage(f"Layer renamed to: {feature_name} ({formatted_area})")

# Create a new variable for Crown cabins string (for further operations if needed)
new_crown_cabins_str = f"{feature_name}_Cabins_{Crown_Num_Values_String}" if Crown_Num_Values else "Trapline_Cabin_No_Values"
arcpy.AddMessage(f"New Crown cabins variable: {new_crown_cabins_str} created.")

#Define the path to both of the Feature Layers ( Trapline Boundary and Trapline Cabins )
trapline_cabin_fc_path = clipped_cabins_output
trapline_bnd_fc_path = application_trapline_boundary

new_trapline_bnd_layer = arcpy.MakeFeatureLayer_management(trapline_bnd_fc_path, f'{feature_name}.shp')

new_trapline_cabin_layer = arcpy.MakeFeatureLayer_management(trapline_cabin_fc_path, f'{feature_name}_Cabins.shp')

# Find the text element (Map Title) and update its text to the feature name as:(Trapline TR0###T###)
text_updated = False

# for lyt in aprx.listLayouts("Layout"):
for elm in layout.listElements("TEXT_ELEMENT"):
    if elm.name == "Map Title":
        elm.text = f"Trapline {feature_name}"
        arcpy.AddMessage("\tProponent text element changed")

#try to rename the original feature layers for all cabin and all trapline boundaries as they are already symbolized correctly..
layer_name1 = all_trapline_cabins_obj
new_name1 = f"{new_crown_cabins_str}"
layer_name2 = all_trapline_boundaries_obj
new_name2 = f"{feature_name} ({formatted_area})"


#Apply a definition query using arcpy.AddFieldDelimiters
expression = f"{arcpy.AddFieldDelimiters(arcpy.env.workspace, 'TRAPLINE_1')} = '{feature_name}'"
layer_name2.definitionQuery = expression
arcpy.AddMessage(f"Definition query applied to {feature_name} = {expression}")

#zoom to Trapline Boundary feature and round scale to 250,000

#set the name of the zoom layer created earlier in the script "new_boundaries_layer_name = f"{feature_name} ({formatted_area})""
zoom_feature_layer =  all_trapline_boundaries_obj

#get the layout and map frame 
lyt = aprx.listLayouts()[0]

map_obj = lyt.listElements('MAPFRAME_ELEMENT', 'Map Frame')[0]

# use the zoom_feature_layer for zooming
arcpy.SelectLayerByAttribute_management(zoom_feature_layer, "NEW_SELECTION", "1=1")

# zoom to selected featues within the map fram using the zoom feature layer = new_boundaries_layer_name
map_obj.zoomToAllLayers(True)
#clear selection
arcpy.SelectLayerByAttribute_management(zoom_feature_layer, "CLEAR_SELECTION")

map_obj.camera.scale = 250000
arcpy.AddMessage(f"Zoomed to feature in {zoom_feature_layer} and set scale to : {map_obj.camera.scale}")

# Refresh the map frame to reflect the symbology changes
map_obj.camera.setExtent(map_obj.camera.getExtent())  # Forces a redraw by resetting the map extent

# Save aprx for trapline boundary map automation project to preserve changes and allow users to open the aprx to view map and make adjustments if required.
aprx.saveACopy(os.path.join(aprx_subdir, f'{feature_name}.aprx'))
arcpy.AddMessage(f"ArcPro Project {feature_name}.aprx has been saved successfully here: {aprx_subdir}")

arcpy.AddMessage("----------------------------------------------------")
arcpy.AddMessage("----------------------------------------------------")
arcpy.AddMessage("Trapline Boundary Map Automation has been completed")
arcpy.AddMessage("----------------------------------------------------")
arcpy.AddMessage("----------------------------------------------------")