#===========================================================================
# Script name: Site and Overview Toolbox
# Author: Initial scripts created by Jordan Foy
# Rewritten/Refactored by Chris Sostad
# Created on: 08/03/2021
# Edited on: 12/12/2023
# 
#
# Description: This script will allow a user to automate the lands authorization process.
# There will be separate tools within the toolbox to accomplish different land authorization maps. 
# 
#
# 
#============================================================================

#import all necessary libraries 
import arcpy 
import traceback 
import sys 
import os 
import datetime
import openpyxl as xl




## Dec 02,2023 
    # Updated many of the connection properties scripts to use the update_layer_connection function.
    # Added error handling for the centroid DMS formatting function.
    # Need to fix the WHSE layer sometimes failing. It might only be running if the correct layout is open.

# Dec 04, 2023
    # Split lines was not updating properly. It is replacing the splitline line file with a polygon
    #  newConnPropDict = {'connection_info': {'database': lands_file_path + str(file_num)},
    #                                 'dataset': f"{file_num}.shp", **Should be layer_name.shp
    #                                 'workspace_factory': 'Shape File'}
    # Fixed the update connection properties function as above
    # Cut and pasted the working copy of Create Full Site map and overview and export tool from its 
    # individual tool into the Lands_Testing Toolbox and tested. Successfully.
    # Still need to test the edits made to the tool SiteAndOverviewToolbox TESTING

# Dec 05, 2023
    # Fixed application layer not updating properly. It was because the application layer shapefile is actually called the file number, 
    # not application. So I changed the update_layer_connection function to use the file number instead of application.
    # Added functionality to create a new folder if the file number already exists. It will create a new folder with the month 
    # and year appended to the folder name.


# Dec 06, 2023
    # Brought in export layout with manual inset from Lands_Testing 1.1 Toolbox
    # Developed in SiteAndOverviewToolbox_Development and then copies to SiteAndOverviewToolbox
    # All stored in Desktop - csostad work backup because github seemed to erase the Testing folder

# Dec 07, 2023
    # Added the zoom to feature extent function. Now you can pass in the map name, map frame, layer name, zoom factor, and layout name and it will zoom to the feature extent
    # Todays working file was SiteAndOverviewToolbox_Development stored in csostad_csostad_work_backup_TESTING

# December 12th
    # Added the functionality to handle amendments. Now, if an application has an amendment attached to it, you can run the amendment tool and it will create a new folder with the
    # amendment file number and then create basic shapefile, centroid with lat/long and splitlines with length calculated and then add these to the site map, overview map, and inset maps
    # These would then be symbolized in blue hatched for "area to be removed" or red hatched for "area to be added"
    # Currently the amendment folder is created in the main Lands folder, need to have it create the amendment in the file_num folder

    # Fixed issue where imagery credits and imagery NA were not being turned on for site imagery. The issue arose because of multiple layouts rather than just one
    # as was the case with the initial code from Lands_4 toolbox

# December 13th
    # Added export overview functionality
    # Added the optional parameter for client name which will be used on the Overview map.
    # Added function to automatically open the folder where the pdfs are stored after the export is complete

# December 14th 
    # Split the old _DEV file into DEV_DEF which indicates I removed the temp_layer selection and went with a Definition query instead
    # Add the SORT function in case the def query returns multiple records. This will sort the records in ascending order and then the first record will be the one in the application stage
    # Fixed Overview map and overview Inet map failing to zoom properly.

# December 27th
    # If the Client name is left empty, it will use the existing client name. 
    # Started to get a file_num not defined error. Tried the back up scripts and it still had an issue. In order to fully roll back, would have use a file from before today
    # Or before adding the client name optional leave blank snippet.
    # Reverted to the backup version of the Arc Pro Folder and the Pre- don’t change client name snippet and it seems to be working normal again. Put the possibly 
    # broken version in the csostad_work_backup folder in case we need to go back to it.labeled it broken.


# January 11th, 2024
    # Updated comments
    # Changed splitline rounding to 0 decimal places

# January 15th/16th
    # Had github issues. Made quarentine folder and copied the toolbox to it as a backup.

# January 16th, 2024
    # Fixed the amendment tool so that it is no longer using temporary layers and is now using definition queries instead.
    #!!Still needs to pull the file number from the layout and use that to find the lands folder and create the amendment folder inside that folder
    # Worked on the connection properties bug. Still not fixed. But took the DEV version and moved it to production (minus the connection properties fix)
    # Change the "Broken Connections" version to Dev

# January 18th, 2024
    # Fixed the connection properties bug. lands_file_path was still being used instead of crown_file_folder. Changed it to crown_file_folder and it works now.
    # os.path.join(lands_file_path, file_num)} was changed to os.path.join(lands_file_path, file_num, crown_file_folder)}
    # Amendment tool now places files in the request file folder
    # Change amendment search to def query
    # remove file name concatenation and updated with os.path.join
    
# January 19th, 2024
    # FIXED Multiple amendments need to be uniquely named (centroids, splitlines) or they will overwrite during a single day. Ie, you need to remove files 1234 and 4321 from file 9876, you run the amendment tool twice,
    # it would overwrite the second time. 

# January 22nd, 2024
    # Added the auto inset functionality based on the size of the area. If area is greater than 1 hectare, it will not create an inset map. If it is less than 1 hectare, it will create an inset map
    # Furthermore, the zoom factor on smaller parcels was not suitable, now if the the areas is less than 1 hectare it will zoom out 160%(?)
    # If the area is greater than 100 hectares, the overview map will have the inset turned off


    
    
############################################################################################################################################################################
#
# Known Issues
#
############################################################################################################################################################################







############################################################################################################################################################################
#
# Define Global Variables to be used in the various versions of the Export Layout Tools
#
############################################################################################################################################################################

# Set up variables for the various layers to be turned on and off in the site map
layers_for_site = ["Base Map Auto Scale (1:7,500,000-1:20,000)", "0-7.5K SITE MAP", "7.5-15K SITE MAP", "15-35K SITE MAP", "35-75K SITE MAP", "75-400K SITE MAP", "centroid", "baseDataText", "ALL Freshwater Atlas Labels GROUP"]
layers_for_imagery = ["Latest BC RGB Spot"]

# Set up variables for the various text elements to be turned on and off in the imagery site map
elements_for_inset = ["Inset Map Words", "Inset Map Dynamic Scale"]

# Get the current date to be used in naming of the pdf 
current_date = datetime.datetime.now()
formatted_date = current_date.strftime("%b%Y")

# Assign the layout names to variables
site_layout = "FCBC_Site_FileNo_REF_mmmyyyy_85X14"
overview_layout = "FCBC_Overview_FileNo_REF_mmmyyyy_17X11_inset"

#assign variables to identify aprx project, and map 
aprx = arcpy.mp.ArcGISProject("CURRENT")
site_map = aprx.listMaps('Layers')[0] 
inset_map = aprx.listMaps('SiteInsetMap')[0]
overview_map = aprx.listMaps('OverviewLayers')[0]
overview_inset_map = aprx.listMaps('OverviewInsetMap')[0]
lands_file_path = '\\\\spatialfiles.bcgov\\work\\lwbc\\nsr\\Workarea\\fcbc_prg\\FCBC\\Lands_Files\\'
site_map_frame = 'Layers Map Frame'
site_map_inset_frame = 'Inset Data Frame Map Frame' # Could assign these strings to objects right here  
overview_map_frame = 'Overview Layers Map Frame'
overview_map_inset_frame = 'Inset Data Frame Map Frame'

# Assign Application to Layer for the Export tool. This will be removed later if integrated with the Site and Overview tool
layer = site_map.listLayers("Application")[0] #This could cause issues in other scripts if Application is not the key layer
crown_tenures_layer = site_map.listLayers("WHSE_TANTALIS.TA_CROWN_TENURES_SVW")[0]

############################################################################################################################################################################
#
# Set up global functions to be used in the various versions of the Export Layout Tools
#
############################################################################################################################################################################



# Global exportToPdf function
def exportToPdf(layout, workSpace_path, pdf_file_name):
    out_pdf = f"{workSpace_path}\\{pdf_file_name}"
    layout.exportToPDF(
        out_pdf=out_pdf,
        resolution=300,  # DPI
        image_quality="BETTER",
        jpeg_compression_quality=80  # Quality (0 to 100)
    )

# This function will turn on the site map layers, turn on the inset map and turn off the imagery layers
# Currently not used
def turn_on_site_map_with_inset():
    global site_layout
    # Step 1 - LAYERS - Turn off the imagery layers and turn ON the site map layers
    for lyr in site_map.listLayers():
        if lyr.name in layers_for_imagery:
            lyr.visible = False
        elif lyr.name in layers_for_site:
            lyr.visible = True

    # Step 2 - TEXT ELEMENTS - Turn off the imagery related elements (Imagery Credits and ImageryNA)
    for lyt in aprx.listLayouts(site_layout):
        for elm in lyt.listElements("TEXT_ELEMENT"):
            if elm.name == "Imagery Credits":
                
                elm.visible = False
            
            if elm.name == "ImageryNA":
                elm.text = "Imagery: NA"
                
        if elm.name in elements_for_inset:
            elm.visible = True # Turn the inset map elements back on

    # Step 3 - MAPFRAME ELEMENTS - Turn on the inset map and extent indicator            
    for elm in lyt.listElements("MAPFRAME_ELEMENT"):
            if elm.name == "Inset Data Frame Map Frame":
                elm.visible = True # Turn the inset map back on
            elif elm.name == "ExtentIndicator":
                elm.visible = True # Turn the extent indicator back on
 
 
#This function will turn on the inset map and associated supporting elements
def turn_on_inset(layout_to_turn_on):
   
    for lyt in aprx.listLayouts(layout_to_turn_on):         
        for elm in lyt.listElements("MAPFRAME_ELEMENT"):
                if elm.name == "Inset Data Frame Map Frame":
                    elm.visible = True # Turn the inset map back on
                elif elm.name == "ExtentIndicator":
                    elm.visible = True # Turn the extent indicator on 
        for elm in lyt.listElements("TEXT_ELEMENT"):
                if elm.name in elements_for_inset:
                    elm.visible = True # Turn the inset map elements on
        for elm in lyt.listElements("GRAPHIC_ELEMENT"):
                if elm.name == "Inset Map Scale Rectangle":
                    elm.visible = True # Turn the scale rectangle on

#This function will turn off the inset map and associated supporting elements
def turn_off_inset(layout_to_turn_off):
   
    for lyt in aprx.listLayouts(layout_to_turn_off):         
        for elm in lyt.listElements("MAPFRAME_ELEMENT"):
                if elm.name == "Inset Data Frame Map Frame":
                    elm.visible = False # Turn the inset map back off
                elif elm.name == "ExtentIndicator":
                    elm.visible = False # Turn the extent indicator off 
                elif elm.name == "Inset Map Scale Rectangle":
                    elm.visible = False # Turn the scale rectangle off
        for elm in lyt.listElements("TEXT_ELEMENT"):
                if elm.name in elements_for_inset:
                    elm.visible = False # Turn the inset map elements off
        for elm in lyt.listElements("GRAPHIC_ELEMENT"):
                if elm.name == "Inset Map Scale Rectangle":
                    elm.visible = False # Turn the scale rectangle on
                
# This function will turn on the site map layers and turn off the imagery layers
def turn_on_site_map():
    
    global site_layout, aprx
    # Step 1 - LAYERS - Turn off the imagery layers and turn ON the site map layers
    for lyr in site_map.listLayers():
        if lyr.name in layers_for_imagery:
            lyr.visible = False
        elif lyr.name in layers_for_site:
            lyr.visible = True

    # Step 2 - TEXT ELEMENTS - Turn off the imagery related elements (Imagery Credits and ImageryNA)
    for lyt in aprx.listLayouts(site_layout):
        
        
        for elm in lyt.listElements("TEXT_ELEMENT"):
            
            if elm.name == "Imagery Credits":
                
                elm.visible = False
            
            if elm.name == "ImageryNA":
                elm.text = "Imagery: NA"
                

# This function will turn on the imagery layers and turn off the site map layers and update the necessary text elements
def turn_on_imagery():
    global site_layout # IS IT BETTER TO DO IT THIS WAY OR TO PASS IN THE LAYOUT AS A PARAMETER?
    for lyr in site_map.listLayers():
        if lyr.name in layers_for_imagery: # If the any of the layers are in the list of layers defined in the layer for imagery variable, turn it on
            lyr.visible = True #Turn on the imagery layers
        
        elif lyr.name in layers_for_site: # If the layer name is in the list of layers defined in the layer for site variable, turn it off
            lyr.visible = False # Turn off the site map layers
        
    #Iterate through the list elements in the layout and find the text element named "Imagery Credits" and turn it on, find ImageryNA and replace NA with Spot 2021
    for lyt in aprx.listLayouts(site_layout):
        arcpy.AddMessage(f"Layout name is: {lyt.name}")
        for elm in lyt.listElements("TEXT_ELEMENT"):
            #arcpy.AddMessage(f"Text element name is: {elm.name}")
            if elm.name == "Imagery Credits":
                arcpy.AddMessage("Imagery Credits found")
                elm.visible = True
            elif elm.name == "ImageryNA":
                arcpy.AddMessage("IMAGERY: N/A found")
                elm.text = "Imagery: Spot 2021"
            

# Function to zoom to feature extent
def zoom_to_feature_extent(map_name, map_frame, layer_name, zoom_factor, layout_name):


    ''' This function will focus the layout on the selected feature and then pan out x% (depending on zoom factor) 
    to show the surrounding area. If you use an 0.8 (80%) Zoom Factor, to calculate the zoom percentage: Original zoom 
    is 100% (the initial extent of the splitline layer).The new extent is 160% larger than the original. This is because 
    you are adding 80% of the width to both sides (left and right) and 80% of the height to both top and bottom.'''\
        
    arcpy.AddMessage("         Running zoom_to_feature_extent function")
    # Verify layout existence
    layouts = aprx.listLayouts(layout_name)
    if not layouts:
        arcpy.AddError(f"No layout found with the name: {layout_name}")
        return
    lyt_name = layouts[0]

    # Verify map frame existence
    map_frames = lyt_name.listElements("MAPFRAME_ELEMENT", map_frame)
    if not map_frames:
        arcpy.AddError(f"No map frame found with the name: {map_frame} in layout: {layout_name}")
        return
    mf = map_frames[0]

    # Verify layer existence
    maps = aprx.listMaps(map_name)
    if not maps:
        arcpy.AddError(f"No map found with the name: {map_name}")
        return
    map_obj = maps[0]
    layers = map_obj.listLayers(layer_name)
    if not layers:
        arcpy.AddError(f"No layer found with the name: {layer_name} in map: {map_name}")
        return
    lyr_name = layers[0]

    # Get and adjust the extent
    try:
        current_extent = mf.getLayerExtent(lyr_name, False, True)
        x_min = current_extent.XMin - (current_extent.width * zoom_factor)
        y_min = current_extent.YMin - (current_extent.height * zoom_factor)
        x_max = current_extent.XMax + (current_extent.width * zoom_factor)
        y_max = current_extent.YMax + (current_extent.height * zoom_factor)
        new_extent = arcpy.Extent(x_min, y_min, x_max, y_max)
        mf.camera.setExtent(new_extent)
        arcpy.AddMessage(f"         Zoomed to {layer_name} with zoom factor {zoom_factor}")
        # # Get the current scale of the map frame
        # current_scale = mf.camera.getScale()
        # arcpy.AddMessage(f"Current scale is: {current_scale}")
        # # Round the scale to the nearest 10,000
        # rounded_scale = round(current_scale / 10000) * 10000
        # arcpy.AddMessage(f"Rounded scale is: {rounded_scale}")
        # # Set the map frame to the new rounded scale
        # mf.camera.setScale(rounded_scale)
        # arcpy.AddMessage(f'Zooming the map to the new scale of {rounded_scale}')
        
    except Exception as e:
        arcpy.AddError(f"Error in zooming to extent: {str(e)}")

def format_dms(dms_str):
    '''
    The purpose of the format_dms function is to format coordinates (DMS) 
    format and round the seconds to two decimals.
    '''
    parts = dms_str.split()
    
    # Check if the dms_str has the expected format
    if len(parts) != 3:
        raise ValueError("Invalid DMS format in the centroid layer. It should have three parts separated by whitespace.")
    
    # The split() method splits the string at whitespace characters. The result of the split() method is assigned to the parts variable, 
    # which becomes a list containing the individual parts of the dms_str string.The code then extracts specific parts from the parts list using indexing.
    parts = dms_str.split()
    degrees, minutes, seconds_with_dir = parts[0], parts[1], parts[2]
    arcpy.AddMessage("         Function to truncate seconds to two decimal places is running.....")


    
    # Split seconds and direction into two parts. The seconds and the direction. Seconds is obtained by slicing the seconds with dir string 
    # from the beginning to the second last character. The direction is obtained by slicing the seconds with dir string from the last character.
    seconds, direction = seconds_with_dir[:-1], seconds_with_dir[-1]

    # The code then truncates the seconds to two decimal places using the float function and string formatting.
    truncated_seconds = f"{float(seconds):.2f}"

    # The code then constructs the formatted DMS by combining the degrees, minutes, truncated seconds, and direction.
    formatted_dms = f"{degrees} {minutes} {truncated_seconds}\"{direction}"
    
    # The code then returns the formatted DMS.
    return formatted_dms



class Toolbox(object):
    def __init__(self):
        """Define the toolbox (name of toolbox is name of the file)"""
        self.label = "Toolbox"
        self.alias = ""

        #List of tool classes associated with this toolbox
        self.tools = [FullSiteOverviewMaps, ExportSiteAndImageryLayout, Amendment]


class FullSiteOverviewMaps(object):
    def __init__(self):
        """This tool will prep all required data - it will create the application polygon, centroid, and splitline, 
        and calculate geometries for centroid and splitline, then update the datasources. It will then create the overview map and export all layouts"""
        self.label = "Create Full Site Map and Overiew Map"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 
        #This parameter is the file number of the application
        file_num = arcpy.Parameter(
            displayName = "Lands File Number",
            name="file_num",
            datatype="String",
            parameterType="Required",
            direction="Input")

        parameters = [file_num]

        return parameters

    
    def execute(self,parameters,messages):
        try:
               
            # Bring in parameters to the function to be used as variables 
            file_num = parameters[0].valueAsText
            
            # Create a valid output folder path
            crown_file_folder = os.path.join(lands_file_path, file_num)                    
            
            # Use if-then logic to verify that the intended directory has not been created already
            # If the directory exists create a new folder inside the existing folder and append the month, date and the year 
            # to the folder name.
            
            #####################################################################
            #
            # Step 1 - Folder Creation
            #
            #####################################################################
            
            arcpy.AddMessage(f"Step 1 - Creating folder for {file_num}")
            
            # Check the Lands Directory to see if a folder for this file number already exists, if the folder already exists,
            # create a new folder inside the existing folder and append the month, date and the year to the folder name. Then
            # set the crown_file_folder variable to point to the new folder. If the folder doesn't exist, create a new one
            # with the file number as the name of the folder.
            
            if os.path.isdir(crown_file_folder):
                arcpy.AddWarning("A directory for this lands file has already been created. Creating a new subfolder inside with date appended.")
                
                # Open the folder as an indication that the folder has already been created
                try:
                    os.startfile(crown_file_folder)
                except Exception as e:
                    arcpy.AddError(f"Could not open folder {crown_file_folder}: {e}")
                
                # Create a subfolder with the current month and year appended to its name
                
                # Create the file name for the new subfolder
                new_folder_name = file_num + "_" + datetime.datetime.now().strftime("%m_%d_%Y")
                arcpy.AddMessage(f"         New folder name is: {new_folder_name}")
                
                # If dir exists, create a new folder path name with the crown_file_folder and the new_folder_name
                # It should look like this: 1234567/1234567_Jan012021
                new_folder_path = os.path.join(crown_file_folder, new_folder_name)
                arcpy.AddMessage(f"         New folder path is: {new_folder_path}")
                
                # Check to see if it already exits, if it doesn't, create it
                if not os.path.exists(new_folder_path):
                    os.makedirs(new_folder_path)
                    arcpy.AddMessage(f"        The date appended folder does not exist. Creating: {new_folder_path}")
                # Update the crown_file_folder variable to point to the new subfolder
                crown_file_folder = new_folder_path
                arcpy.AddMessage(f"         Crown_File_Folder variable updated to: {crown_file_folder}")
            else:
                # Create the directory for the request number
                os.makedirs(crown_file_folder)
                arcpy.AddMessage("         There is no existing Directory. New Directory created for " + str(file_num) in {crown_file_folder})
            

            # Create a path name for the output of the shapefile by joining the crown_file_folder and the file_num
            output_shp_path = os.path.join(crown_file_folder, f"{file_num}.shp")
            arcpy.AddMessage(f'         Output path for shapefiles is: {output_shp_path}')

            
            ################################################################################################################################
            #
            # Step 2 - Create the Application Polygon
            #
            #############################################################################################################################
            
            arcpy.AddMessage("Step 2 - Creating Application Polygon")
            
            # Set the definition query and assign it to a variable
            expression = arcpy.AddFieldDelimiters(arcpy.env.workspace, 'CROWN_LANDS_FILE') + f" = '{file_num}'"
            
            # Apply the expression to the layer "Application" in the Site Map (Defined as a global variable at the top of the script)
            crown_tenures_layer.definitionQuery= expression
            arcpy.AddMessage(f"         Definition query {expression} set for layer: {crown_tenures_layer}")

        
            # Error Handling -  Use a SearchCursor to count the number of records returned by the definition query
            count = 0
            with arcpy.da.SearchCursor(crown_tenures_layer, "*") as cursor:
                for row in cursor:
                    count += 1

            # Check if the count is 0 and display a message if the definition query was successful or not.
            if count == 0:
                arcpy.AddMessage(f"         No records returned by definition query: {expression}, please check the file number and try again")
            else:
                arcpy.AddMessage(f"         Definition query {expression} set for layer: {crown_tenures_layer}. Records returned: {count}")

        
                                        
            # Allow overwrite of existing application file - used for testing during the development of this code
            # Comment out if not needed
            arcpy.env.overwriteOutput = True
            
            # Pass the Crown Tenures Layer that has the definition query on it into this sort function to sort the results in ascending order. 
            # This should ensure that if you have multiple records returned by the definition query, the first record will be the one that is 
            # in the application stage. This is the one that will be used to create the rest of the map. 
            
            arcpy.management.Sort(crown_tenures_layer, output_shp_path, [["TENURE_STAGE", "ASCENDING"]])
            arcpy.AddMessage("         Sorting the results of the definition query by TENURE_STAGE in ascending order")
            arcpy.AddMessage("         Shapefile created")
            
            ###################################################################################################################################
            #
            # Step 3 - Update the data sources of the application layers in the Site Map, Overview Map, and Overview Inset Map
            #
            ######################################################################################################################################
            
            def update_layer_connection(layer_name, map_name):
                
                """
                Updates the data source connection properties for a specified layer in a the map.

                This function is designed to update the connection properties of a layer identified by 'layer_name' in the map 
                specified by 'map_name'. The function first retrieves the target layer from the map. It then logs the original 
                connection properties of this layer for reference. The new connection properties are set up to point to the new shapefile 
                found in the file folder.
    

                Parameters:
                layer_name (str): The name of the layer whose connection properties are to be updated.
                map_name (arcpy._mp.Map): The ArcGIS map object containing the target layer.

                The function assumes that 'lands_file_path' and 'file_num' variables are available in the scope where this function 
                is called, and uses these to construct the path to the new shapefile. The connection properties of the target layer 
                are then updated to this new path, effectively redirecting the layer to a new data source.
                """
                # List the layers and assign the layer that was passed to the function as an argument to a variable
                target_lyr = map_name.listLayers(layer_name)[0]
                arcpy.AddMessage("         Updating connection properties for " + str(target_lyr) + "")
            

                # Set a variable that represents layer original connection properties
                origConnPropDict = target_lyr.connectionProperties
                
                
                # Set new connection properties based on the layer shapefile exported in earlier step 
                newConnPropDict = {'connection_info': {'database': os.path.join(lands_file_path, file_num, crown_file_folder)},
                                'dataset': f"{layer_name}.shp",
                                'workspace_factory': 'Shape File'}
                
                
                # Update connection properties 
                target_lyr.updateConnectionProperties(origConnPropDict, newConnPropDict)
                arcpy.AddMessage(f"         {target_lyr} Connection properties successfully updated") 
            
            # Call the update_layer_connection function to update the connection properties of the application layer in the site map

            # Had to create a separate function for the Application layer because the shapefile stored in the folder
            # is named the file number, not "Application", whereas the other layers are named the same as the layer name
            def update_application_layer_connection(map_name, file_num, lands_file_path):
                
                """
                Updates the data source connection properties of the 'Application' layer in a specified map.


                Parameters:
                map_name: The ArcGIS map object in which the 'Application' layer exists.
                file_num (str): The file number used to name the shapefile.
                lands_file_path (str): The file path where the shapefiles are stored.

                The function identifies the 'Application' layer, constructs the new connection properties with the appropriate 
                shapefile path and name, and updates the layer's connection properties accordingly.
                """
                
                # The layer name in the contents pane is "Application"
                layer_name = "Application"
                
                # Find the "Application" layer in the specified map
                layers = map_name.listLayers(layer_name)
                if not layers:
                    arcpy.AddWarning(f"No layer named '{layer_name}' found in {map_name.name}.")
                    return
                target_lyr = layers[0]
               

                # Prepare the new connection properties
                # The shapefile's name is based on file_num
                newConnPropDict = {
                    'connection_info': {'database': os.path.join(lands_file_path, file_num, crown_file_folder)},
                    'dataset': f"{file_num}.shp",
                    'workspace_factory': 'Shape File'
                }
                # arcpy.AddMessage(f"        New connection properties for {target_lyr} are: {newConnPropDict}")
                
                # Update the connection properties
                origConnPropDict = target_lyr.connectionProperties
                target_lyr.updateConnectionProperties(origConnPropDict, newConnPropDict)
                arcpy.AddMessage(f"         Data source updated for {map_name}")

            update_application_layer_connection(site_map, file_num, lands_file_path)
            update_application_layer_connection(overview_map, file_num, lands_file_path)
            update_application_layer_connection(overview_inset_map, file_num, lands_file_path)

                        
            #########################################################################################################################################
            # Step 4 - Create the splitline layer to be used in the inset map. 
            # Not needed if Tenure_1_A is greater than 1 Hectare. But at this point, 
            # it is just creating it for all applications
            #
            ##########################################################################################################################################

            # Set a variable for the Splitlines feature class filename
            splitline_fc = os.path.join(crown_file_folder, "splitline.shp")
            
            # Create splitline from the previosuly created shapefile in the crown_file_folder 
            arcpy.management.SplitLine(output_shp_path, splitline_fc)
            arcpy.AddMessage(f"Step 4 -  Splitline created for {file_num}")
            
            # Add field named "length" to the splitline for labelling purposes
            arcpy.AddField_management(splitline_fc, "Length", "DOUBLE")
            arcpy.AddMessage("         Length field added to splitline.shp")

                                
            # Use an update cursor to iterate through each feature in the splitline feature class
            # and calculate the length of the feature geometry in meters then round the length value 
            # to 1 decimal place before storing it in the 'Length' field
            with arcpy.da.UpdateCursor(splitline_fc, ["SHAPE@", "Length"]) as cursor:
                for row in cursor:
                    length_m = round(row[0].length, 0)
                    
                    # Set the calculated length to the 'Length' field of the current row
                    row[1] = length_m 
                    
                    # Update the row with the new length value
                    cursor.updateRow(row)  
            
            arcpy.AddMessage("         Length field calculated for splitline.shp")


            #Call the update_layer_connection function to update the connection properties of the splitline layer in the Inset map
            update_layer_connection("splitline", inset_map) 
            arcpy.AddMessage("         Splitline layer data source updated for inset map")
        
            # Call the function to zoom to the feature extent of the splitline layer and then pan out 160% to show the surrounding area
            arcpy.AddMessage("         Zooming to the feature extents for the Application/Splitline layer for Overview Map, SiteMap and Inset Map and then panning out 160% to show the surrounding area")
            arcpy.AddMessage("-------------------------------------------------------------------------------------------- ")
            arcpy.AddMessage("-----------------THE SCALE WILL NOT BE A ROUND NUMBER. BE SURE TO ADJUST--------------------- ")
            arcpy.AddMessage("--------------------------------------------------------------------------------------------- ")
            
            
            zoom_to_feature_extent(site_map.name, site_map_frame, "Application", 0.8, site_layout)
            
            zoom_to_feature_extent(inset_map.name, site_map_inset_frame, "splitline", 0.8, site_layout)
            
            zoom_to_feature_extent(overview_inset_map.name, overview_map_inset_frame, "Application", 0.1, overview_layout)
                
    
            
            
            ############################################################################################################################################################################
            #
            # Step 5 - Create the Centroid Layer to power the datadriven pages in the map series
            #
            ############################################################################################################################################################################
            
            arcpy.AddMessage("Step 5 - Creating Centroid Layer")
            
            # Set a variable for the Centroid feature class
            centroid_fc = os.path.join(crown_file_folder, "centroid.shp")
            
            # Create points from the application layer
            arcpy.management.FeatureToPoint(output_shp_path, centroid_fc, "CENTROID")
            arcpy.AddMessage("         Centroid successfully created in " + centroid_fc)        
            
            # Add Lat and Long fields to the Centroid  
            
            arcpy.management.AddField(centroid_fc, "Lat", "TEXT")

            arcpy.management.AddField(centroid_fc,"Long", "TEXT")
            arcpy.AddMessage(f"         Lat and Long fields added to {centroid_fc}")
            
            # Calculate the Lat and Long fields
            arcpy.management.CalculateGeometryAttributes(in_features = os.path.join(crown_file_folder,"centroid.shp"), geometry_property=[["Lat", "POINT_Y"], ["Long", "POINT_X"]], coordinate_format="DMS_DIR_FIRST")[0]
            arcpy.AddMessage("         Lat and Long fields calculated for centroid.shp")
            
            
            # Use an update cursor to iterate through each feature in the centroid feature class
            # Update the Lat and Long fields with formatted values returned from the format dms function
            with arcpy.da.UpdateCursor(centroid_fc, ["Lat", "Long"]) as cursor:
                for row in cursor:
                    try:
                        # Format each field using the custom function
                        row[0] = format_dms(row[0])
                        row[1] = format_dms(row[1])
                        cursor.updateRow(row)
                    except ValueError as e:
                        arcpy.AddWarning(str(e))
            arcpy.AddMessage("         Lat and Long fields formatted to 2 decimals for centroid.shp")
            
            update_layer_connection("centroid", site_map)
            arcpy.AddMessage("         Centroid layer data source updated for site map")

            arcpy.AddMessage("Step 6 - Checking the Area field of the Application Layer to determine if the inset map should be turned on or off")
            # turn_off_inset(site_layout)
            # turn_off_inset(overview_layout)
            
            # Check the Area (Tenure_A_1) field of the Application Layer, if the Area is < 1 hectare, turn on the site_map_inset_frame
            # If Area is > 100 hectares, turn off 
            with arcpy.da.UpdateCursor(layer, ["TENURE_A_1"]) as cursor:
                for row in cursor:
                    if row[0] < 1:
                        arcpy.AddMessage(f"         Area is less than 1 hectare ({row[0]} Ha). Turning on the Site Layout and Overview Layout inset maps")
                        turn_on_inset(site_layout)
                        turn_on_inset(overview_layout)
                        
                        # If the area is less than 1 hectare, zoom to the feature extent of the application layer and then pan out 160% to show the surrounding area
                        # as the preview amount of zoom was not suitable for such a small area
                        arcpy.AddMessage("         Readjusting zoom to compensate for extremely small area. Zoom factor changed to 2.0 and 12.0")
                        zoom_to_feature_extent(overview_inset_map.name, overview_map_inset_frame, "Application", 2.0, overview_layout)
                        zoom_to_feature_extent(site_map.name, site_map_frame, "Application", 12.0, site_layout)
                    elif 1 < row[0] <= 100:
                        arcpy.AddMessage(f"         Area is greater than 1 hectare and less than or equal to 100 hectares ({row[0]} Ha). Turning off the Site Layout inset map and keeping the Overview Layout inset map on")
                        turn_off_inset(site_layout)
                        # Assuming the overview_layout is already on, if not use turn_on_inset(overview_layout)
                    elif row[0] > 100:
                        arcpy.AddMessage(f"         Area is greater than 100 hectares ({row[0]} Ha). Turning off the Overview Layout inset map")
                        turn_off_inset(overview_layout)
                        turn_off_inset(site_layout)


            arcpy.AddMessage("SCRIPT COMPLETED SUCCESSFULLY")
            
        except arcpy.ExecuteError:
            msgs = arcpy.GetMessages(2)
            arcpy.AddError(msgs)

        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            #return python error messages for use in script tool
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)



class Amendment(object):
    def __init__(self):
        
        """This tool will prep all required data for an individual crown tenure - to be used to add/subtract amendment - it will create the 
        amendment polygon, centroid, and splitline, and calculate geometries for centroid and splitline"""
        
        self.label = "Create Amendment Polygon, Centroid, Splitlines"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 
        #This parameter is the file number of the application
        amend_file_num = arcpy.Parameter(
            displayName = "Lands Amendment File Number",
            name="file_num",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        parameters = [amend_file_num]

        return parameters

    
    def execute(self,parameters,messages):
        try:
            # Bring in parameters to the function to be used as variables 
            amend_file_num = parameters[0].valueAsText
            
            ############################################################################################################################################################################
            #
            # Create the shapefile polygon layer to be used for the Amendment.
            #
            ############################################################################################################################################################################

            
            # Find the application layer in the site map and then use a .da search cursor to iterate through the first feature in the layer
            # and find the field called CROWN_LANDS. This will be used to create the folder for the amendment
            
            with arcpy.da.SearchCursor(layer, "CROWN_LAND") as cursor:
                for row in cursor:
                    request_file_num = row[0]
                    arcpy.AddMessage(f"Current Project File Number is: {request_file_num}") 
                    break
                                       
            # Create a valid output folder path using the file number pulled from the layout
            amend_file_folder = os.path.join(lands_file_path, request_file_num, request_file_num + str('_Amendment')) 

                                   
            # Use if-then logic to verify that the intended directory has not been created already
            # If the directory exists create a new folder inside the existing folder and append the month and the year 
            # to the folder name.
                     
            if os.path.isdir(amend_file_folder):
                arcpy.AddWarning("A directory for this lands file has already been created. Creating a new subfolder inside with date appended.")
                
                # Create a subfolder with the current month and year appended to its name
                new_folder_name = amend_file_folder + "_" + datetime.datetime.now().strftime("%m_%d_%Y")
                new_folder_path = os.path.join(amend_file_folder, new_folder_name)
                if not os.path.exists(new_folder_path):
                    os.makedirs(new_folder_path)

                # Update the amend_file_folder variable to point to the new subfolder
                amend_file_folder = new_folder_path
            else:
                # Create the directory for the request number
                os.makedirs(amend_file_folder)
                arcpy.AddMessage(f"Directory:{amend_file_folder} created for amendment {amend_file_num}")
            
            # Use a definition query to select the amendment file number from the crown tenures layer and then create a shapefile, centroid, and splitline
            
            output_shp_path = os.path.join(amend_file_folder, f"Amend_{amend_file_num}.shp")
            arcpy.AddMessage(f'         Amendment Output shapefile path is: {output_shp_path}')

            expression = arcpy.AddFieldDelimiters(arcpy.env.workspace, 'CROWN_LANDS_FILE') + f" = '{amend_file_num}'"
            arcpy.AddMessage(f'         Amendment Expression is: {expression}')

            # Apply the expression to the layer "Application" in the Site Map (Defined as a global variable at the top of the script)
            crown_tenures_layer.definitionQuery= expression
            arcpy.AddMessage(f"        Definition query {expression} set for layer: {crown_tenures_layer}")
            
            # Error Handling -  Use a SearchCursor to count the number of records returned by the definition query
            count = 0
            with arcpy.da.SearchCursor(crown_tenures_layer, "*") as cursor:
                for row in cursor:
                    count += 1

            # Check if the count is 0 and display a message if the definition query was successful or not.
            if count == 0:
                arcpy.AddMessage(f"         No records returned by definition query: {expression}, please check the file number and try again")
            else:
                arcpy.AddMessage(f"         Definition query {expression} set for layer: {crown_tenures_layer}. Records returned: {count}")


            # Set overwrite to true
            arcpy.env.overwriteOutput = True
                     
            # Sort the result of the definition query by TENURE_STAGE in ascending order so that you will end up with "A" at the top of the list thus choosing the "Application" stage
            arcpy.management.Sort(crown_tenures_layer, output_shp_path, [["TENURE_STAGE", "ASCENDING"]])
            arcpy.AddMessage("         Sorting the results of the definition query by TENURE_STAGE in ascending order")
            arcpy.AddMessage("         Amendment Shapefile created")
            
          
            # Add the Amendment feature to the site map, overview map, site inset map, and overview inset map
            site_map.addDataFromPath(output_shp_path)
            overview_map.addDataFromPath(output_shp_path)
            inset_map.addDataFromPath(output_shp_path)
            overview_inset_map.addDataFromPath(output_shp_path)
            
                                     
            ############################################################################################################################################################################
            #
            # Create the splitline layer to be used in the inset map. Not needed if Tenure_1_A is greater than 1 Hectare. But at this point, it is just creating it for all applications
            #
            ############################################################################################################################################################################

            # Set a variable for the Splitlines feature class
            amend_splitline_fc = os.path.join(amend_file_folder, f"Amend_{amend_file_num}_splitlines.shp")
           
            #create splitline from selection 
            arcpy.management.SplitLine(output_shp_path, amend_splitline_fc)

            
            # Add field length to the Split Line
            arcpy.AddField_management(amend_splitline_fc, "Length", "DOUBLE")
            arcpy.AddMessage("Length field added to amend_splitline.shp")

                                
            # Use an update cursor to iterate through each feature in the splitline feature class
            with arcpy.da.UpdateCursor(amend_splitline_fc, ["SHAPE@", "Length"]) as cursor:
                for row in cursor:
                    # Calculate the length of the feature geometry in meters
                    # Round the length value to 1 decimal place before storing it in the 'Length' field
                    length_m = round(row[0].length, 0)
                    row[1] = length_m  # Set the calculated length to the 'Length' field of the current row
                    cursor.updateRow(row)  # Update the row with the new length value
            arcpy.AddMessage("Length field calculated for amend_splitline.shp")

            # Add the splitline feature to the site inset map
            inset_map.addDataFromPath(amend_splitline_fc)
            
        
            ############################################################################################################################################################################
            #
            # Create the Centroid Layer to power the datadriven pages in the map series
            #
            ############################################################################################################################################################################
            
            
            # Set a variable for the Centroid feature class file name
            amend_centroid_fc = os.path.join(amend_file_folder, f"Amend_{amend_file_num}_centroid.shp")
            
            # Use the feature to point tool on the shapefile created in the previous step to create a centroid
            arcpy.management.FeatureToPoint(output_shp_path, amend_centroid_fc, "CENTROID")
            arcpy.AddMessage("Amendment Centroid created in " + amend_centroid_fc)        
            
            # Add Lat and Long fields to the Centroid  
            arcpy.management.AddField(amend_centroid_fc, "Lat", "TEXT")

            arcpy.management.AddField(amend_centroid_fc,"Long", "TEXT")
            
            
            # Calculate the Lat and Long fields
            arcpy.management.CalculateGeometryAttributes(in_features = amend_centroid_fc, geometry_property=[["Lat", "POINT_Y"], ["Long", "POINT_X"]], coordinate_format="DMS_DIR_FIRST")[0]
                                                         
            arcpy.AddMessage("Lat and Long fields calculated for Amendment centroid")
            
            # Use an update cursor to iterate through each feature in the centroid feature class
            # Update the Lat and Long fields with formatted values
            
            with arcpy.da.UpdateCursor(amend_centroid_fc, ["Lat", "Long"]) as cursor:
                for row in cursor:
                    try:
                        # Format each field using the custom function
                        row[0] = format_dms(row[0])
                        row[1] = format_dms(row[1])
                        cursor.updateRow(row)
                    except ValueError as e:
                        arcpy.AddWarning(str(e))
            
            arcpy.AddMessage("Lat and Long fields formatted to 2 decimals for amend_centroid.shp")
            
            # Add the centroid feature to the site map
            site_map.addDataFromPath(amend_centroid_fc)
            
            

        except arcpy.ExecuteError:
            msgs = arcpy.GetMessages(2)
            arcpy.AddError(msgs)

        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            #return python error messages for use in script tool
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)


class ExportSiteAndImageryLayout(object):
    def __init__(self):
        self.label = "Export Site, Imagery, Overview Layout With Manual Inset"
        self.description = ""
        self.canRunInBackground = False



    def getParameterInfo(self):
        
        '''
        This parameter prompts the user to enter the folder where they want the maps exported to. 
        If it is left blank, it will default to the file folder in the Lands Folder
        
        '''
        workspace = arcpy.Parameter(
            displayName = "OPTIONAL - Select the output folder for the Site and Imagery Layouts (Will Overwrite) - Default is Lands - File Number folder",
            name="workspace",
            datatype="DEWorkspace",
            parameterType="Optional",
            direction="Input"
        )
        
        '''
        This parameter prompts the user to enter the client name to be used on the overview map. If it is left blank the client name will not be changed.
        '''
        client_name = arcpy.Parameter(
            displayName="OPTIONAL - Client Name (To be used on the Overview Map's Client Name Field)",
            name="client_name",
            datatype="String",
            parameterType="Optional",
            direction="Input")
        
        workspace.filter.list = ["Local Database", "File System"]
        parameters = [workspace, client_name]
        return parameters

    def execute(self,parameters,messages):
        
        # Pull in the global variables
        global site_layout, overview_layout, lands_file_path, layer, aprx, site_map, inset_map, overview_map, overview_inset_map, file_num, formatted_date 
        try:
            try:
                # Get "file_num" value to help build the pdf file name !!!!Could be changed later to use the file_num from the first tool
                with arcpy.da.SearchCursor(layer, "CROWN_LAND") as cursor:
                    for row in cursor:
                        file_num = row[0]
                        break
                    
            # Error handling - if the script cannot find the Crown Land file number in the Application layer - user will have to manually update the pdf name        
            except arcpy.ExecuteError as e:
                arcpy.AddError("An error occurred while retrieving the file number, you will need to edit the file name manually.")
                arcpy.AddError(e)
                arcpy.AddMessage("File number could not be retrieved, you will need to manually edit the file name of the pdf.")
                # You can choose to handle the error differently based on your requirements.
            
            finally: 
                # Create a valid output folder path
                crown_file_folder = os.path.join(lands_file_path, file_num)
                        
                # Bring in the optional workspace path parameter
                workspace_input = parameters[0].valueAsText
                
                # Bring in the optional client name parameter
                client_name = parameters[1].valueAsText
                
                # Check to see if the workspace input was left black. If the parameter was left blank, set the workspace path to the crown_file_folder 
                # (It wont export to the sub folder that may have been created during the Site_Overview Tool if the file folder already exists.... YET)
                if workspace_input in [None, ""]:
                    arcpy.AddMessage("No output folder selected. Defaulting to the current project folder.")
                    workspace_path = crown_file_folder
                # if the user navigated to a folder, set the workspace path to the folder they selected
                else:
                    workspace_path = workspace_input
                
                # Get the layouts from the site map and the overview map and assign them to variables to be passed into the export function later
                site_layout_obj = aprx.listLayouts(site_layout)[0]
                overview_layout_obj = aprx.listLayouts(overview_layout)[0]
                
                # Assign layout name to a variable to be used later in creating the file name of the pdf
                site_layout_name = site_layout_obj.name  
                overview_layout_name = overview_layout_obj.name


                ########################################################################################################################
                #
                # SETUP THE SITE_LAYOUT FOR THE SITE MAP
                #
                # Make sure the layout is set up with the default settings so the script doesn't fail
                #
                ########################################################################################################################
                
                turn_on_site_map()
                
                # Allow overwriting of output, this is useful when you are making frequent changes to the layout and keep exporting it to the same folder
                arcpy.env.overwriteOutput = True  

                # Build the file name of the output pdf by replacing mmmyyyy found in the layout name with the current month and year, the replacing the words file_num with
                # the file number found in the application layer
                pdf_file_name = f"{site_layout_name.replace('mmmyyyy', formatted_date).replace('FileNo', file_num)}.pdf"
                

                # Call the global exportToPDF function
                exportToPdf(site_layout_obj, workspace_path, pdf_file_name)
                
                
                ########################################################################################################
                #
                # This Section will turn on the imagery layers, turn off unnecessary site map layers, turn off the inset
                # and then export the imagery version of the map
                #
                ########################################################################################################

                turn_on_imagery()
                        
                # Enable overwriting of output
                arcpy.env.overwriteOutput = True  
                
                # Site Map Output PDF Name
                pdf_file_name = f"{site_layout_name.replace('mmmyyyy', formatted_date).replace('FileNo', file_num)}_Imagery.pdf"
                
                # Overview Map Output PDF Name
                overview_pdf_file_name = f"{overview_layout_name.replace('mmmyyyy', formatted_date).replace('FileNo', file_num)}.pdf"
                
                # Call the global exportToPDF function to export the site map
                exportToPdf(site_layout_obj, workspace_path, pdf_file_name)
                
    

                ##RESET THE LAYOUT TO DEFAULT##
                # Turn off the layers "Latest BC RGB Spot" and "Imagery Credits" and turn on the layers "Base Map Auto Scale", "0-7.5K SITE MAP", etc
                # This resets the layout back to normal so if you have to run the tool again, it is ready to go
                
                turn_on_site_map()

                
                # Check to see if the user input a Client Name, if they did, insert the client name into the overview map
                # If they didn't, leave the client name as is
                if client_name in [None, ""]:
                    arcpy.AddMessage("No Client Name was input. Leaving the Client Name on the Overview Map as is.")
                else:
                    arcpy.AddMessage("Client Name was input. Updating the Client Name on the Overview Map.")
                    # Replace the existing Client Name on the Overview map with the parameter input by the user
                    # for lyt in aprx.listLayouts(overview_layout):
                    for elm in overview_layout_obj.listElements("TEXT_ELEMENT"):

                        
                        if elm.name == "ClientName":
                            elm.text = client_name
                    
                # Call the global exportToPDF function to export the overview map
                exportToPdf(overview_layout_obj, workspace_path, overview_pdf_file_name)
    
                arcpy.AddMessage(f"Site Maps, Imagery Maps and Overview Map exported to {crown_file_folder}")

                
                # Open the folder where the pdfs were exported to for preview before sending to Lands Teams
                try:
                    os.startfile(workspace_path)
                except Exception as e:
                    arcpy.AddError(f"Could not open folder {workspace_path}: {e}")
                 
                 
        except arcpy.ExecuteError:
            msgs = arcpy.GetMessages(2)
            arcpy.AddError(msgs)

        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = "PYTHON ERRORS:\nTraceback info:\n" + tbinfo + "\nError Info:\n" + str(sys.exc_info()[1])
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)