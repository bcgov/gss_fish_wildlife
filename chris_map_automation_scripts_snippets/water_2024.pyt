#===========================================================================
# Script name: water.pyt
# Initial Author: Jordan Foy
# Rewritten by: Chris Sostad 08/23/23
# Created on: 08/10//2021
# Edited on: 11/01/2023
#
# Description: This tool fully automates the water plat creating process. The user will enter the water file number, licence number, and PID number 
# (if they have it) and the tool wil create the plat for them.
#
# Dependancies: APRX must contain one layout and one map, must use in ArcGIS Pro 
# 
# Script tool parameters
# 
# water_file: string, type: required, direction: input
# licensenum: string, type: required, direction: input 
#
# Limitations: 
# Can only run this script within the template, as it is coded to run with
# certain data. Must have a fresh, unaltered template. 
#
#  
# 
#============================================================================

# Dec 07, 2023
    # Added error handling for Null Values in the PW Application Layer     

# Dec 15, 2023
    # Added Function to Zoom to Feature Extent
    #lsdkjflksdjlskdjfdlk


# December 18, 2023
    # Need to adjust file name based on if PD/PW is chosen
    # Moved variables to global
    # Removed If/Or to find if folder exists. Just overwrite the folder, the maps are so quick to produce, it's easier this way
    # Added functionality, if you don't have a PID, it is now optional. Script will use search by location on the PWD or PD (depending what you entered as Pod Type)
    # to find the appurtenance in which the PWD is contained by.

# January 4th, 2024
    # Add automatic file labelling for PDF export
    # Need to add a feature to turn off "appurtenance" in legend because it is automatically in there as a graphic
    # Need to add a feature, if PW is selected as POD type, then turn off the layers for PD and turn on the PW Application layer. If PD is selected as the POD Type, turn off the PW Layers and turn on the PD Application layer
    # Added error handling for cases where no PID is given, the script will attempt to find the appurtenance using a select by location "CONTAINS" on the Point of DIVERSION application layer

# February 20th, 2024
    # Added regional folder handling for the outputs. The script will now save the outputs to the correct regional folder based on the region the chosen pd is in.
    # Added further handling for WTN that are null. If the WTN is null, the script will simply add "unknow" to the label.

#import necessary libraries - 
import arcpy 
import traceback 
import sys 
import os 
import datetime





# Set the default workspace
aprx = arcpy.mp.ArcGISProject("CURRENT")

# Take the only map in the project and assign it to a variable
mapx = aprx.listMaps()[0]

# Take the only layout in the project and assign it to a variable
lyt  = aprx.listLayouts("FCBC_PW_PLAT_mmmyyyy_85X11")[0]

# The name of the layout as a string
layout_name = "FCBC_PW_PLAT_mmmyyyy_85X11"

# Get the current date to be used in naming of the pdf 
current_date = datetime.datetime.now()
formatted_date = current_date.strftime("%b%Y")

#creates a map frame variable for the map elements in the Layers Map Frame used in the layout
mf = lyt.listElements("mapframe_element", "Layers Map Frame")[0]

# Define the path to the water file folder
# water_file_path = "\\\\spatialfiles.bcgov\\work\\lwbc\\nsr\\Workarea\\fcbc_prg\\FCBC\water_files\Authorizations\Water"  MOF path for testing
water_file_path = "\\\\spatialfiles.bcgov\work\srm\wml\Workarea\Authorizations\Water"  # WLRS path for production

# Set the PMBC Parcel Cadastre as the layer to be used for a definition query 
pmbc_layer = mapx.listLayers("PMBC Parcel Cadastre - Fully Attributed - Outlined")[0]

# Set the 4 possibilties of Point of Diversion layers to variables
pwd_app_layer = mapx.listLayers("Point of Well Diversion - Application")[0]
pwd_lic_layer = mapx.listLayers("PWD - Licenced")[0]
pod_app_layer = mapx.listLayers("PD - Application")[0]
pod_lic_layer = mapx.listLayers("PD - Licenced")[0]

# Define the land layer
land_districts_layer = mapx.listLayers("Land Districts")[0]

# Define the water wells layer
water_wells_layer = mapx.listLayers("Water Wells")[0]

# Define the Appurtenance layer
appurtenance_layer = mapx.listLayers("appurtenance")[0]

# Define the Natural Resource Regions layer
regions_layer = mapx.listLayers("Natural Resource Regions - Outlined")[0]

# Global exportToPdf function
def exportToPdf(layout, workSpace_path, pdf_file_name):
    out_pdf = f"{workSpace_path}\\{pdf_file_name}"
    layout.exportToPDF(
        out_pdf=out_pdf,
        resolution=300,  # DPI
        image_quality="BETTER",
        jpeg_compression_quality=80  # Quality (0 to 100)
    )

# Update Layer Connection Function
def update_layer_connection(layer_name, map_name, full_path):
    arcpy.AddMessage("                                                          ")
    arcpy.AddMessage("Step 6: Running update_layer_connection function.")
    arcpy.AddMessage("        (If you receive a 'List Index Out of Range' Error, make sure you have an existing appurtenance layer)")
   
    """
    Updates the data source connection properties for a specified layer in a given map.

    This function is designed to update the connection properties of a layer in a map. The function first retrieves the target layer from the map. It then logs the original 
    connection properties of this layer for reference. The new connection properties are set up to point to a shapefile 
    named after the layer itself, which found in a specific folder defined by a combination of 
    'lands_file_path' and 'file_num' variables.

    Parameters:
    layer_name (str): The name of the layer whose connection properties are to be updated.
    map_name: The map object containing the target layer.


    """
    
    global mapx
    
    # The below code is used to update the data source for the  layer. 
    # First find the application layer 
    target_lyr = map_name.listLayers(layer_name)[0]
    


    # Set a variable that represents layer connection Properties
    origConnPropDict = target_lyr.connectionProperties
    arcpy.AddMessage(f"        Old {target_lyr} connection properties retrieved")
    
    # Set new connection properties based on the layer shapefile exported in earlier step 
    newConnPropDict = {'connection_info': {'database': full_path},
                    'dataset': layer_name,
                    'workspace_factory': 'Shape File'}
    arcpy.AddMessage(f"        New {target_lyr} layer connection properties updated")
    
    # Update connection properties 
    target_lyr.updateConnectionProperties(origConnPropDict, newConnPropDict)
    arcpy.AddMessage(f"        {target_lyr} source updated") 


# Function to zoom to feature extent
def zoom_to_feature_extent(map_name, map_frame, layer_name, zoom_factor, layout_name):
    global mapx
    ''' This function will focus the layout on the selected feature and then pan out a given zoom factor to show the surrounding area. 
    the function validates the existence of the specified layout, map frame, and layer. If any of these do not exist, 
    it logs an appropriate error message and exits. Once the existence of these elements is confirmed, the function retrieves 
    the current extent of the specified layer within the map frame. It then adjusts this extent based on the provided zoom factor, 
    effectively zooming out to provide a broader view that includes the surrounding area of the feature. The adjusted 
    extent is applied to the map frame's camera, updating the view in the layout. The function also handles any exceptions
    that may occur during the zooming process, logging an error message if an issue arises.'''
    arcpy.AddMessage("                                                       ")
    arcpy.AddMessage("Step 7: Running zoom_to_feature_extent function")
    
    # Verify layout is present and assign it to a variable
    layouts = aprx.listLayouts(layout_name)
    if not layouts:
        arcpy.AddError(f"No layout found with the name: {layout_name}")
        return
    lyt_name = layouts[0]

    # Verify map frame exists
    map_frames = lyt_name.listElements("MAPFRAME_ELEMENT", map_frame)
    if not map_frames:
        arcpy.AddError(f"No map frame found with the name: {map_frame} in layout: {layout_name}")
        return
    mf = map_frames[0]

    # Check if the map and layer exist
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

    # Get the current extent and expand the Xmax/Min Ymax/Min by the given zoom factor
    try:
        current_extent = mf.getLayerExtent(lyr_name, False, True)
        x_min = current_extent.XMin - (current_extent.width * zoom_factor)
        y_min = current_extent.YMin - (current_extent.height * zoom_factor)
        x_max = current_extent.XMax + (current_extent.width * zoom_factor)
        y_max = current_extent.YMax + (current_extent.height * zoom_factor)
        new_extent = arcpy.Extent(x_min, y_min, x_max, y_max)
        mf.camera.setExtent(new_extent)
        arcpy.AddMessage(f"        Zoomed to {layer_name} with zoom factor {zoom_factor}")
    except Exception as e:
        arcpy.AddError(f"Error in zooming to extent: {str(e)}")



# Function to check if a layer exists in a map (Not used in this script)
def layer_exists(layer_name):
    """Check if a layer exists in the global map object."""
    global mapx  # Reference the global map object
    for layer in mapx.listLayers():
        if layer.name == layer_name:
            arcpy.AddMessage(f"{layer_name} found in map.")
            return True
    return False

class Toolbox(object):
    def __init__(self):
        """Define the toolbox (name of toolbox is name of the file)"""
        self.label = "Toolbox"
        self.alias = ""

         
        self.tools = [FullWaterPlatTool, ExportSingleLayout, ExportWaterPlat]



class FullWaterPlatTool(object):
    """This tool combines the 3 different steps into one tool. Hopefully it 
    will complete all three steps in one go. Hit run and pour yourself a coffee.""" 
    
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Full Water Authorization Tool for PD or PW - Development"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """Defines the parameters (this is what the 
        user enters before they run the tool)"""
        
        proj_req = arcpy.Parameter(
            displayName = "Project Request Number",
            name="proj_req",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        water_file = arcpy.Parameter(
            displayName = "Water File Number",
            name="water_file",
            datatype="String",
            parameterType="Required",
            direction="Input")
    
        licensenum = arcpy.Parameter(
            displayName = "License Number",
            name="licensenum",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        pid = arcpy.Parameter(
            displayName = "Optional but preferred - PID Number as Text i.e. 017475741",
            name="pid",
            datatype="String",
            parameterType="Optional",
            direction="Input")

        pod_type = arcpy.Parameter(
            displayName="POD Type",
            name="pod_type",
            datatype="String",
            parameterType="Required",
            direction="Input")

        pw_list = arcpy.Parameter(
            displayName = "PW List (comma separated)",
            name="pw_list",
            datatype="String",
            parameterType="Optional",
            direction="Input")

        label_multiple_pw = arcpy.Parameter(
            displayName = "Label Plat with Multiple PW's",
            name="label_multiple_pw",
            datatype="Boolean",
            parameterType="Optional",
            direction="Input")

        well_tag_num = arcpy.Parameter(
            displayName="Well Tag Number eg.128302 (Optional if PID is not provided)",
            name="well_tag_num",
            datatype="String",
            parameterType="Optional",
            direction="Input")

        pod_type.filter.type = "ValueList"
        pod_type.filter.list = ["PD", "PW"]

        parameters = [proj_req, water_file, licensenum, pid, pod_type, pw_list, label_multiple_pw, well_tag_num]
        
        return parameters

    def execute(self, parameters, messages):
        try:
            # Retrieve parameter values
            proj_req = parameters[0].valueAsText
            water_file = parameters[1].valueAsText
            licence_num = parameters[2].valueAsText
            pid = parameters[3].valueAsText
            pod_type = parameters[4].valueAsText
            pw_list = parameters[5].valueAsText
            label_multiple_pw = parameters[6].value
            well_tag_num = parameters[7].valueAsText
            
            #########################################################################################################
            #
            # Section 1 - Find the Point of Diversion on either the PD or PW Application Layer
            #
            #########################################################################################################
            

            # Step 1: Set the path to the water file folder
            # path = os.path.join(water_file_path, str(water_file))

            #Add a message confirming the file path
            # arcpy.AddMessage(f"Step 1: The file path is {path} \n")
            
            
            # Step 2: 
            
            # Use a conditional statement to determine which layer to work on based on if the user entered PW or PD in the pod type, and assign that to a variable name
            # Use Try/Except to catch errors
            try:
                
                if pod_type == "PW":
                    # If the podtype was entered as PW, then assign the Point of Well Diversion Application layer to a variable that will be used throughout
                    # the rest of the script.
                    chosen_pdpw_layer = mapx.listLayers("Point of Well Diversion - Application")[0]
                    
                    # Turn on the Point of Well Diversion Layer
                    chosen_pdpw_layer.visible = True
                    
                    # Turn off the Point of Diversion Application Layer
                    pod_app_layer.visible = False
                    
                    # Apply a definition query to the layer based on the water file number entered by the user
                    chosen_pdpw_layer.definitionQuery= "FILE_NUMBER = '" + water_file + "'"
                    arcpy.AddMessage(f"Step 2: Applying Definition Query PWD application, Definition query is: {chosen_pdpw_layer.definitionQuery}")
                    
                    # Confirm if there are any records on that layer that meet the query
                    result = arcpy.GetCount_management(chosen_pdpw_layer)
                    record_count = int(result.getOutput(0))
                    
                    # If no records are found, display a message and proceed to check the PW -  Licenced Layer
                    if record_count == 0:
                        # if record count on the last query was 0, then change the chosen_pdpw_layer to the PWD Licenced layer and apply the same definition query
                        chosen_pdpw_layer = mapx.listLayers("PWD - Licenced")[0]
                        chosen_pdpw_layer.definitionQuery= "FILE_NUMBER = '" + water_file + "'"
                        arcpy.AddWarning("        Step 2: Your Definition Query on the PWD Application Layer resulted in Zero records. Checking PWD Licenced Layer. Be sure to confirm your results")

                    
                    else:
                        arcpy.AddMessage(f"Step 2: Definition Query found {record_count} records on PWD Application that meet the query.")
                
                # If PW not chosen as the POD Type (PD was chosen), then use the PD layers
                else: 
                    # Assign the Point of Diversion Application layer to a variable that will be used throughout the rest of the script.
                    chosen_pdpw_layer = mapx.listLayers("PD - Application")[0]
                    
                    # Apply a definition query to the layer based on the water file number entered by the user
                    chosen_pdpw_layer.definitionQuery= "FILE_NUMBER = '" + water_file + "'"
                    arcpy.AddMessage(f"Step 2: Applying Definition Query to PD Application, Definition query is: {chosen_pdpw_layer.definitionQuery}")
                    
                    # Turn on the PD Application Layer
                    chosen_pdpw_layer.visible = True
                    
                    # Turn off the PWD Layer
                    pwd_app_layer.visible = False
                    
                    
                    
                    # Confirm if there are any records that meet the query
                    result = arcpy.GetCount_management(chosen_pdpw_layer)
                    record_count = int(result.getOutput(0))
                    
                    # If no records are found, display a message and proceed to check the POD Licenced Layer
                    if record_count == 0:
                        arcpy.AddWarning("        Step2: Your Definition Query on the PD Application Layer resulted in Zero records. Checking PD Licenced Layer. Be sure to confirm your results")
                        
                        # Assign POD Licenced layer to the variable that will be used throughout the rest of the script
                        chosen_pdpw_layer = mapx.listLayers("PD - Licenced")[0]
                        
                        # Apply a definition query to the layer based on the water file number entered by the user
                        chosen_pdpw_layer.definitionQuery= "FILE_NUMBER = '" + water_file + "'"
                        

                    else:
                        arcpy.AddMessage(f"Step 2: Definition Query on PD Application found {record_count} records on PWD Application that meet the query.")

            except Exception as e:
                arcpy.AddError(f"Error in selecting the PD/PW layer, it doesn't exist: {e}")
                
                raise e
                   
            # Create a map frame variable for the map elements in the Layers Map Frame in the layout
            mf = lyt.listElements("MAPFRAME_ELEMENT", "Layers Map Frame")[0]
            
            # Use the map frame and camera object to zoom the the extent of the Point of Diversion layer 
            mf.camera.setExtent(mf.getLayerExtent(chosen_pdpw_layer, False, True))
            
        
            #########################################################################################################
            #
            # Section 2 - Appurtenance
            #
            # Script will check if you have entered a PID. If you have a PID it will use a definition query on the PMBC layer
            # to find the appurtenance and save it to the file path. If you don't have a PID, it will use a select by location using the 
            # Water Wells layer to find the appurtenance
            #
            #########################################################################################################
        
        
            
            # Set overwrite to true. Used during development. Can delete or comment out after testing is complete.
            arcpy.env.overwriteOutput = True
            
            # Step 3:
            
            # Perform a select by location to find which region the chosen_pdpw_layer is in. Using the Natural Resource Regions\Natural Resource Regions - Outlined
            # as the input layer and the chosen_pdpw_layer as the select features (intersect). 
            
            # Step 3: Select by Location to find the Natural Resource Region
            
            arcpy.SelectLayerByLocation_management(in_layer=regions_layer,
                                                overlap_type="INTERSECT",
                                                select_features=chosen_pdpw_layer,
                                                search_distance="",
                                                selection_type="NEW_SELECTION")

            # Step 4: Use a Search Cursor to Extract the Region Name
            region_name = None
            with arcpy.da.SearchCursor(regions_layer, ["Region_Name"]) as cursor:
                for row in cursor:
                    region_name = row[0]
                    break  # Assuming one region is selected, break after the first match

            # Mapping of region names to folder names
            region_to_folder = {
                "Omineca Natural Resource Region": "Omineca",
                "Cariboo Natural Resource Region": "Cariboo",
                "Skeena Natural Resource Region": "Skeena",
                "Kootenay Natural Resource Region": "Kootenay-Boundary",
                "Thompson-Okanagan Natural Resource Region": "Thompson-Okanagan",
                "Northeast Natural Resource Region": "NorthEast",
                "West Coast Natural Resource Region": "West Coast",
                "South Coast Natural Resource Region": "South Coast"
            }

            # Determine the correct output folder based on the region name
            output_folder_name = region_to_folder.get(region_name, "Unknown")
            output_path = os.path.join(water_file_path, output_folder_name) #looks like this: Authorization\Water\Omineca

            # Step 5: Create the output folder if it doesn't exist
            if not os.path.exists(output_path):
                os.makedirs(output_path)
                arcpy.AddMessage(f"Output folder created: {output_path}")
            else:
                arcpy.AddMessage(f"Output folder already exists: {output_path}")

            # Step 6: Create the output folder with water_file + proj_req
            folder_name = f"{water_file}_{proj_req}"
            file_num_output_path = os.path.join(output_path, folder_name)
            
            # Create the folder if it doesn't exist
            if not os.path.exists(file_num_output_path):
                os.makedirs(file_num_output_path)
                arcpy.AddMessage(f"Output subfolder created: {file_num_output_path}")
            else:
                arcpy.AddMessage(f"Output subfolder already exists: {file_num_output_path}")

            # Allow overwrite of existing appurtenance file - used for testing during the development of this code
            arcpy.env.overwriteOutput = True
            
            
            
            if pod_type == "PW":
                
                arcpy.AddMessage("        Locating appurtenance  using select by location on the Point of Well Diversion - Application layer")
                                
                # Call the select by location function to find which PMBC parcel the pw application is in. This is the alternative to finding the appurtenance using
                # the PID number. Be sure to check your results
                arcpy.management.SelectLayerByLocation(
                in_layer= pmbc_layer, 
                overlap_type="CONTAINS",
                select_features=chosen_pdpw_layer,
                search_distance=None,
                selection_type="NEW_SELECTION",
                invert_spatial_relationship="NOT_INVERT"
                )
                
                # Copy the selected features and send them as a shapefile to the water_file directory for this request                
                arcpy.management.CopyFeatures(pmbc_layer, file_num_output_path  + "\\appurtenance.shp")
                arcpy.AddMessage("        Select by location on PW Groundwater Point Of Diversion - Application SUCCESS! appurtenance.shp created by using Select by Location 'Contains' PW Application")
                arcpy.AddMessage("        *** Be Sure to Confirm Your Results***")
            
            # If there was no PID entered and the pod type is PD    
            if pid == None and pod_type == "PD":
                arcpy.AddMessage("Step 4: Creating Appurtenance Polygon.... The PID parameter is null. Attempting to locate the appurtenance using Select by Location 'Contains' 'Point of Diversion - Application'.")    
                
                # #Set the definition query on the Water Wells layer
                # POD_expression = "FILE_NUMBER = '" + water_file + "'"
                # PD - Application.definitionQuery= water_wells_expression 
                # arcpy.AddMessage(f"        Water Wells definition query set to {water_wells_expression}")
                
                
                # Confirm if there are any records that meet the query
                result = arcpy.GetCount_management(chosen_pdpw_layer)
                record_count = int(result.getOutput(0))
                
                # If no records are found, display a message and proceed to check the Licenced Layer
                if record_count == 0:
                    arcpy.AddWarning("        Step 4: Your Definition Query on the PD Application Layer resulted in Zero records. Out of options. Email the Water Authorizations Team to confirm the File Number?")
           
                else:
                    arcpy.AddMessage(f"Step 4: Definition Query found {record_count} records on PD Application that meet the query, attemping to select by location to find appurtenance.")
            
                

                
                # Call the select by location function to find which PMBC parcel the water well is in. This is the alternative to finding the appurtenance using
                # the PID number. Be sure to check your results
                arcpy.management.SelectLayerByLocation(
                in_layer= pmbc_layer, 
                overlap_type="CONTAINS",
                select_features=chosen_pdpw_layer,
                search_distance=None,
                selection_type="NEW_SELECTION",
                invert_spatial_relationship="NOT_INVERT"
                )
                
                # Copy the selected features and send them as a shapefile to the water_file directory for this request                
                arcpy.management.CopyFeatures(pmbc_layer, file_num_output_path  + "\\appurtenance.shp")
                arcpy.AddMessage("        Select by location on Point Of Diversion - Application SUCCESS! appurtenance.shp created by using Select by Location 'Contains' POD Application")
                arcpy.AddMessage("        *** Be Sure to Confirm Your Results***")


            if pid != None:
                
                # If you have the PID number, then use a definition query on the PMBC layer to find the appurtenance 
                pmbc_expression = "PID_NUMBER = '" + pid + "'"
                pmbc_layer.definitionQuery = pmbc_expression
                arcpy.AddMessage(f"Step 4: PID Number Entered. Attempting to locate the appurtenance. With {pmbc_expression}")
                
                # Step 5: Create the appurtenance shapefile and save to file path
            
                # This copies the selected features and sends them as a shapefile to the water_file directory for this request                
                arcpy.management.CopyFeatures(pmbc_layer, file_num_output_path + "\\appurtenance.shp")
                arcpy.AddMessage("Step 5: Appurtenance created via PMBC layer and saved to folder")

         
            # Reset the PMBC layer definitionQuery to none to prevent errors later on
            pmbc_layer.definitionQuery = None
            
            
            # Step 6 - Update the source of the appurtenance layer and zoom to feature on the layout
            
            update_layer_connection("appurtenance", mapx, file_num_output_path )
            arcpy.AddMessage("        Connection Properties Updated for appurtenance layer")
            
            # Step 7 - Focus the map frame in the layout onto the appurtenance layer and then zoom out to 80% border
            
            # Call the Zoom to feature extent function with 0.8 zoom factor
            zoom_to_feature_extent("Layers", "Layers Map Frame", "appurtenance", 0.1, layout_name)
            

        
            ##########################################################################################################################
            #
            # Section 3 - Labels
            #
            ##########################################################################################################################
            arcpy.AddMessage("                                                 ")
            arcpy.AddMessage("Step 8: Labelling the Water Plat")
            
            # Would be better if lands_districts was a temp layer
            
            # Clear any existing selection on the land_districts_layer
            arcpy.management.SelectLayerByAttribute(land_districts_layer, "CLEAR_SELECTION")        
            
            
            # Perform Select Layer by Location to find which land district the POD is in
            arcpy.management.SelectLayerByLocation(land_districts_layer, overlap_type="INTERSECT", select_features = pwd_app_layer,
                                                    search_distance="", selection_type="NEW_SELECTION",
                                                    invert_spatial_relationship="NOT_INVERT")

            # Set overwrite to true
            arcpy.env.overwriteOutput = True

            arcpy.management.CopyFeatures(land_districts_layer, file_num_output_path  + "\\land_districts.shp")
            arcpy.AddMessage("        land_districts.shp created")

            # Change the licence label
            for elm in lyt.listElements("TEXT_ELEMENT"):
                if elm.name == "LICENCE_NUMBER":
                    elm.text = licence_num
                    arcpy.AddMessage("        Updating Text Elements:")
                    arcpy.AddMessage("        Licence Number Changed")

            # Change the land district setting
            ld_name = None  # Initialize ld_name outside of the with block

            ld = os.path.join(file_num_output_path,  "land_districts.shp")
            arcpy.AddMessage(f"        Land District Layer path: {ld}")
            with arcpy.da.SearchCursor(ld, "LAND_DISTR") as cursor:
                for row in cursor:
                    ld_name = f"{row[0]}"

            for elm in lyt.listElements("TEXT_ELEMENT"):
                if elm.name == "LAND_DISTRICT":
                    elm.text = ld_name
                    arcpy.AddMessage("        Land District text element changed")
                    
                elif elm.name == "FILE_NUMBER":
                    elm.text = str(water_file)  # Ensure water_file is converted to a string
                    arcpy.AddMessage("        File Number text element changed")
            
       
            for elm in lyt.listElements("TEXT_ELEMENT"):
                try:
                    if elm.name == "PW":
                        if pod_type == "PW":
                            if label_multiple_pw and pw_list:
                                elm.text = pw_list
                                arcpy.AddMessage("        Multiple PW text element set")
                            else:
                                elm.text = "PW"
                        elif pod_type == "PD":
                            elm.text = "PD"
                            arcpy.AddMessage("            PW vs PD text element changed")

                    if elm.name == "PW_AND_WTN":
                        pod_num = None
                        wt_num = None

                        # Check POD_NUMBER
                        with arcpy.da.SearchCursor(chosen_pdpw_layer, "POD_NUMBER") as cursor:
                            for row in cursor:
                                if row[0] is not None:
                                    pod_num = u'{0}'.format(row[0])
                                    break
                                else:
                                    arcpy.AddMessage("           POD Number is Null!")

                        if pod_type == "PW" and not label_multiple_pw:
                            # Check WELL_TAG_NUMBER for PW
                            with arcpy.da.SearchCursor(chosen_pdpw_layer, "WELL_TAG_NUMBER") as cursor:
                                for row in cursor:
                                    if row[0] is not None:
                                        wt_num = u'{0}'.format(row[0])
                                        break
                                    else:
                                        arcpy.AddMessage("            Well Tag Number is Null!")

                            # Update the label with POD_NUMBER and WTN for PW
                            if pod_num is not None and wt_num is not None:
                                elm.text = f"{pod_num} WTN({int(float(wt_num))})"
                                arcpy.AddMessage("        PW and WTN text element changed")

                        elif pod_type == "PD":
                            # Update the label with only POD_NUMBER for PD
                            if pod_num is not None:
                                elm.text = pod_num
                                arcpy.AddMessage("            PD text element changed")

                except Exception as e:
                    arcpy.AddMessage(f"Error encountered: {e}")
                                    
                                    
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





#Export any single layout by choosing your destination folder
class ExportSingleLayout(object):
    def __init__(self):
        """This tool will export a single layout to a chosen directory. You should only have one layout in your project"""
        self.label = "ExportSingleLayout"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 

        # Parameter for the directory where the PDF will be stored
        workSpace = arcpy.Parameter(
            displayName = "Navigate to the output folder for your PDF",
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )
        
        # Parameter for the name of the exported PDF
        pdfName = arcpy.Parameter(
            displayName = "Enter the name for your PDF",
            name="pdfName",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        pdfName.value = "ChangeMe"

        # Add a filter to allow only folders to be chosen
        workSpace.filter.list = ["Local Database", "File System"]

        # List of parameters
        parameters = [workSpace, pdfName]

        return parameters

    def execute(self,parameters,messages):
        try:
            workSpace_path = parameters[0].valueAsText
            pdf_file_name = parameters[1].valueAsText

            # Check if '.pdf' extension is already included in the name, if not add it
            if not pdf_file_name.endswith('.pdf'):
                pdf_file_name += '.pdf'

            # Set the path to the current project
            aprx = arcpy.mp.ArcGISProject("CURRENT")

            # Access the layout in the project
            layout = aprx.listLayouts()[0]

            # Export the layout to a PDF with the specified settings
            layout.exportToPDF(
                out_pdf=workSpace_path + "\\" + pdf_file_name,
                resolution=300,  # DPI
                image_quality="BETTER",
                jpeg_compression_quality=80  # Quality (0 to 100)
            )

            arcpy.AddMessage(f"Layout exported to {workSpace_path}")

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

# This tool will export the Water plat the the existing working directory for the water file. It will overwrite any existing files with the same name and will name the PDF based on the file number and the date
class ExportWaterPlat(object):
    def __init__(self):
        self.label = "Export Water Plat to the water file folder and name it appropriately"
        self.description = ""
        self.canRunInBackground = False


# This parameter prompts the user to enter the folder where they want the maps exported to. If it is left blank, it will default to the file folder in the Lands Folder
    def getParameterInfo(self):
        workspace = arcpy.Parameter(
            displayName = "OPTIONAL - Select the output folder for the Water Plat (Will Overwrite) - Default is Water_Files - File Number folder",
            name="workspace",
            datatype="DEWorkspace",
            parameterType="Optional",
            direction="Input"
        )
        
  
        workspace.filter.list = ["Local Database", "File System"]
        parameters = [workspace]
        return parameters

    def execute(self,parameters,messages):
        
        # Pull in the global variables
        global lyt, mapx, file_num, formatted_date, layout_name, water_file_path 
        try:
            try:
                # Scan the map elements for the file number to be used in the pdf file name
                for elm in lyt.listElements("TEXT_ELEMENT"):
       
                    if elm.name == "FILE_NUMBER":
                        # Assign the file number to a variable
                        file_num = elm.text
                        
                        # Assign in to a variable to be used in the file name  
                        arcpy.AddMessage("File Number text element assigned to variable")
                        break
                    
            # Error handling - if the script cannot find the file number in the layout - user will have to manually update the pdf name        
            except arcpy.ExecuteError as e:
                arcpy.AddError("An error occurred while retrieving the file number, you will need to edit the file name manually.")
                arcpy.AddError(e)
                arcpy.AddMessage("File number could not be retrieved, you will need to manually edit the file name of the pdf.")
                # You can choose to handle the error differently based on your requirements.
            
            finally: 
                # Create a valid output folder path
                plat_output_folder = os.path.join(water_file_path, file_num)
                
                
                # Bring in the optional workspace path parameter
                workspace_input = parameters[0].valueAsText
                
 
                # Check to see if the workspace input was left black. If the parameter was left blank, set the workspace path to the water_file_folder 
                # If the workspace that the user input is in the list [None, or blank] then assign the plat_output_folder to the workspace_path variable
                if workspace_input in [None, ""]:
                    arcpy.AddMessage("No output folder selected. Defaulting to the current project folder.")
                    workspace_path = plat_output_folder
                
                # if the user navigated to a folder, set the workspace path to the folder they selected
                else:
                    workspace_path = workspace_input

                
                # Allow overwriting of output, this is useful when you are making frequent changes to the layout and keep exporting it to the same folder
                arcpy.env.overwriteOutput = True  

                
                # Check to see if you are creating a PW or PD Plat. This is done by scanning the layout for the PW or PD text element. 
                # If PW is found, read the text element and assign the result to the variable pw_or_pd
                
                try:
                    # Scan the map elements for the pod type to be used in the pdf file name
                    for elm in lyt.listElements("TEXT_ELEMENT"):
                        
                        
                        # Scan the text elements in the layout. If one of the text elements is PW, then scan the elm.text and assign the result to the variable named pw_or_pd
                        if elm.name == "PW":
                            pw_or_pd = elm.text
                            
                            arcpy.AddMessage(f" PW found in layout contents as {pw_or_pd}, PDF file name will indicate {pw_or_pd}")
                            
                            # Replace the month and year, replace PLAT with the file number and PW with the pod type
                            pdf_file_name = f"{layout_name.replace('mmmyyyy', formatted_date).replace('PLAT', file_num).replace('PW', pw_or_pd)}.pdf"
                        

                            arcpy.AddMessage("Pdf file name updated")

                                                  
                        
                # Error handling - if the script cannot find the file number in the layout - user will have to manually update the pdf name        
                except arcpy.ExecuteError as e:
                    arcpy.AddError("An error occurred while determining if the file name should contain PW or PD, you will need to edit the file name manually.")
                    arcpy.AddError(e)
                    arcpy.AddMessage("PW or PD could not be retrieved, you will need to manually edit the file name of the pdf.")

                

                # Call the global exportToPDF function
                exportToPdf(lyt, workspace_path, pdf_file_name)
                
                arcpy.AddMessage(f"Pdf Export is complete. Opening the folder for you to double check the output")
                

                # Open the folder where the pdfs were exported to for preview before sending to Water Team
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
