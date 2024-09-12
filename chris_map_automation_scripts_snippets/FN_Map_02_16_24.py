
# Author: Chris Sostad
# Based on FOR TOOLS by: Wes Smith
# Ministry of Forests
# Created Date: January 30th, 2024
# Updated Date: 
# Description:
#   This script will create the FN Map and evaluate the intersection with UWR, WMA, and OGMA

# --------------------------------------------------------------------------------
# * SUMMARY

# - INPUTS:
#   - Proponent ID
#   - Cutting Permit ID
#   - Permit Type (CP, RP, FSR, etc.)


# - OUTPUTS

# --------------------------------------------------------------------------------
# * IMPROVEMENTS
# * Suggestions...
# --------------------------------------------------------------------------------
# * HISTORY



#February 2nd, 2024
    # Changed the process of using a def query to find the cutting permit to make temp feature layer
    # Script will start by creating a temp feature layer with the proponent_id and cp_ID
    # Then it will use the temp feature layer to create a new shapefile in the project output folder
    # The temp  shapefile will be used to evaluate the intersection with UWR, WMA, 
    
# February 5th, 2024
    # Finished up to but not including the appraisal area tool

# February 6th, 2024
    # Got most of the way through the appraisal area tool, it is working up to the commented point
    # Accidentally committed to the appraisal area branch, need to merge the two branches together

# February 9th
    # Updated the script into functions
    # Added Create FN Appraisal Area function
    # Added Create Selected FN Features function
    
# February 13th 
    # Fixed attribute table of selected FN Features formatting
    
# February 14th
    # Added the creation of the permit features via def query
    # Created a proper FN folder based on the original format
    # Went through appraisals and updated all of the variables to point to new directories
    # redirected the directory for selected fn features from fgdb to permit dir
    # Added the update application layer connection function - currently not working
    
# February 15th
#     Now saves aprx to the permit_dir folder
#     Saves .kml to the permit_dir folder
#     Update the application layer is working
    


# Need to do
    # Change the name in the layout to application
    # Use update data source and data driven pages to update the layout
    # Go through MakE FN map, find out why cutting permit application is different than A15384_CP_M10_FN
    # Go through the rest of make fn map.

#import all necessary libraries 
import arcpy 
import sys 
import os 
import datetime

# Assign variables

# Variables to be changed each time you run the script
proponent_id = "A15384" #Used to be proponentName
proponent_input = "Conifex" # Not used yet
proponentName = proponent_input
cp_ID = "M32"

# Create the permit_str
permit_str = f"{proponent_id}_CP_{cp_ID}"

# Permit type is used in checkUWR function to determine if the search distance should be -1 or not, CP is -1, all others are 0

permit_type = "CP"

# General Project Variables
aprx = arcpy.mp.ArcGISProject("CURRENT")
workspace = arcpy.env.workspace
fn_consult_map = aprx.listMaps('FN Consult Map')[0] 
# permit_dir = r'\\Spatialfiles2.bcgov\Work\FOR\RNI\RNI\General_User_Data\CSostad\csostad_work_DEVELOPMENT\DMK_Clearances\Arc_Pro_FTA_Clearances_Combined\Notebook_Outputs'
ex_a_fn_folders = r"\\spatialfiles2.bcgov\archive\FOR\RNI\DMK\Library\ExA_FN_Folders"
cut_permit_layer = fn_consult_map.listLayers("Cutting Permit Application")[0]
thlb_data = r'\\spatialfiles2.bcgov\archive\FOR\RNI\DMK\Local_Data\THLB\Consolidated_THLB_2016.gdb\DMK_THLB_over_0'

# Set up the datetime module
now = datetime.datetime.now()
day = now.strftime("%d")
month = now.strftime("%M")
year = now.strftime("%Y")
suffix = year + month + day

# Asssign the FN Consultation Forestry Layer to a variable
# forestry_consultation_data = fn_consult_map.listLayers("FN_Consultation_Forestry_DMK")[0]
forestry_consultation_data = (r'\\Spatialfiles2.bcgov\WORK\FOR\RNI\DMK\Projects\10_FirstNations\FN_Project\First_Nations_Consultation_DMK.gdb\FN_Consultation_Features_DMK\FN_Consultation_Forestry_DMK')

kmz_lyr_file = r'\\spatialfiles2.bcgov\work\FOR\RNI\DMK\Templates_Utilities\FN_Templates\Polygon_FN_Symbology.lyr'


# root directory for all proponents--from which to build
root_dir = os.path.join(r"\\spatialfiles2.bcgov\archive\FOR\RNI\DMK\Library\ExA_FN_Folders", year, "Licensee_NRFL_TESTING")

# Create the Permit root directory This will result in the folder: 2024\Licensee_NRFL\A15384\Canfor 
global permit_root
permit_root = os.path.join(root_dir, proponentName)

# Create the Permit directory which will look like: 2024\Licensee_NRFL\A15384\Canfor\CP\A15384_CP_H47
global permit_dir
permit_dir = os.path.join(permit_root, "CP", permit_str)


# Assign layout names to name variables
layout_portrait_11x17_name = "FN_Consult_Site_Map_11x17_Portrait"

# Access the specific layout
layout_portrait_11x17 = aprx.listLayouts("FN_Consult_Site_Map_11x17_Portrait")[0]  


# Assign Map Frame names to name variables
layers_map_frame = 'Layers Map Frame'

# Ungulate Winter Range Variables
uwr_layer = fn_consult_map.listLayers("Ungulate Winter Range")[0]
uwr_fields = ["UWR_NUMBER", "UWR_UNIT_NUMBER", "TIMBER_HARVEST_CODE"]

# Wildlife Management Area Variables
wma_layer = fn_consult_map.listLayers("Wildlife Management Areas")[0]
wma_fields = ["TAG", "TIMBER_HARVEST_CODE"]

# OGMA Variables
ogma_layer = fn_consult_map.listLayers("OGMA Legal Current")[0]
ogma_fields = ["LEGAL_OGMA_PROVID", "OGMA_TYPE"]


################################################################################################################################
#
# COMMONLY USED FUNCTIONS
#
#############################################################################################################################

# Global exportToPdf function !Not used yet
def exportToPdf(layout, workSpace_path, pdf_file_name):
    out_pdf = f"{workSpace_path}\\{pdf_file_name}"
    layout.exportToPDF(
        out_pdf=out_pdf,
        resolution=300,  # DPI
        image_quality="BETTER",
        jpeg_compression_quality=80  # Quality (0 to 100)
    )

# Function to zoom to feature extent ! Not used yet
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



# Root directory for all proponents
# root_dir = os.path.join(ex_a_fn_folders, now.year, "Licensee_NRFL")
# permit_root = os.path.join(root_dir, proponentName)



###############################################################################################################################
#
# Step 1 - Folder Creation
# Use if-then logic to verify that the intended directory has not been created already
# If the directory exists create a new folder inside the existing folder and append the month, date and the year 
# to the folder name.
#
###############################################################################################################################

arcpy.AddMessage(f"Step 1 - Creating folder for {permit_str}")

# Check the  Directory to see if a folder for this file number already exists, if the folder already exists,
# create a new folder inside the existing folder and append the month, date and the year to the folder name. Then
# set the permit dir variable to point to the new folder. If the folder doesn't exist, create a new one
# with the file number as the name of the folder.

if os.path.isdir(permit_dir):
    arcpy.AddWarning("A directory for this  file has already been created. Creating a new subfolder inside with date appended.")
    print("A directory for this  file has already been created. Creating a new subfolder inside with date appended.")
    
    # Open the folder as an indication that the folder has already been created
    try:
        os.startfile(permit_dir)
    except Exception as e:
        arcpy.AddError(f"Could not open folder {permit_dir}: {e}")
        print(f"Could not open folder {permit_dir}: {e}")
    
    # Create a subfolder with the current month and year appended to its name
    
    # Create the file name for the new subfolder
    new_folder_name = permit_str + "_" + datetime.datetime.now().strftime("%m_%d_%Y")
    arcpy.AddMessage(f"         New folder name is: {new_folder_name}")
    print(f"         New folder name is: {new_folder_name}")
    
    # If dir exists, create a new folder path name with the permit_dir and the new_folder_name
    # It should look like this: 2024\Licensee_NRFL\A15384\Canfor\CP\A15384_CP_H47\A15384_CP_H47_02_15_2024
    new_folder_path = os.path.join(permit_dir, new_folder_name)
    arcpy.AddMessage(f"         New folder path is: {new_folder_path}")
    print(f"         New folder path is: {new_folder_path}")
    
    # Check to see if it already exits, if it doesn't, create it
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)
        arcpy.AddMessage(f"        The date appended folder does not exist. Creating: {new_folder_path}")
        print(f"        The date appended folder does not exist. Creating: {new_folder_path}")
    
    # Update the permit_dir variable to point to the new subfolder
    permit_dir = new_folder_path
    arcpy.AddMessage(f"         permit_dir variable updated to: {permit_dir}")
else:
    # Create the directory for the request number
    os.makedirs(permit_dir)
    arcpy.AddMessage("         There is no existing Directory. New Directory created for " + str(permit_str) in {permit_root})
    print("         There is no existing Directory. New Directory created for " + str(permit_str) in {permit_root})
################################################################################################################################


################################################################################################################################
#
# Step 2 - Create a Def Query on the FTEN Pending Layer
#
#############################################################################################################################

arcpy.AddMessage("Step 2 - Setting up FTEN Layer Polygon")
print("Step 2 - Setting up FTEN Layer Polygon")

# Create a select by attribute query to filter FTEN Harvest Authority (cut_permit_layer) by proponent_id and cp_ID
query = "FOREST_FILE_ID = '" + proponent_id + "' and CUTTING_PERMIT_ID = '" + cp_ID + "'"
print(query)

# Create a definition query on the cut_permit_layer
cut_permit_layer.definitionQuery = query

# Set overwrite to true
arcpy.env.overwriteOutput = True

# Copy the selected feature from the temp layer to the output folder
arcpy.management.CopyFeatures(cut_permit_layer, permit_dir + f"\\{proponent_id}_CP_{cp_ID}_FN.shp")


# Check UWR Function - Start - didnt use temp layer option, might change later
def chekUWR():
    try:
        arcpy.AddMessage("Evaluating intersection with UWR...")
        print("Evaluating intersection with UWR...")

        # Selecting intersecting UWR features
        arcpy.AddMessage("\tSelecting intersecting UWR features...")
        print("\tSelecting intersecting UWR features...")

        if permit_type == "RP" or permit_type == "FSR":
            selection = arcpy.management.SelectLayerByLocation(
                in_layer=uwr_layer,
                overlap_type="INTERSECT",
                select_features=cut_permit_layer,
                selection_type="NEW_SELECTION",
                invert_spatial_relationship="NOT_INVERT"
            )
        else:
            selection = arcpy.management.SelectLayerByLocation(
                in_layer=uwr_layer,
                overlap_type="INTERSECT",
                select_features=cut_permit_layer,
                search_distance=-1,
                selection_type="NEW_SELECTION",
                invert_spatial_relationship="NOT_INVERT"
            )

        # Get attributes and print
        count_result = arcpy.management.GetCount(selection)
        featCount = int(count_result.getOutput(0))

        if featCount == 0:
            arcpy.AddMessage('* No intersecting UWR Features *')
            print('* No intersecting UWR Features *')
        else:
            arcpy.AddWarning('Conflict with UWR:')
            print('\tConflict with UWR:')
            uwr_list = []
            with arcpy.da.SearchCursor(selection, uwr_fields) as cursor:
                for row in cursor:
                    uwr_row_list = [f"\t{field}\t\t\t{row[idx]}" for idx, field in enumerate(uwr_fields)]
                    uwr_list.append(uwr_row_list)
                    arcpy.AddMessage("\n".join(uwr_row_list))
                    print("\n".join(uwr_row_list))


    except Exception as e:
        arcpy.AddError(f"Error in checkUWR: {e}")
        print(f"Error in checkUWR: {e}")

chekUWR()


def checkWMA():
    try:
        arcpy.AddMessage("Evaluating intersection with WMA...")
        print("Evaluating intersection with WMA...")


        # Select intersecting WMA features
        search_distance = -1 if permit_type != "RP" and permit_type != "FSR" else None
        selection = arcpy.management.SelectLayerByLocation(
            in_layer=wma_layer,
            overlap_type="INTERSECT",
            select_features=cut_permit_layer,
            search_distance=search_distance,
            selection_type="NEW_SELECTION",
            invert_spatial_relationship="NOT_INVERT"
        )
        # Get attributes and print
        count_result = arcpy.management.GetCount(selection)
        featCount = int(count_result.getOutput(0))

        #wma_fields = ["TAG", "TIMBER_HARVEST_CODE"]
        wma_list = []

        if featCount == 0:
            arcpy.AddMessage('\t* No intersecting WMA Features *')
            print('\t* No intersecting WMA Features *')
        else:
            arcpy.AddWarning('\tConflict with WMA:')
            print('\tConflict with WMA:')
            with arcpy.da.SearchCursor(selection, wma_fields) as cursor:
                for row in cursor:
                    wma_row_list = [
                        "\tTAG\t\t\t" + row[0],
                        "\tTIMBER_HARVEST_CODE\t" + row[1] + "\n"
                    ]
                    wma_list.append(wma_row_list)
                    arcpy.AddMessage("\n".join(wma_row_list))

    except Exception as e:
        arcpy.AddError(f"Error in checkWMA: {e}")
        arcpy.AddMessage(f"Error in checkWMA: {e}")
        print(f"Error in checkWMA: {e}")
        return e
    return wma_list

# Call the function
wma_result = checkWMA()




# Check appraisal areas function


def appraisalArea():                        # for CP only
    global permit_str
    
    arcpy.AddMessage("Running Appraisal Area Tool...")
    print("Running Appraisal Area Tool...")
    # # Preliminaries

    # Workspace
    appraisalWorkspace = os.path.join(permit_dir, "Appraisal")
    print(appraisalWorkspace)
    arcpy.env.workspace = appraisalWorkspace
    arcpy.env.overwriteOutput = True

    # THLB Data defined at the top of the script
    

    # # Implement if else statement to determine the query based on the proponent ID later on!

    #                     # # Make temp permit feature layer and select attributes by query
    #                     # if proponentName == "BCTS":     # if there is no proponent ID (e.g., BCTS, FLTs, etc.)
    #                     #     temp_query = ex_a_query
    #                     # else:                           # if application has BOTH proponent ID and CP ID
    #                     #     temp_query = query

    # Create a select by attribute query to filter FTEN Harvest Authority (cut_permit_layer) by proponent_id and cp_ID
    # query = "FOREST_FILE_ID = '" + proponent_id + "' and CUTTING_PERMIT_ID = '" + cp_ID + "'"

    # Create the permit_str
    #permit_str = f"{proponent_id}_CP_{cp_ID}"


    # Make feature layer is creating a temp shape so no need to use os.path.join ie. temp_CP_shape = os.path.join(permit_str, "_appraisal_temp.shp")  
    # Create the file name for temp_CP_shape
    
    
    temp_CP_shape = f"{permit_str}_appraisal_temp.shp" # is the same as cut_permit_layer

    # Create the selection layer from the cut permit layer. The reason it creates a temp copy is because appraisal_fc_shape needs to have area field added and calculated
    selection = arcpy.management.MakeFeatureLayer(cut_permit_layer, temp_CP_shape, query, permit_dir) 



    # Catch empty FC
    count_result = arcpy.GetCount_management(temp_CP_shape)
    featCount = int(count_result.getOutput(0))
    if featCount == 0:
        arcpy.AddWarning('\tNo CP features were selected.  Appraisal area evaluation was NOT performed...')
        print('\tNo CP features were selected.  Appraisal area evaluation was NOT performed...')
    else:
        arcpy.AddMessage(f"\t{featCount} CP features were selected.  Running Appraisal area evaluation ...")
        print(f"\t{featCount} CP features were selected.  Running Appraisal area evaluation ...")
        

    # Create the Appraisal FC name
    # For example: Tenures\A15384\Canfor\CP\A15384_CP_H47\A15384_CP_H47_appraisal.shp
    appraisal_fc_name = os.path.join(permit_dir, f"{permit_str}_appraisal.shp")
    print(appraisal_fc_name)



    # Create appraisal shape using minimum bounding geometry Creates a feature class containing polygons which 
    # represent a specified minimum bounding geometry enclosing each input feature or each group of input features.
    arcpy.management.MinimumBoundingGeometry(selection, appraisal_fc_name, "CONVEX_HULL")
    print("Minimum Bounding Geometry created")

    # Add field ("Area_HA")
    field_name = 'Area_HA'
    arcpy.management.AddField(appraisal_fc_name, field_name, "DOUBLE")

    # Calculate field ("Area_HA")
 
    appraisalExpression = '!shape.area@hectares!'
   
    # Calculate field ("Area_HA")
    arcpy.management.CalculateField(appraisal_fc_name, field_name, appraisalExpression, "PYTHON3")
    print("Area calculated")
    # Get area value

    with arcpy.da.SearchCursor(os.path.join(appraisalWorkspace, appraisal_fc_name), field_name) as cursor:    
        for row in cursor:
            cp_appraisal_area = row[0]



    arcpy.AddMessage(f"\tArea - THLB and non-THLB: {cp_appraisal_area} Hectares")
    print(f"\tArea - THLB and non-THLB: {cp_appraisal_area} Hectares")
    appraisal_message_1 = f"\tArea - THLB and non-THLB: {cp_appraisal_area} Hectares"
    print(f"Cp Appraisal Area is: {cp_appraisal_area}")

    # Check if the appraisal area is less than or equal to 7850
    if cp_appraisal_area <= 7850:
        thlb_checked = False
        thlb_test = "PASSED"
    elif cp_appraisal_area > 7850:
        thlb_checked = True
        thlb_test = "FAILED"
        arcpy.AddMessage("\tRemoving non-THLB...")
        print("\tRemoving non-THLB...")
        
        # Select THLB_FACT if it is greater than 0
        '''
        The Timber Harvesting Land Base (THLB) in British Columbia refers to Crown forest land within the timber
        supply area where timber harvesting is considered both acceptable and economically feasible,
        '''
        # Create the expression THLB_FACT > 0
        thlb_expression = '"THLB_FACT" > 0'
        
        # Perform select by location, new selection, THLB_FACT > 0
        thlb_selection = arcpy.management.SelectLayerByAttribute(thlb_data, "NEW_SELECTION", thlb_expression)
        
        # Create the path for the thlb_clip shapefile
        # thlb_fc_clip = os.path.join(appraisalWorkspace,'thlb_clip.shp')
        thlb_fc_clip = os.path.join(permit_dir, "thlb_clip.shp")
        
        # Clip THLB to Appraisal Shape using the appraisal fc name (appraisal.shp) to clip  the new selection and create a new output called thlb_fc_clip
        # Tenures\A15384\Canfor\CP\A15384_CP_H47\A15384_CP_H47_appraisal.shp
        arcpy.analysis.Clip(thlb_selection, appraisal_fc_name, thlb_fc_clip)
        
        # Add area field
        field_name = field_name + '_2'
        arcpy.management.AddField(thlb_fc_clip, field_name, "DOUBLE")
        
        # Calculate field ("Area_HA")
        appraisalExpression = '!shape.area@hectares!'
                                
        # Calculate field ("Area_HA")
        arcpy.management.CalculateField(thlb_fc_clip, field_name, appraisalExpression, "PYTHON3")

        # Get area from thlb_fc_clip using a search cursor and assign it to cp_appraisal_area
        area_2 = 0
        with arcpy.da.SearchCursor(thlb_fc_clip, field_name) as cursor:     
            for row in cursor:
                area_2 += row[0]
        cp_appraisal_area = area_2
        print(f"cp_appraisal_area is: {cp_appraisal_area}")
        
        # Evaluate the THLB area and print out a message
        if cp_appraisal_area <= 7850:
            thlb_test = "PASSED"
        elif cp_appraisal_area > 7850:
            thlb_test = "FAILED"
            arcpy.AddWarning("\t******* Failed ******")
            arcpy.AddMessage("\t******* Failed ******")
            print("\t******* Failed ******")
        appraisal_message_2 = f"Area - THLB only: {cp_appraisal_area} Hectares"
    
        print(appraisal_message_2)


                        
    # Check to the see if the appraisal area is less than or equal to 7850
    if thlb_checked == False:
        appraisal_message = f"Final if thlb checked - \tAppraisal tool was run: {thlb_test}\n\t   {appraisal_message_1}"
        print(f"\tAppraisal tool was run: {thlb_test}\n\t   {appraisal_message_1} last if")
        arcpy.AddMessage(appraisal_message)
        print(appraisal_message)
    else:
                                # appraisal_message = "\tAppraisal tool was run: {}\n\t   {}\n\t   {}".format(thlb_test, appraisal_message_1, appraisal_message_2)
        appraisal_message = f"\tAppraisal tool was run: {thlb_test}\n\t   {appraisal_message_1}\n\t   {appraisal_message_2} last else"
        arcpy.AddMessage(appraisal_message)

    arcpy.management.Delete(appraisal_fc_name)

if cp_ID:                   # Check appraisal area; BSP permits are exempt from appraisal area
    if cp_ID != "BSP":
        appraisalArea()




# Start make_fn_map function

arcpy.AddMessage("\tSelecting features of interest from BCGW and FN Consult Features...")
print("\tSelecting features of interest from BCGW and FN Consult Features...")

 




'''
make_fn_map() is a function that will make a first nations map for the proponent and cutting permit
It will start Selecting features of interest from BCGW and FN Consult Features
It then creates a layer from the selected features for use in a aprx layout
Next, it will Select consultation layers that intersect with permit feature
'''

# Assign the new shapefile name to a variable
fn_permit_fc_name = permit_dir + f"\\{proponent_id}_CP_{cp_ID}_FN.shp" # This is just cut_permit_layer

# Assign the fn_permit_fc_name layer to an object
fn_permit_lyr = fn_consult_map.listLayers(f"{proponent_id}_CP_{cp_ID}_FN")[0]

 ## Select consultation layers that intersect with permit feature
 
# Asssign the FN Consultation Forestry Layer to a variable
# forestry_consultation_data = fn_consult_map.listLayers("FN_Consultation_Forestry_DMK")[0]
# forestry_consultation_data = (r'\\Spatialfiles2.bcgov\WORK\FOR\RNI\DMK\Projects\10_FirstNations\FN_Project\First_Nations_Consultation_DMK.gdb\FN_Consultation_Features_DMK\FN_Consultation_Forestry_DMK')

# Make temp feature
temp_consultation_fc = "temp_consultation_fc"
arcpy.MakeFeatureLayer_management(forestry_consultation_data, temp_consultation_fc)

# Select by location
selection = arcpy.management.SelectLayerByLocation(temp_consultation_fc, "INTERSECT", fn_permit_fc_name, \
                                               "", "NEW_SELECTION")

 #QC
featCount = arcpy.GetCount_management(temp_consultation_fc)
print(f"Consult Feature Count: {featCount}")
arcpy.AddMessage(f"Consult Feature Count: {featCount}")


selected_fn_features_fc_name = permit_str + "_SelectedFeatures.shp"
selected_fn_features_path = os.path.join(permit_dir, selected_fn_features_fc_name)
# Copy FC as SHP
arcpy.management.CopyFeatures(temp_consultation_fc, selected_fn_features_path)


# Delete un-needed fields (the named fields should not be in the KMZ)

arcpy.management.DeleteField(selected_fn_features_fc_name, "PROPO_ID;AGENCY;TEN_TYPE;LOCATION;CNSLT_AREA;CNSLT_YEAR;INTAKE_DAT;\
    SHAPE_Leng;QUARTER;CNSLT_CP;CNSLT_CB;CNSLT_HA;CNSLT_STAR;CNSLT_DONE;CNSLT_LEAD;GIS_ENTRY;GIS_NAME;\
    GIS_UPDATE;GIS_NAME_U;CNSLT_REPO;IMPACT_POT;RIGHTS;TITLE;COMMENT;CB_COUNT;FN_1;FN_2;FN_3;FN_4;FN_5;FN_6;FN_7;IBS;SHAPE_Le_1;SHAPE_Area")



arcpy.management.Delete(temp_consultation_fc)
print("REMOVE MESSAGE - Deleted temp_consultation_fc")


def update_application_layer_connection(map_name, permit_string, permit_directory):
    layer_name = "FN Selected Features"
    layers = map_name.listLayers(layer_name)
    if not layers:
        print(f"No layer named '{layer_name}' found in {map_name.name}.")
        return

    target_lyr = layers[0]
    # Construct the path to the new shapefile
    new_shapefile_path = os.path.join(permit_directory, f"{permit_string}_SelectedFeatures.shp")
    # Prepare the new connection properties
    new_conn_props = {
        'connection_info': {'database': os.path.dirname(new_shapefile_path)},
        'dataset': os.path.basename(new_shapefile_path).replace('.shp', ''),
        'workspace_factory': 'Shape File'
    }
    
    # Update the connection properties
    target_lyr.updateConnectionProperties(target_lyr.connectionProperties, new_conn_props)
    print(f"Data source updated for layer: {layer_name} in map: {map_name.name}")

update_application_layer_connection(fn_consult_map, permit_str, permit_dir)

def fn_make_kmz():

        arcpy.AddMessage("Exporting permit features to KMZ...")

        # preliminaries
        fn_kmz_name = permit_str + ".kmz"
        fn_kmz_output = os.path.join(permit_dir, fn_kmz_name)

        temp_permit_layer = "permit_layer_temp"
        fn_kmz_lyr_output = os.path.join(permit_dir, permit_str + ".lyr")

        # Make Feature Layer
        arcpy.MakeFeatureLayer_management(fn_permit_lyr, temp_permit_layer)

        # Apply Symbology From Layer
        # arcpy.ApplySymbologyFromLayer_management(temp_permit_layer, kmz_lyr_file)
        arcpy.management.ApplySymbologyFromLayer(temp_permit_layer, kmz_lyr_file)
        arcpy.AddMessage("\tSymbology applied to layer--exporting to KMZ...")
        print("\tSymbology applied to layer--exporting to KMZ...")
        # Save To Layer File
        # arcpy.SaveToLayerFile_management(temp_permit_layer, fn_kmz_lyr_output, version="CURRENT")
        arcpy.management.SaveToLayerFile(temp_permit_layer, fn_kmz_lyr_output)

        # Layer To KML
        fn_kmz_scale = "20000"
        arcpy.LayerToKML_conversion(temp_permit_layer, fn_kmz_output, fn_kmz_scale, "false", \
                                    "DEFAULT", "1024", "96", "CLAMPED_TO_GROUND")

        arcpy.Delete_management(temp_permit_layer) 

        arcpy.AddMessage("\tFN Map and data has been prepared.")
        print("\tFN Map and data has been prepared.")
fn_make_kmz()





# Save the .aprx to the permit_dir folder using the permit_str as the file name
aprx.saveACopy(os.path.join(permit_dir, f"{proponent_id}_CP_{cp_ID}_FN.aprx"))
arcpy.AddMessage(f"Map saved to: {permit_dir}")
print(f"Map saved to: {permit_dir}")



print("Script Complete")