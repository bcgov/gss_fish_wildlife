# Author: Chris Sostad
# Based on Model by: Viktor Brumovsky
# Ministry of Forests
# Created Date: February 21st, 2024
# Updated Date: 
# Description:
#   This script will create a manual Exhibit A fors time when SNCSC is broken or when the user wants to create an Exhibit A map themselves.

# --------------------------------------------------------------------------------
# * SUMMARY

# The Exhibit A script is integrated into an ArcGIS Pro Project designed to function as a single space to assist with FTA Clearnces. This template allows users to visually address Clearance conflicts within a dedicated Conflicts map and subsequently generate a PDF of the Exhibit A map using a GUI-based ArcGIS Pro Toolbox. The Exhibit A tool within the project requires the user to input the Proponent ID along with one of the following: a Cutting Permit (with or without amendments), a Road Permit (with or without amendments), Road Section IDs, or a Special Use Permit. Once the input is provided, the tool automatically executes all the necessary steps to create an Exhibit A map.

# Features of the Exhibit A Tool:

# 1. Creates a new directory to hold the outputs
# 2. Creates the Pending Application/Tenure Road Application and associated Points of Commencement and Points of Termination
# 3 Creates the Conflict Magnitude for Cutting Permits
# 4. Creates a bounding box around the application and then automatically determines the correct layout size and format i.e. Legal Portrait for applications with blocks close together vs ANSI E Landscape for applications that have blocks far apart from one another
# 5. Automatically selects one of 9 possible layouts that best displays the Pending Application
# 6. The layout mapframe automatically zooms to the pending application, pans out an additional 10% for visual appeal and then rounds the scale to the nearest 10,000
# 7. Creates centroids for each cutblock or road section, then calculates the correct mapsheet number for each one and writes that into the layout surround
# 8. Finds and extracts other necessary labelling data from the pending application features, including District, Cut Block, planned gross area, Forest File ID etc and labels those on the layout surround
# 9. Adds tables to the Exhibit A map displaying the POC POT Point Type, Point ID, Planned Gross Area/Length (depending on permit type) and UTM Coordinates
# 10. Adds the ESF ID to the layout
# 11. Labels the Title with Amendment on the layout
# 11. Cleans any transitory layers from the map contents pane

# In summary, once the user has input the required parameters, the tool will fully create an Exhibit A map ready for export in under a minute. This will help eliminate the need to rely on SNC to create Exhibit A maps for the user and help centralize some of the steps in the Clearance process in a single ArcGIS Pro Document.



# - INPUTS:
#   - Proponent ID
#   - Cutting Permit ID or Road Permit Id with Optional Amendment and Section Ids
#   - ESF Submission ID


# - OUTPUTS
# - All outputs are performed in Memory and will be lost when the project is closed.
# - Fully Drawn Exhibit A Layout that can be exported using the Quick Export Toolbox

# --------------------------------------------------------------------------------
# \\spatialfiles2.bcgov\Work\FOR\RNI\DMK\Templates_Utilities\ExA_Templates\ExA_Templates\Exhibit A Mapping Tools.docx - Shortcut.lnk
#
# --------------------------------------------------------------------------------



 
#TODO

    # Add error handling if def query returns empty, script should exit and display a message
    # Currently the script will replace the data source of the old permit that was run by the user with the new permit. Using this method, if you run the same permit twice, replacing
        # the old permit with a copy of the same one will delete the layer, so handling has been built it to stop the script and provide a warning. The script could be changed to 
        # apply symbology from layer but according to forums there is a bug with this. 
    # When launched to different districts, the user should apply an extent limitation on the map so it isnt processing data from the entire province
    # Add handling for cases where there is a single Cut Block, the zoom factor should be around 4.0(?) where as, 
    # if there are multiple cut blocks, the zoom factor should be around 0.1
    # Need to tweak the ESF Submission IDs and the layouts,ll
    
# FEATURES    
    # Use the Table Element X,Y function to automatacally place the table in a given area each time
    # When you run the script but are unhappy with the map size, feature should be that you
        # Can manually choose a layout from a drop down menu and then regenerate the map without bumping out 
        # The data sources.
    




import arcpy
import os
import datetime

# Define Commonly Used Functions
def hide_layer(layer_name):
    '''
    Simple function that assigns the layer name to a layer object for easier handling
    Put in the layer name as a string and it will hide the layer in the map
    '''
    arcpy.AddMessage(f"Hiding Layer: {layer_name}")
    layer = Ex_A_map.listLayers(layer_name)[0]
    layer.visible = False

def reveal_layer(layer_name):
    '''
    Simple function that assigns the layer name to a layer object for easier handling
    Put in the layer name as a string and it will reveal the layer in the map
    '''
    arcpy.AddMessage(f"Revealing Layer: {layer_name}")
    layer = Ex_A_map.listLayers(layer_name)[0]
    layer.visible = True



class Toolbox(object):
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [Tool]


class Tool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Fully Automatic Exhibit A Tool"
        self.description = ""
        self.canRunInBackground = False

    def getParameterInfo(self):
            """Define parameter definitions"""


            proponent_id = arcpy.Parameter(
                displayName="Proponent ID",
                name="Proponent ID",
                datatype="String",
                parameterType="Required",
                direction="Input"
                )


            proponent_name = arcpy.Parameter(
                displayName="Proponent Name",
                name="Proponent Name",
                datatype="String",
                parameterType="Optional",
                direction="Input"
                )


            cp_ID = arcpy.Parameter(
                displayName="Cuttng Permit/TSL",
                name="Cutting Permit",
                datatype="String",
                parameterType="Optional",
                direction="Input"
                )
            
            esf_id = arcpy.Parameter(
                displayName="ESF ID",
                name="ESF ID",
                datatype="String",
                parameterType="Optional",
                direction="Input"
                )
            
            # cp_amendment = arcpy.Parameter(
            #     displayName="Cutting Permit Amendment/ TSL Amendment (Optional)",
            #     name="cp_amendment",
            #     datatype="String",
            #     parameterType="Optional",
            #     direction="Input"
            #     )
            
            rp_ID = arcpy.Parameter(
                displayName="Road Permit i.e. R11085 (Optional)",
                name="rp_ID",
                datatype="String",
                parameterType="Optional",
                direction="Input"
                )
            
            rp_amendment = arcpy.Parameter(
                displayName="Road Permit Amendment (Entered as  digits ie. 18) (Optional)",
                name="rp_amendment",
                datatype="String",
                parameterType="Optional",
                direction="Input"
                )
            
            rp_sections = arcpy.Parameter(
                displayName="Road Section Ids (Entered as comma separated numbers i.e. 2,6,9 *No Leading Zeros)",
                name="rp_sections",
                datatype="String",
                parameterType="Optional",
                direction="Input"
                )
            
            # sup_ID = arcpy.Parameter(
            #     displayName="Special Use Permit (Enter the Forest File ID ie. S27214 (Optional)",
            #     name="sup_ID",
            #     datatype="String",
            #     parameterType="Optional",
            #     direction="Input"
            #     )
            

            # NOTE implement this later. A click box in the gui that allows the user to override the automatic selection of layouts and choose a layout manually. 
            # ex_a_map_size = arcpy.Parameter(
            #     displayName="Exhibit A Map Size (Optional)",
            #     name="ex_a_map_size",
            #     datatype="String",
            #     parameterType="Optional",
            #     direction="Input"
            #     ) 
            
                

            parameters = [proponent_id, proponent_name, cp_ID, esf_id,  rp_ID, rp_amendment, rp_sections] #cp_amendment, ex_a_map_size
            
            return parameters


    def isLicensed(self):
        """Set whether tool is licensed to execute."""
        return True


    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter.  This method is called after internal validation."""
        return

    def execute(self, parameters, messages): 

        arcpy.AddMessage("***Running Auto Exhibit A Tool - ALL OUTPUTS WILL BE SAVED TO T:\ DRIVE AND WILL BE DELETED UPON LOGOUT***")
        arcpy.AddMessage("Step 1 - Initializing User Inputs...")

        # *** ENVIRONMENTS ***
        # To allow overwriting outputs change overwriteOutput option to True.
        arcpy.env.overwriteOutput = True
        
        # *** PARAMETERS ***
        
        def user_inputs():
            '''
            This function assigns the user inputs to variables
            '''
            global proponent_id, proponent_name, cp_ID, esf_id, rp_ID, rp_amendment, rp_sections, sup_ID
            
            proponent_id = parameters[0].valueAsText
            proponent_name = parameters[1].valueAsText
            cp_ID = parameters[2].valueAsText
            esf_id = parameters[3].valueAsText
            rp_ID = parameters[4].valueAsText
            rp_amendment = parameters[5].valueAsText
            rp_sections = parameters[6].valueAsText
            sup_ID = parameters[7].valueAsText
            
            arcpy.AddMessage(f"Proponent Id is: {proponent_id}, cp_ID is: {cp_ID}, esf_id is: {esf_id}") 
            
            
        def define_project_variables(proponent_name):
            
            arcpy.AddMessage("Step 1 - Defining Project Variables...")
            
            
            global aprx, permit_root, Ex_A_map, ften_cutblock_pending, ften_road_sections, sup_pending, mapsheet_20k, cp_pocpot_str, rp_pocpot_str, ex_a_map_frame, query
            # Unable to get the script to write to memory in arcgis pro. Using the T drive of the user because the T drive erases after you log out
            root_dir = os.path.join("T:\\", "Exhibit A Tool Outputs")

            # General Project Variables
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            
            # Create the Permit root directory 
            permit_root = os.path.join(root_dir, proponent_name)

    
            # *** Creating the Layer/Map Objects***
            
            # Assign the Ex A map to a variable

            Ex_A_map_str = "Exhibit A Map"
            Ex_A_map = aprx.listMaps(Ex_A_map_str)[0] 
            
            # Assign the FTEN Cut Block SVW (All) and FTEN Cut Block SVW (Pending) layers to variables as objects
            ften_cutblock_pending = Ex_A_map.listLayers("FTEN Cut Block SVW (Pending)")[0]

            # Assign the FTEN Road Sections SVW Layer to a variable as object
            ften_road_sections = Ex_A_map.listLayers("FTEN Road Sections SVW (DefQuery)")[0]
            
            # Assign the Special Use Permit Layer to a variable as object
            sup_pending = Ex_A_map.listLayers("Special Use Permit - Pending")[0]
            
            # Assign the 20k BCGS Grid to a variable as object
            # global mapsheet_20k
            mapsheet_20k = Ex_A_map.listLayers("20k BCGS Grid")[0]
        
        
            # Assign the Road Permit and Cutting Permit Layer Names to Variables
            cp_pocpot_str = "Cutting Permit P of C"
            rp_pocpot_str = "Road Permit P of C and P of T"
            
            # Assign the Map Frame to a variable
            ex_a_map_frame = "Main Map Frame"
            
            # Set query to None to avoid issues
            query = None
        
        
           


            # *** FUNCTIONS / CLASSES ***


        
        ###############################################################################################################################
        #
        # Setting the Variables for the CP, RP, SUP, BCTS TSL, FSR, and BCTS RP Permits
        #
        ###############################################################################################################################
        
        '''
        There are several types of permits that the script needs to run on. The logic is that regardless of which type of permit, or which
        Variables() function is called, the output of the function will be the variable permit_str. This allows the rest of the script to remain
        virtually unchanged. Another output is permit_dir. Since the code has been adapted to 
        be multi-district and the creation of data has been moved to in memory, permit_dir may be obsolete.
        '''
        def set_cpVariables():
            '''
            This function sets the variables for the Cutting Permit, 
            '''
            arcpy.AddMessage("Step 2 - Setting CP Variables....")  
            

            permit_type = "CP"
            
            permit_str = f"{proponent_id}_CP_{cp_ID}"
            
            # Create the Permit directory which will look like: 2024\Licensee_NRFL\A15384\Canfor\CP\A15384_CP_H47 this will be the main workspace
            permit_dir = os.path.join(permit_root, "CP", permit_str)

            
            def create_cp_def_query():
                  
            
                # Create a select by attribute query to filter FTEN Harvest Authority (cut_permit_layer) by proponent_id and cp_ID
                query = "CUT_BLOCK_FOREST_FILE_ID = '" + proponent_id + "' and HARVEST_AUTH_CUTTING_PERMIT_ID = '" + cp_ID + "'"

                # Create a definition query on the cut_permit_layer
                ften_cutblock_pending.definitionQuery = query
                arcpy.AddMessage(f"\tDefinition Query Applied to FTEN Cutblocks Pending: {query}")


                # Error Handling -  Use a SearchCursor to count the number of records returned by the definition query
                count = 0

                with arcpy.da.SearchCursor(ften_cutblock_pending, "*") as cursor:
                    for row in cursor:
                        count += 1

                # Check if the count is 0 and display a message if the definition query was successful or not.
                if count == 0:
                    arcpy.AddMessage(f"\tNo records returned by definition query: {query}, please check the file number and try again")
                else:
                    arcpy.AddMessage(f"\tDefinition query {query} set for layer: {ften_cutblock_pending}. Records returned: {count}")

           # Return the query to be used elsewhere in the script if needed     
            query = create_cp_def_query()

            return query, permit_str, permit_dir, permit_type
        

             
        # Set the variables for the Road Permit
        def set_rpVariables():
            '''
            This function sets the variables for the Road Permit
            '''
            arcpy.AddMessage("Step 2 - Setting RP Variables...")

            permit_type = "RP"
            if rp_amendment:
                permit_str = f"{proponent_id}_{rp_ID}_Am{rp_amendment}"
                arcpy.AddMessage(f"\tThis is an amendment. Permits String is: {permit_str}")
            else:
                permit_str = f"{proponent_id}_{rp_ID}"
                arcpy.AddMessage(f"\tRoad Permit ID: {permit_str} (No Amendment)")

            permit_dir = os.path.join(permit_root, "RP", permit_str)

            # Start with the base query
            query = f"FOREST_FILE_ID = '{rp_ID}'"

            # Handling rp_sections input by the user
            if rp_sections:
                # Split the input string by commas, strip whitespace, and ensure proper SQL string format
                section_ids = [id.strip() for id in rp_sections.split(',')]
                sections_query = "ROAD_SECTION_ID IN ('" + "', '".join(section_ids) + "')"
                query += f" AND {sections_query}"

            arcpy.AddMessage(f"\tRoad Query: {query}")

            # Apply the query to the definition of ften_road_sections layer
            try:
                ften_road_sections.definitionQuery = query
                arcpy.AddMessage(f"\tDefinition Query Applied to FTEN Road Sections: {query}")
            except Exception as e:
                arcpy.AddMessage(f"Failed to apply definition query with error: {e}")

            # Error Handling - Use a SearchCursor to count the number of records returned by the definition query
            count = 0
            with arcpy.da.SearchCursor(ften_road_sections, "*") as cursor:
                for row in cursor:
                    count += 1

            # Check the records, if the number of records is 0, display a message, if not, display a message with the number of records
            if count == 0:
                arcpy.AddMessage(f"\tNo records returned by definition query: {query}, please check the file number and try again")
            else:
                arcpy.AddMessage(f"\tDefinition query {query} set for layer: {ften_road_sections}. Records returned: {count}")
    
            return query, permit_str, permit_dir, permit_type
            

        
        # Set the variables for special use permits
        def set_supVariables():

            # # Error handling for proponent name. 
            # if not proponent_name:
            #     arcpy.AddError("You need to enter a proponent name before continuing. Exiting script.")
            #     return  # Exit the function if proponent_name is not provided
            
            arcpy.AddMessage("2. Setting SUP Variables...")
            

            permit_type = "SUP"
            
            
            permit_str = f"{sup_ID}"
            arcpy.AddMessage(f"\tThis is a Special Use Permit. Permits String is: {permit_str}")
            
            # Create the Permit directory which will look like: 2024\Licensee_NRFL\A15384\Canfor\CP\A15384_CP_H47 this will be the main workspace
            permit_dir = os.path.join(permit_root, "SUP", permit_str)

            
            def create_sup_def_query():
              
                # Create a select by attribute query to filter FTEN Harvest Authority (cut_permit_layer) by proponent_id and cp_ID
                query = "LIFE_CYCLE_STATUS_CODE = '" + "PENDING" + "' and FOREST_FILE_ID = '" + sup_ID + "'"

                
                # Create a definition query on the cut_permit_layer
                sup_pending.definitionQuery = query
                arcpy.AddMessage(f"\tDefinition Query Applied to Special Use Permit Pending: {query}")
                arcpy.AddMessage("\tDefinition Query Applied")


                # Error Handling -  Use a SearchCursor to count the number of records returned by the definition query
                count = 0

                with arcpy.da.SearchCursor(sup_pending, "*") as cursor:
                    for row in cursor:
                        count += 1

                # Check if the count is 0 and display a message if the definition query was successful or not.
                if count == 0:
                    arcpy.AddMessage(f"\tNo records returned by definition query: {query}, please check the file number and try again")
                else:
                    arcpy.AddMessage(f"\tDefinition query {query} set for layer: Special Use Permit Pending. Records returned: {count}")

                
            query = create_sup_def_query()

            return query, permit_str, permit_dir, permit_type
        
  
               
        ###################################################################################################################################################
        #
        # The following 30 lines of code are for future implementation of BCTS TSL, FSR, and BCTS RP Permits.
        #
        ###################################################################################################################################################
        
        def set_bcts_tslVariables():
            arcpy.AddMessage("2. Setting BCTS TSL Variables- Not working yet")
            pass
     
        
        def set_fsrVariables():
            arcpy.AddMessage("2. Setting FSR Variables- Not working yet")
            pass
      
        
        def set_bcts_rpVariables():
            arcpy.AddMessage("2. Setting BCTS RP Variables- Not working yet")
            pass
        
                
        def select_permit_variables():
            '''
            This function will select the correct permit variables based on the proponent name
            '''
            if proponent_name == "BCTS":
                if cp_ID:
                    set_bcts_tslVariables()
                if rp_ID:
                    if rp_ID[0] == "R":
                        set_bcts_rpVariables()
                    else:
                        set_fsrVariables()
            else:
                if cp_ID:
                    set_cpVariables()
                elif rp_ID:
                    set_rpVariables()
                elif sup_ID:
                    set_supVariables()
        
        ####################################################################################################################################################
        

        def check_existing_data_source(layer_name):  
            '''
            This function checks to make sure the GIS user is not running the script twice for the same permit. If they are
            this will cause an error and delete the Permit Application from the Contents Pane
            '''
            arcpy.AddMessage("Step 3 - Checking for existing data source conflicts...")    

            # Try to get the layer by name
            layer = Ex_A_map.listLayers(layer_name)[0] if Ex_A_map.listLayers(layer_name) else None

            # If the layer is found, check the data source path for the permit string
            if layer:
                data_source = layer.dataSource
                
                # Check if the permit string is in the data source path
                if permit_str in data_source:
                    arcpy.AddError(f"Whoops! This is the same permit as the last one you ran, please change the data source of {layer_name} to a different source and run the script with the same permit again.")
                    raise Exception(f"Whoops! You have already run this Permit Request, please change the data source of {layer_name} to a different source and run the script again.")
                else:
                    arcpy.AddMessage(f"\tNo conflict with the data source. Proceeding with the script.")
            else:
                arcpy.AddError(f"Layer named '{layer_name}' not found in the current project.")
                raise Exception(f"Layer named '{layer_name}' not found in the current project.")

            # If no error is raised, continue with further processing
        


        ###############################################################################################################################
        #
        # Step 1 - Folder Creation
        # Use if-then logic to verify that the intended directory has not been created already
        # This is a holdover script from DMK because they saved their files to disk. It can be moved evenutally
        #
        ###############################################################################################################################


        def make_dir():  
            '''
            Folder Creation
            Use if-then logic to verify that the intended directory has not been created already
            If the directory exists create a new folder inside the existing folder and append the month, date and the year 
            to the folder name.
            '''
            arcpy.AddMessage(f"Step 4 - Creating folder for {permit_str}")

            arcpy.AddMessage(f"\tChecking for existing folder.")
            if not os.path.exists(permit_dir):
                # If it doesn't exist, create it
                os.makedirs(permit_dir)
                arcpy.AddMessage(f"\tThe folder does not exist. Creating: {permit_dir}")
                
            else:
                # If it exists, use the existing directory
                arcpy.AddMessage(f"\tUsing existing directory: {permit_dir}")
                

               
                    
        
        def create_pending_fc():
            """
            Creates a pending feature class based on the provided arguments and returns the output path.

            This function creates a feature class from the pending feature class based on the input arguments
            It checks if the feature class has any records and raises an error if it is empty.

            Returns:
                str: The output path of the created pending feature class.
            """
            try:
                # Create the output path for the output feature class
                pending_tenure_output_fc = os.path.join(permit_dir, f"{permit_str}_pending_tenure")
                
                # Determine which pending feature class to use based on the input identifiers
                if cp_ID:
                    pending_fc = ften_cutblock_pending
                elif rp_ID:
                    pending_fc = ften_road_sections
                elif sup_ID:
                    pending_fc = sup_pending
                else:
                    raise ValueError("The permit you entered is no valid. Please check your Proponent Id, CP ID, or RP ID.")
                
                # Copy the features to the new feature class
                arcpy.management.CopyFeatures(pending_fc, pending_tenure_output_fc)
                
                # Check to see if the feature class has any records
                count = int(arcpy.management.GetCount(pending_tenure_output_fc)[0])
                if count == 0:
                    arcpy.AddMessage(f"\tFeature class {pending_tenure_output_fc} has no records. Exiting script.")
                    arcpy.AddError("Script Ended - No records found while executing the definition query. Check your Proponent Id, CP ID, or RP ID and try again.")
                    raise ValueError("Script Ended - No records found while executing the definition query. Check your Proponent Id, CP ID, or RP ID and try again.")
                else:
                    arcpy.AddMessage(f"\tPending feature class has {count} records.")
                
                # Return the output path of the created pending feature class
                return pending_tenure_output_fc

            except Exception as e:
                arcpy.AddError(f"An error in creating a pending feature class occurred: {e}")
                raise



            
        ################################################################################################################################
        #
        # Select Layout
        #
        #############################################################################################################################
        
        def evaluate_page_size(input_pending_fc, target_map, lyt_type):
            
            arcpy.AddMessage("Evaluating Page Size...")
            '''
            # Bounding output name could be exA_bounding_box.shp or fn_bounding_box.shp as a string
            # Input pending fc is either FTEN PENDING or FT HARVEST AUTHORITY (FOR FN MAP)
            # Working Dir is permit dir for ex a or fn_dir for fn map
            # lyt_type = [ExA] or [FN]
            # target_map = Ex_A_map or FN_map
            '''
            global chosen_layout_str

            
            evaluate_temp_fc = "memory\\evaluate_temp_fc"
            if arcpy.Exists(evaluate_temp_fc):
                arcpy.AddMessage(f"\tDeleting existing temporary feature class {evaluate_temp_fc}")
                arcpy.Delete_management(evaluate_temp_fc)
            arcpy.management.MakeFeatureLayer(input_pending_fc, evaluate_temp_fc)
            arcpy.AddMessage(f"\tTemporary feature class created: {evaluate_temp_fc}")

            # Generate a minimum bounding rectangle by area for the temporary layer
            bounding_box_fc = "memory\\bounding_box"
            arcpy.AddMessage(f"\tCreating bounding box feature class: {bounding_box_fc}")
            arcpy.management.MinimumBoundingGeometry(evaluate_temp_fc,
                                                bounding_box_fc,
                                                geometry_type="RECTANGLE_BY_AREA",
                                                group_option="ALL",
                                                mbg_fields_option="MBG_FIELDS")
            arcpy.AddMessage(f"\tBounding box feature class created: {bounding_box_fc}")

            # Make a feature layer, and then use the .getOutput() method on the results object to get the actual layer object. 
           
            bb_mem_fc = arcpy.management.MakeFeatureLayer(bounding_box_fc, "bounding_box_mem_lyr")
            bb_mem_lyr_obj = bb_mem_fc.getOutput(0) #output is now a layer object
            
            # Add the bounding box layer to the map
            target_map.addLayer(bb_mem_lyr_obj)

            
            # Check if the bounding box feature class exists
            if not arcpy.Exists(bb_mem_lyr_obj):
                arcpy.AddMessage(f"\tBounding box feature class {bb_mem_lyr_obj} does not exist.")
                arcpy.AddError(f"Bounding box feature class {bb_mem_lyr_obj} does not exist.")
                raise ValueError(f"Bounding box feature class {bb_mem_lyr_obj} does not exist.")

            # Check to see if the feature class has any records
            count = arcpy.management.GetCount(bb_mem_lyr_obj)
            if count == 0:
                arcpy.AddMessage(f"\tBounding box feature class {bb_mem_lyr_obj} has no records.")
                arcpy.AddError(f"Bounding box feature class {bb_mem_lyr_obj} has no records.")
                raise ValueError(f"Bounding box feature class {bb_mem_lyr_obj} has no records.")
            
            # Initialize variables to store geometry and area
            polygon = None
            area = None
            
            # Use a search cursor to retrieve the geometry and area of the feature class
            with arcpy.da.SearchCursor(bb_mem_lyr_obj, ['SHAPE@', 'SHAPE@AREA']) as cursor:
                try:
                    polygon, area = next(cursor)
                except StopIteration:
                    raise ValueError(f"No valid geometry or area found in {bb_mem_lyr_obj}.")
            
            # Proceed only if polygon and area have been successfully retrieved
            if polygon is not None and area is not None:
                extent = polygon.extent
                width, height = extent.width, extent.height
        

            # Determine page orientation based on width and height
            portrait_orientation = width < height  
            
            # Initialize orientation and page variables
            orientation = "PORTRAIT" if portrait_orientation else "LANDSCAPE"
            i = 0  # Default value in case none of the conditions match
            page = ""  # Default value
            
            # Evaluate the page size needed based on width and area
            if portrait_orientation == True:
                orientation = 'PORTRAIT'
                if area <= 23139401.1175:
                    i = 1
                    page = "legal"
                elif area > 23139401.1175 and area <= 75344603.4531:
                    i = 3
                    page = "AnsiC"
                else:
                    i = 4
                    page = "AnsiD"
            

            elif portrait_orientation == False:
                orientation = 'LANDSCAPE'
                if area <= 22131657.1665:
                    i = 2
                    page = "legal"
                if area > 22131657.1665 and area <= 35234909.1795:
                    i = 9
                    page = "Tabloid"
                if area > 35234909.1795 and area <= 153846020.67:
                    i = 5
                    page = "AnsiD"
                if area > 153846020.67:
                    i = 2
                    page = "legal"

            
            # Build the layout name based on the page size and orientation. The  layout name should have the format eg. 1_PORTRAIT_ExA_legal
            chosen_layout_str = f"{i}_{orientation}_{lyt_type}_{page}"
            arcpy.AddMessage(f"\tChosen Layout: {chosen_layout_str}") 
        
            # Assign the chosen layout str to a layout object
            global chosen_layout_obj
            chosen_layout_obj = aprx.listLayouts(chosen_layout_str)[0]
            
            # Open the chosen layout
            chosen_layout_obj.openView()   
            arcpy.AddMessage(f"\tOpening layout: {chosen_layout_str}")
            
  
            return chosen_layout_str, chosen_layout_obj
        
        
        ################################################################################################################################
        #
        # Update the layer sources PENDING APPLICATION
        #
        #############################################################################################################################

        def update_layer_sources(layer_to_update):
            '''
            Update the data sources for the pending application and poc_pot layers
            '''

            arcpy.AddMessage("Updating Layer Sources...")

            # Set the layer to be updated
            pending_application_layer = Ex_A_map.listLayers(layer_to_update)[0]
            orig_lyr = pending_application_layer

            # Check to make sure the layer exists
            if not orig_lyr:
                raise ValueError("\tLayer not found")
            else:
                arcpy.AddMessage(f"\tLayer found: {orig_lyr.name}")



            new_conn_dict = {
                'connection_info': {'database': permit_dir}, # This is the path to the folder where the new data source is located
                'dataset': f'{permit_str}_pending_tenure.shp', # This has to be the file name including the .shp
                'workspace_factory': 'Shape File' # This is the type of file you are connecting to (needs the space)
            }
            arcpy.AddMessage(f"\tNew connection properties created: {new_conn_dict}")

            # Update the connection properties
            new_conn_prop = orig_lyr.updateConnectionProperties(orig_lyr.connectionProperties, new_conn_dict)
            arcpy.AddMessage(f"\tData source updated for layer new connection properties are now: {new_conn_prop}")
        

        ################################################################################################################################
        #
        # Update the layer sources POC POT
        #
        #############################################################################################################################


        def update_pocpot_layer_connection(layer_to_update, cp_or_rp_pocpot):
            
            arcpy.AddMessage(f"Updating Layer Sources for {layer_to_update}...")
            orig_lyr = Ex_A_map.listLayers(layer_to_update)[0]


            # Check to make sure the layer exists
            if not orig_lyr:
                raise ValueError("Layer not found")
            else:
                arcpy.AddMessage(f"\tLayer found: {orig_lyr.name}")

            # Set a variable that represents layer connection Properties
            # origConnPropDict = orig_lyr.connectionProperties
            arcpy.AddMessage(f"\tOld {orig_lyr} connection properties retrieved")



            new_conn_dict = {
                'connection_info': {'database': permit_dir}, # This is the path to the folder where the new data source is located
                'dataset': cp_or_rp_pocpot, #f'{permit_str}_u_POCPOT.shp', # This has to be the file name including the .shp
                'workspace_factory': 'Shape File' # This is the type of file you are connecting to (needs the space)
            }
            arcpy.AddMessage(f"\tNew connection properties created: {new_conn_dict}")

            # Update the connection properties
            new_conn_prop = orig_lyr.updateConnectionProperties(orig_lyr.connectionProperties, new_conn_dict)
            arcpy.AddMessage(f"\tData source updated for {layer_to_update}. New connection properties are now: {new_conn_prop}")
        
        
        ################################################################################################################################
        #
        # Create Point of Commencement
        #
        #############################################################################################################################
        def create_cp_poc_pot():
            # Create the output path for the Point of Commencement output feature class
            cp_pocpot = f"{permit_str}_CP_POCPOT"
            u_pocPot_fc = os.path.join(permit_dir, cp_pocpot)

            # Convert vertices to points
            with arcpy.EnvManager(outputCoordinateSystem="PROJCS[\"NAD_1983_UTM_Zone_10N\",GEOGCS[\"GCS_North_American_1983\",DATUM[\"D_North_American_1983\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",500000.0],PARAMETER[\"False_Northing\",0.0],PARAMETER[\"Central_Meridian\",-123.0],PARAMETER[\"Scale_Factor\",0.9996],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]"):             
                arcpy.management.FeatureVerticesToPoints(pending_tenure_output_fc, u_pocPot_fc, point_location="START")
            arcpy.AddMessage("\tFeature Vertices To Points Completed")
            

            
            # Add Fields to the new Point of Commencement feature class
            arcpy.management.AddFields(u_pocPot_fc, field_description=[["PT_TYPE", "TEXT", "POINT TYPE", "10", "", ""], ["PT_ID", "TEXT", "POINT ID", "10", "", ""], ["EASTING", "LONG", "", "", "", ""], ["NORTHING", "LONG", "", "", "", ""]])
            arcpy.AddMessage("\tAdd Fields Completed")

            arcpy.management.CalculateFields(u_pocPot_fc, expression_type="PYTHON3", fields=[["PT_TYPE", "\"POC\"", ""], ["PT_ID", "\"POC\" + str(!FID! + 1)", ""]])
            arcpy.AddMessage("\tCalculate Fields Completed") 


            # Calculate Easting and Northing Fields
            arcpy.management.CalculateGeometryAttributes(u_pocPot_fc, geometry_property=[["EASTING", "POINT_X"], ["NORTHING", "POINT_Y"]], coordinate_system="PROJCS[\"NAD_1983_UTM_Zone_10N\",GEOGCS[\"GCS_North_American_1983\",DATUM[\"D_North_American_1983\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",500000.0],PARAMETER[\"False_Northing\",0.0],PARAMETER[\"Central_Meridian\",-123.0],PARAMETER[\"Scale_Factor\",0.9996],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]")[0]
            arcpy.AddMessage("\tCalculate Geometry Attributes Completed")
            arcpy.AddMessage(f"\tPOCT POT Fields Added, PT ID Calculated, Calculate Geometry Attributes Completed")

            return cp_pocpot


        def create_rp_poc_pot():

            arcpy.AddMessage("Creating POC POT for Road Permit...") 
            # Create the output path for the Point of Commencement output feature class
            u_poc_fc = os.path.join(permit_dir, f"{permit_str}_RP_POC")

            # Convert vertices to point of commencement
            with arcpy.EnvManager(outputCoordinateSystem="PROJCS[\"NAD_1983_UTM_Zone_10N\",GEOGCS[\"GCS_North_American_1983\",DATUM[\"D_North_American_1983\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",500000.0],PARAMETER[\"False_Northing\",0.0],PARAMETER[\"Central_Meridian\",-123.0],PARAMETER[\"Scale_Factor\",0.9996],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]"):             
                arcpy.management.FeatureVerticesToPoints(pending_tenure_output_fc, u_poc_fc, point_location="START")
            arcpy.AddMessage("\tFeature Vertices To Points (START) Completed")
            
                        
            # Add Fields to the new Point of Commencement feature class
            arcpy.management.AddFields(u_poc_fc, field_description=[["PT_TYPE", "TEXT", "PT TYPE", "10", "", ""], ["PT_ID", "TEXT", "PT ID", "10", "", ""],  ["PT_STRING", "TEXT", "PT STRING", "50", "", ""], ["EASTING", "LONG", "EASTING", "", "", ""], ["NORTHING", "LONG", "NORTHING", "", "", ""]])
            arcpy.AddMessage("\tAdd Fields Completed")

            arcpy.management.CalculateFields(u_poc_fc, expression_type="PYTHON3", fields=[["PT_TYPE", "\"POC\"", ""], ["PT_ID", "\"POC\" + str(!FID! + 1)", ""]])[0]
            arcpy.AddMessage("\tCalculate Fields Completed") 


            # Calculate Easting and Northing Fields
            arcpy.management.CalculateGeometryAttributes(u_poc_fc, geometry_property=[["EASTING", "POINT_X"], ["NORTHING", "POINT_Y"]], coordinate_system="PROJCS[\"NAD_1983_UTM_Zone_10N\",GEOGCS[\"GCS_North_American_1983\",DATUM[\"D_North_American_1983\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",500000.0],PARAMETER[\"False_Northing\",0.0],PARAMETER[\"Central_Meridian\",-123.0],PARAMETER[\"Scale_Factor\",0.9996],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]")[0]
            arcpy.AddMessage("\tCalculate Geometry Attributes Completed")
            
            
            arcpy.management.CalculateField(u_poc_fc, field="PT_STRING", expression="!PT_ID! + \" UTM10 \" + str(!EASTING!) + \", \" + str(!NORTHING!)")
            arcpy.AddMessage(f"\tPOCT POT Fields Added, PT ID Calculated, Calculate Geometry Attributes Completed")
            
            ################################################################################################################################
            
            # Convert vertices to point of termination
            
            u_pot_fc = os.path.join(permit_dir, f"{permit_str}_u_POT")
            with arcpy.EnvManager(outputCoordinateSystem="PROJCS[\"NAD_1983_UTM_Zone_10N\",GEOGCS[\"GCS_North_American_1983\",DATUM[\"D_North_American_1983\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",500000.0],PARAMETER[\"False_Northing\",0.0],PARAMETER[\"Central_Meridian\",-123.0],PARAMETER[\"Scale_Factor\",0.9996],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]"):             
                arcpy.management.FeatureVerticesToPoints(pending_tenure_output_fc, u_pot_fc, point_location="END")
            arcpy.AddMessage("\tFeature Vertices To Points END Completed")
            

            
            # Add Fields to the new Point of Commencement feature class
            arcpy.management.AddFields(u_pot_fc, field_description=[["PT_TYPE", "TEXT", "", "10", "", ""], ["PT_ID", "TEXT", "", "10", "", ""], ["PT_STRING", "TEXT", "", "50", "", ""], ["EASTING", "LONG", "", "", "", ""], ["NORTHING", "LONG", "", "", "", ""]])[0]
            arcpy.AddMessage("\tPOT Add Fields Completed")

            arcpy.management.CalculateFields(u_pot_fc, expression_type="PYTHON3", fields=[["PT_TYPE", "\"POT\"", ""], ["PT_ID", "\"POT\" + str(!FID! + 1)", ""]])
            arcpy.AddMessage("\tPOT Calculate Fields Completed") 


            # Calculate Easting and Northing Fields
            arcpy.management.CalculateGeometryAttributes(u_pot_fc, geometry_property=[["EASTING", "POINT_X"], ["NORTHING", "POINT_Y"]], coordinate_system="PROJCS[\"NAD_1983_UTM_Zone_10N\",GEOGCS[\"GCS_North_American_1983\",DATUM[\"D_North_American_1983\",SPHEROID[\"GRS_1980\",6378137.0,298.257222101]],PRIMEM[\"Greenwich\",0.0],UNIT[\"Degree\",0.0174532925199433]],PROJECTION[\"Transverse_Mercator\"],PARAMETER[\"False_Easting\",500000.0],PARAMETER[\"False_Northing\",0.0],PARAMETER[\"Central_Meridian\",-123.0],PARAMETER[\"Scale_Factor\",0.9996],PARAMETER[\"Latitude_Of_Origin\",0.0],UNIT[\"Meter\",1.0]]")[0]
            arcpy.AddMessage("\tCalculate Geometry Attributes Completed")

            
            arcpy.management.CalculateField(u_pot_fc, field="PT_STRING", expression="!PT_ID! + \" UTM10 \" + str(!EASTING!) + \", \" + str(!NORTHING!)")
        
            arcpy.AddMessage(f"\tPOT Fields Added, PT ID Calculated, Calculate Geometry Attributes Completed")
            
            rp_pocpot = f"{permit_str}_RP_POCPOT"
            
            u_poc_pot_fc = os.path.join(permit_dir, rp_pocpot)
            arcpy.management.Merge([u_poc_fc, u_pot_fc], u_poc_pot_fc)
            return rp_pocpot

        ################################################################################################################################
        #
        # Adjust the layout mapframe
        #
        #############################################################################################################################
        def zoom_to_feature_extent(map_name, map_frame, layer_name, zoom_factor, layout_name):


            ''' This function will focus the layout on the selected feature and then pan out x% (depending on zoom factor) 
            to show the surrounding area. If you use an 0.8 (80%) Zoom Factor, to calculate the zoom percentage: Original zoom 
            is 100% (the initial extent of the splitline layer).The new extent is 160% larger than the original. This is because 
            you are adding 80% of the width to both sides (left and right) and 80% of the height to both top and bottom.'''\
                
            arcpy.AddMessage("Running zoom_to_feature_extent function....")
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
                arcpy.AddMessage(f"\tNo map found with the name: {map_name}")
                
                return
            map_obj = maps[0]
            layers = map_obj.listLayers(layer_name)
            if not layers:
                arcpy.AddError(f"No layer found with the name: {layer_name} in map: {map_name}")
                arcpy.AddMessage(f"\tNo layer found with the name: {layer_name} in map: {map_name}")
                
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
                
                arcpy.AddMessage(f"\tZoomed to {layer_name} with zoom factor {zoom_factor}")
                arcpy.AddMessage("\tZoom to Feature Extent Completed")
                
            except Exception as e:
                arcpy.AddError(f"Error in zooming to extent: {str(e)}")
               
        
        def round_scale(map_frame, chosen_layout_obj):
            '''This function will round the scale of the map frame to the nearest 10,000 or to 5000 if less than 5000'''
            
            arcpy.AddMessage("Running round_scale function....")
            
            mf_list = chosen_layout_obj.listElements("MAPFRAME_ELEMENT", map_frame)
            
            if mf_list:  # Check if the list is not empty
                mf = mf_list[0]  # Access the first map frame in the list
                current_scale = mf.camera.scale  # Access the scale property directly
                arcpy.AddMessage(f"\tCurrent scale is: {current_scale}")

                # Check if scale is less than 5000
                if current_scale < 5000:
                    rounded_scale = 5000
                else:
                    # Round the scale to the nearest 10,000
                    rounded_scale = round(current_scale / 10000) * 10000

                arcpy.AddMessage(f"\tRounded scale is: {rounded_scale}")

                # Set the map frame to the new rounded scale
                mf.camera.scale = rounded_scale  # Set the scale property directly
                arcpy.AddMessage(f'\tZooming the map to the new scale of {rounded_scale}')
            else:
                arcpy.AddMessage("\tNo map frame found with the given name.")


        # Remove the temporay layer from the map
        def remove_temp_layer():
            global bounding_box_lyr_obj
            
            arcpy.AddMessage("Removing Temp Layer from Map...")
            # Find the layer
            bounding_box_lyr_obj = Ex_A_map.listLayers("bounding_box_mem_lyr")[0]
            
            # Check if the layer exists
            if not bounding_box_lyr_obj:
                arcpy.AddError("When trying to remove layer, Layer not found")
                
            # Remove the layer    
            Ex_A_map.removeLayer(bounding_box_lyr_obj)
            arcpy.AddMessage("\tTemp Layer Removed from Map")
        
        
        ################################################################################################################################
        #
        # Locate Mapsheet numbers for labelling
        #
        #############################################################################################################################
        
        
        def create_mapsheet_centroid():
            '''
            This function will calculate feature to point and put a point into each of the pending cut blocks on FTEN Cut Block SVW (Pending)
            It will then intersect the points with the 20k BCGS Grid to find the mapsheet number for each cut block. Next it creates a set()
            of unique mapsheet numbers and makes them available to the rest of the script. The output of this function is the mapsheet_centroid, 
            and ften pending centroid which is saved to the permit directory. Future modification could either make these temporary layers or delete
            them after the script is complete.
            '''
            
            # Set overwrite to True
            arcpy.env.overwriteOutput = True
            arcpy.AddMessage("Calculating Mapsheet Centroid...")
            

            ften_pending_centroid_check = "memory\\ften_pending_centroid" 
            arcpy.AddMessage(f"Creating centroid from {pending_tenure_output_fc}")
            
            ften_pending_centroid = arcpy.ValidateTableName(ften_pending_centroid_check)
            
            arcpy.management.FeatureToPoint(in_features=pending_tenure_output_fc, out_feature_class=ften_pending_centroid, point_location="INSIDE")

            # Check to see if ften_pending_centroid has records
            count = arcpy.management.GetCount(ften_pending_centroid)
            if count == 0:
                arcpy.AddMessage(f"\tFeature class {ften_pending_centroid} has no records.")
                arcpy.AddError(f"Feature class {ften_pending_centroid} has no records.")
                raise ValueError(f"Feature class {ften_pending_centroid} has no records.")
            else:
                arcpy.AddMessage(f"\tFeature class {ften_pending_centroid} has {count} records, intersecting with 20k BCGS Grid.")
            
            # Intersect the ften pending Centroid with the 20k BCGS Grid
            # mapsheet_centroid = os.path.join(permit_dir, "mapsheet_centroid.shp")
            mapsheet_centroid_check = "memory\\mapsheet_centroid"
            
            # Validate the new file name before proceeding. There was an invalid character bug prior to implementing this fix. 
            mapsheet_centroid = arcpy.ValidateTableName(mapsheet_centroid_check)
   
            
            # Check to see if the mapsheet_centroid feature class exist
            # if arcpy.Exists(mapsheet_centroid):
            #     arcpy.AddMessage(f" {mapsheet_centroid} exists.")

        
            arcpy.analysis.Intersect(in_features=[[ften_pending_centroid], [mapsheet_20k]], out_feature_class=mapsheet_centroid)
            arcpy.AddMessage("\tMapsheet Intersect Completed")

            
            
            # Check to see if the mapsheet_centroid feature has records, if it has no records, skip the search cursor and arcpy.AddMessage an error message.
            count = arcpy.management.GetCount(mapsheet_centroid)
            if count == 0:
                arcpy.AddMessage(f"\tFeature class {mapsheet_centroid} has no records.")
                arcpy.AddError(f"Feature class {mapsheet_centroid} has no records. You will have manually enter mapsheet numbers.")
                raise ValueError(f"Feature class {mapsheet_centroid} has no records.")
            
            # Use a search cursor to read each row and create a list of mapsheets. Print only each unique mapsheet number to messages.
            arcpy.AddMessage(f"Feature class {mapsheet_centroid} has {count} records, adding Mapsheet Number.")
            unique_mapsheet_set = set()
            with arcpy.da.SearchCursor(mapsheet_centroid, "MAP_TILE_DISPLAY_NAME") as cursor:
                for row in cursor:
                    unique_mapsheet_set.add(row[0])
                    arcpy.AddMessage(f"\tUnique Mapsheet Numbers: {unique_mapsheet_set}")
            


            return unique_mapsheet_set
             

        
        ################################################################################################################################
        #
        # Labelling the layout
        #
        #############################################################################################################################
        # Read the selected Pending Application Layer and arcpy.AddMessage the attributes. If you need to access individul rows it will need to be changed

        def label_cp_layout():
            arcpy.AddMessage("Labelling Layout with CP Parameters...")
            
            # Fields to retrieve from the FTEN Pending Layer * Might not need all of them.
            pending_tenure_output_fc_fields = ['CUT_BLOCK_', 'HARVEST__1', 'CUT_BLOCK1', 'TIMBER_MAR', 'PLANNED_GR', 'CLIENT_NAM', 'ADMIN_DIST']

            # Fetching cursor data
            total_planned_gross_area = 0
            cursor_data = []
            with arcpy.da.SearchCursor(pending_tenure_output_fc, pending_tenure_output_fc_fields) as cursor:
                for row in cursor:
                    cursor_data.append(row)
                    # Accumulate total planned gross area
                    total_planned_gross_area += row[4]  # Assuming 'PLANNED_GR' is at index 4

            if not cursor_data:
                raise ValueError("No data returned from the cursor. Check the path and fields.")

            # Use the first row to display other details
            row = cursor_data[0]
            cut_block_str, cutting_permit_id_str, cut_block_id_str, timber_mark_str, planned_gross_area_str, client_name_str, admin_dis_str = row
            record_string = f"Cut Block: {cut_block_str}, Cutting Permit ID: {cutting_permit_id_str}, Cut Block ID: {cut_block_id_str}, Timber Mark: {timber_mark_str}, Planned Gross Area: {planned_gross_area_str}, Client Name: {client_name_str}, Admin District: {admin_dis_str}"
            arcpy.AddMessage(record_string)
            arcpy.AddMessage(f"\tYour chosen record has the following details: {record_string}")

            # Update the text elements in the layout
            for lyt in aprx.listLayouts(chosen_layout_str):
                for elm in lyt.listElements():
                    if elm.name == "Map_Title":
                        elm.text = f"MAP OF {cut_block_str} CP {cutting_permit_id_str} (shown in bold black)"
                        arcpy.AddMessage("\tMap of... text element changed")
                    
                    if elm.name == "Area/Length Heading":
                        elm.text = f'Area (Ha)'
                        arcpy.AddMessage("\tArea/Length Heading text element changed to Area")
                    
                    if elm.name == "Area":
                        # Update the text to show the total planned gross area
                        elm.text = f'{total_planned_gross_area}'  # Adjust unit if necessary
                        arcpy.AddMessage("\tTotal Planned Gross Area text element changed")
                        
                    if elm.name == "ESF Submission ID":
                        elm.text = f'{esf_id}'
                        arcpy.AddMessage("\tESF Submission ID text element changed")
                        
                    if elm.name == "Map Sheet":
                        # Convert each item in the set to a string, then join with commas
                        elm.text = ', '.join(unique_mapsheet_set)
                        arcpy.AddMessage("\tMap Sheet text element changed")
                        
                    if elm.name == "FOREST DISTRICT":
                        elm.text = f'{admin_dis_str}'
                        arcpy.AddMessage("\tESF Submission ID text element changed")
                    
                    if elm.name == "RP POCPOT TABLE":
                        elm.visible = False
                        arcpy.AddMessage("\tRP POCPOT TABLE visibility set to off")
                        
                    if elm.name == "CP POCPOT TABLE":
                        elm.visible = True
                        arcpy.AddMessage("\tCP POCPOT TABLE visibility set to on")

       
        def label_rp_layout():
            arcpy.AddMessage("Labelling Layout with RP Parameters...")
            
            # Fields to retrieve from the FTEN Pending Layer * Might not need all of them.
            pending_tenure_output_fc_fields = ['FOREST_FIL', 'SECTION_WI', 'GEOGRAPHIC', 'GEOGRAPH_1', 'FEATURE_LE']

            # Fetching cursor data
            total_planned_length = 0
            cursor_data = []
            with arcpy.da.SearchCursor(pending_tenure_output_fc, pending_tenure_output_fc_fields) as cursor:
                for row in cursor:
                    cursor_data.append(row)
                    # Accumulate total road length
                    total_planned_length += row[4]  

            if not cursor_data:
                raise ValueError("No data returned from the cursor. Check the path and fields.")

            # Use the first row to display other details
            row = cursor_data[0]
            forest_file_str, section_width_str, geographic_str, geographic_1_str, feature_length_str = row
            record_string = f"Forest File ID: {forest_file_str}, Section Width: {section_width_str}, Geograhic Acronym: {geographic_str}, Resource District: {geographic_1_str}, Feature Length: {total_planned_length}"

            arcpy.AddMessage(f"\tYour chosen record has the following details: {record_string}")

            # Update the text elements in the layout
            for lyt in aprx.listLayouts(chosen_layout_str):
                for elm in lyt.listElements():
                    # Update the Map Title based on the amendment status
                    if elm.name == "Map_Title":
                        if rp_amendment:
                            elm.text = f"MAP OF {forest_file_str} Amendment {rp_amendment} (shown in bold black)"
                            arcpy.AddMessage("\tMap of... text element changed to include amendment number")
                        else:
                            elm.text = f"MAP OF {forest_file_str} (shown in bold black)"
                            arcpy.AddMessage("\tMap of... text element changed without amendment number")

                    # Other text elements update
                    if elm.name == "Area/Length Heading":
                        elm.text = f'Total Length (m)'
                        arcpy.AddMessage("\tArea/Length Heading text element changed to Length")
                    
                    if elm.name == "Area":
                        elm.text = f'{total_planned_length} m'
                        arcpy.AddMessage("\tTotal Planned Length text element changed")
                    
                    if elm.name == "ESF Submission ID":
                        elm.text = f'{esf_id}'
                        arcpy.AddMessage("\tESF Submission ID text element changed")
                    
                    if elm.name == "Map Sheet":
                        elm.text = ', '.join(unique_mapsheet_set)
                        arcpy.AddMessage("\tMap Sheet text element changed")
                    
                    if elm.name == "FOREST DISTRICT":
                        elm.text = f'{geographic_str}'
                        arcpy.AddMessage("\tForest District text element changed")
                    
                    if elm.name == "RP POCPOT TABLE":
                        elm.visible = True
                        arcpy.AddMessage("\tRP POCPOT TABLE visibility set to on")
                        
                    if elm.name == "CP POCPOT TABLE":
                        elm.visible = False
                        arcpy.AddMessage("\tCP POCPOT TABLE visibility set to off")

        
        ################################################################################################################################
        #
        # **Main** Call all of the functions
        #
        #############################################################################################################################

        # Assigned the user inputs to variables
        user_inputs()

        # Define the project variables
        define_project_variables(proponent_name) 

        if cp_ID:
            query, permit_str, permit_dir, permit_type = set_cpVariables()  #TODO Remove Query and Permit type
        if rp_ID:
            query, permit_str, permit_dir, permit_type = set_rpVariables()
        if sup_ID:
            query, permit_str, permit_dir, permit_type = set_supVariables()
        
        # Set the permit variables
        select_permit_variables()  

        if cp_ID:
            check_existing_data_source("Cutting Permit P of C")
        elif rp_ID:
            check_existing_data_source("Road Permit P of C and P of T")
        elif sup_ID:
            check_existing_data_source("Special Use Permit - Pending")

        
        # Make the directory to hold the temp files. This will be in the T:\ drive and will be deleted upon logging out of the GTS
        make_dir()    

        # Close all views in the project - This seems to makes the script run faster because it isn't drawing a map at the same time
        arcpy.AddMessage("Step 5 - Closing All Views...")

        aprx.closeViews("MAPS_AND_LAYOUTS")

        # Create the pending tenure output feature class
        pending_tenure_output_fc = create_pending_fc()
        
        # Create a minimum bounding box around the feature class and use that to determine the correct page size and orientation
        chosen_layout_str, chosen_layout_obj = evaluate_page_size(pending_tenure_output_fc, Ex_A_map, "ExA")    


        # Update the data sources. This is done to preserve symbology and labelling. Could be changed later to import symbology from file but I had issues with that and it may be an esri bug.
        # https://support.esri.com/en-us/bug/a-custom-python-script-tool-for-apply-symbology-from-la-bug-000155515
        if cp_ID:
            update_layer_sources("Pending Application")
        elif rp_ID:
            update_layer_sources("Tenure Road Application")   
        elif sup_ID:
            update_layer_sources("Special Use Permit - Pending") # Not tested yet

        # Change the features on the layout depending if it is a CP or RP. This will create the Point of Commencement and Point of Termination feature classes, update the layer of the existing POCPOT, 
        # and then hide/reveal the appropriate layer on the layout. eg, if you have a RP, you dont want to see any CP POCPOTS on the map. 
        if cp_ID:
            arcpy.AddMessage("Setting up for Cutting Permit")
            cp_pocpot = create_cp_poc_pot()
            update_pocpot_layer_connection(cp_pocpot_str, cp_pocpot)
            hide_layer(rp_pocpot_str)
            reveal_layer(cp_pocpot_str)

        if rp_ID:
            arcpy.AddMessage("Setting up for Road Permit")
            rp_pocpot = create_rp_poc_pot()
            update_pocpot_layer_connection(rp_pocpot_str, rp_pocpot)
            hide_layer(cp_pocpot_str)
            reveal_layer(rp_pocpot_str)

        # Zoom the mapframe to the extent of the permit
        zoom_to_feature_extent(Ex_A_map.name, ex_a_map_frame, "bounding_box_mem_lyr", 0.1, chosen_layout_str)

        # Round the scale of the map frame to the nearest 10,000
        round_scale(ex_a_map_frame, chosen_layout_obj)

        # Remove the temporary bounding box layer from the map
        remove_temp_layer()

        # Create centroids on each feature, this will be interescted with the BCGS Mapsheet to find the mapsheet number of each pending application. It then returns the mapsheets in a list to be used in the labelling
        unique_mapsheet_set = create_mapsheet_centroid()

        # Clear any selected features before going into labelling        
        Ex_A_map.clearSelection()

        # Label the layout with the appropriate information for CPs or RPs
        if cp_ID:
            label_cp_layout()
            
        if rp_ID:
            label_rp_layout()  


        arcpy.AddMessage("Script Complete")
        


    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
