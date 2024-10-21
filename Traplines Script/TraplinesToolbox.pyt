# ArcGIS Pro 3 Toolbox Template

#===========================================================================
# Script name: Arcgis Pro 3 Toolbox Template
# Author: https://pro.arcgis.com/en/pro-app/latest/arcpy/geoprocessing_and_python/a-template-for-python-toolboxes.htm

# Created on: 10/21/2024
# 
#

# 
#
# 
#============================================================================

import arcpy


class Toolbox:
    def __init__(self):
        """Define the toolbox (the name of the toolbox is the name of the
        .pyt file)."""
        self.label = "Descriptive Name of your Toolbox"
        self.alias = "toolbox"

        # List of tool classes associated with this toolbox
        self.tools = [NameOfYourTool]  
        # Insert the name of each tool in your toolbox if you have more than one. 
        # i.e. self.tools = [FullSiteOverviewMaps, ExportSiteAndImageryLayout, Amendment]


class NameOfYourTool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Tool"
        self.description = ""

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 
        
        
        #This parameter is the file number of the application
        feature_name = arcpy.Parameter(
            displayName = "feature_name",
            name="feature_name",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        


        parameters = [feature_name]  # Each parameter name needs to be in here, separated by a comma

        return parameters


    
    def isLicensed(self):
        """Set whether the tool is licensed to execute."""
        return True
    

    def updateParameters(self, parameters):
        """Modify the values and properties of parameters before internal
        validation is performed.  This method is called whenever a parameter
        has been changed."""
        return

    def updateMessages(self, parameters):
        """Modify the messages created by internal validation for each tool
        parameter. This method is called after internal validation."""
        return

    
    #NOTE This is where you cut a past your code
    def execute(self, parameters, messages):
        
        
        import arcpy
        import os

        feature_name = parameters[0].valueAsText  # Take the first parameter **0 indexed!!** and assign in the the variable file_num

        # addded .shp to trapline boundaries export
        # moved the assignment of layout and maps to objects to the top of the script
        # changed variable names so it is easier for others (namely me) to read

        # Create or get the feature_layer object from ArcGIS Pro's content pane or the appropriate source
        aprx = arcpy.mp.ArcGISProject("CURRENT")


        # Define paths for the directories -  - r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\ebreton\WLRS\Projects\Trapline_Territories\Data'
        #workspace = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\SharedWork\Trapline_Territories'
        workspace = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\ebreton\WLRS\Projects\Trapline_Territories\Data'
        #kml_dir = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\SharedWork\Trapline_Territories\Kml'
        kml_dir = os.path.join(workspace, 'Kml')
        aprx_dir = os.path.join(workspace, 'Aprx')
        data_dir = os.path.join(workspace, 'Data')
        pdf_dir = os.path.join(workspace, 'pdf')

        # Path to the template APRX - W:\srm\nel\Local\Geomatics\Workarea\ebreton\WLRS\Projects\Trapline_Territories
        template_aprx_path = r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\ebreton\WLRS\Projects\Trapline_Territories\Temp_Trapline.aprx'

        # Convert main maps and layers to objects early in the script so when others are working in the script, they can scroll to the top and confirm it has been assigned rather
        # than having to search through the script to find where it was assigned.

        #NOTE: 

        map_obj = aprx.listMaps('Map')[0]
        layout = aprx.listLayouts("Layout")[0] 
        all_trapline_cabins_obj = map_obj.listLayers("All Trapline Cabins")[0]
        all_trapline_boundaries_obj = map_obj.listLayers("All Trapline Boundaries")[0]
        # map_name = map_obj

        # Function to create a directory if it doesn't exist
        def create_directory(directory):
            if not os.path.exists(directory):
                os.makedirs(directory)
                print(f"Directory created: {directory}")
            else:
                print(f"Directory already exists: {directory}")

        # Create the main directories
        create_directory(aprx_dir)
        create_directory(data_dir)
        create_directory(pdf_dir)

        # Define the feature name (extracted from the query)
        # feature_name = 'TR0409T003'

        # Create subdirectories named after the feature in each main directory
        aprx_subdir = os.path.join(aprx_dir, feature_name)
        data_subdir = os.path.join(data_dir, feature_name)
        pdf_subdir = os.path.join(pdf_dir, feature_name)

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

        print("Step 2 - Creating Application Polygon")

        # Set the definition query and assign it to a variable

        expression = f"{arcpy.AddFieldDelimiters(arcpy.env.workspace, 'TRAPLINE_1')} = '{feature_name}'"


        # Apply the expression to the layer "Application" in the Site Map (Defined as a global variable at the top of the script)
        all_trapline_boundaries_obj.definitionQuery= expression
        print(f"         Definition query {expression} set for all trapline boundaries")


        # Error Handling -  Use a SearchCursor to count the number of records returned by the definition query
        count = 0
        with arcpy.da.SearchCursor(all_trapline_boundaries_obj, "*") as cursor:
            for row in cursor:
                count += 1
                del cursor
        # Check if the count is 0 and display a message if the definition query was successful or not.
        if count == 0:
            print(f"         No records returned by definition query: {expression}, please check the file number and try again")
        else:
            print(f"         Definition query {expression} set for layer: {all_trapline_boundaries_obj}. Records returned: {count}")



        # NOTE Now that we have a def query ofn trapline bouundaries do we need to export thepolygon?
        # Specify the trapline boundary layer you want to export features from
        application_trapline_boundary_path = os.path.join(data_dir, 'Scripting Data', 'Trapline_Boundaries_Export.shp')
        print(f"Trapline boundary path: {application_trapline_boundary_path}")

        # Specify the output file path for the exported feature
        application_trapline_boundary = os.path.join(data_subdir, f'{feature_name}.shp')
        print(f"Output feature path: {application_trapline_boundary}")


        # # Create a query to select the first feature
        # query = f"TRAPLINE_1 = '{feature_name}'"  # Adjust this query based on the field you're using to identify features

        # print("Query:", query)  # Print the query for debugging purposes

        #NOTE 
        #Put in a check here to see if Def query has 0 records or more than one record.

        ###########################################################################################################################################
        #
        # Step 3 - Export the Feature to a Shapefile
        #
        ###########################################################################################################################################

        # Create a feature layer to select the specific feature
        try:
            arcpy.management.MakeFeatureLayer(all_trapline_boundaries_obj, "temp_layer") 
            # map_obj.addDataFromPath(application_trapline_boundary)
            # Make feature layer always creates a temporary layer. Its held in memory and will be deleted upon exit.  
            print("Feature layer created.")
        except arcpy.ExecuteError as e:
            print("MakeFeatureLayer_management error:", e)
            exit()  # Exit the script if there's an error

        # Check if the temporary layer exists
        if arcpy.Exists("temp_layer"):
            try:
                
                arcpy.management.CopyFeatures("temp_layer", application_trapline_boundary) # Layer has been created (No longer just a path)
                print("Export process complete.")
            except arcpy.ExecuteError as e:
                print("CopyFeatures_management error:", e)
        else:
            print("Temporary layer 'temp_layer' does not exist.")
            exit()

        # Add a new field for the area in hectares if it doesn't already exist
        area_field = "Area_ha"
        if area_field not in [f.name for f in arcpy.ListFields(application_trapline_boundary)]:
            arcpy.management.AddField(application_trapline_boundary, area_field, "DOUBLE")
            print(f"Field '{area_field}' added to the feature class.")


        # Define the function to calculate area in hectares
        def calculate_area_in_hectares(geometry):
            print("Defining calculate_area_in_hectares function...")
            
            area_sq_meters = geometry.area  # Area in square meters
            return area_sq_meters / 10000  # Convert to hectares

        # Calculate the area for each polygon and update the new field
        print("Calculating area in hectares...")
        with arcpy.da.UpdateCursor(application_trapline_boundary, ["SHAPE@", area_field]) as cursor:
            for row in cursor:
                area_hectares = calculate_area_in_hectares(row[0])
                row[1] = area_hectares  # Update the field with the area in hectares
                cursor.updateRow(row)
                del cursor  # Delete the cursor after it has completed its task
                print("Deleted cursor")
                # Format the area to two decimal points
                formatted_area = f"{area_hectares:.2f} ha."
                print(f"Formatted area: {formatted_area}")

        print(f"Area in hectares has been added to the field '{area_field}'.")

        # #  Export the feature layer to KML and save it in the corresponding feature's subdirectory
        # kml_output = os.path.join(kml_dir, f'{feature_name}.kml')

        # # Convert the feature layer to KML
        # try:
        #     arcpy.LayerToKML_conversion("temp_layer", kml_output)
        #     print(f"KML file created at: {kml_output}")
        # except arcpy.ExecuteError as e:
        #     print(f"Error exporting to KML: {e}")



        # Ensure there are maps and layouts in the template project
        #NOTE commented out, not sure if this is necessary we know the map and layout are in the project. Just trying to speed things up bit. Can always uncomment if needed.
        # if not map_obj:
        #     raise ValueError("No maps found in the project.")
        # if not layout:
        #     raise ValueError("No layouts found in the project.")



        #NOTE This can be removed as we can rename the layers lower down in the script
        # Update the feature layer name in the content pane
        # new_layer_name = f"{feature_name} ({formatted_area})"
        # trapline_app_layer.name = new_layer_name
        # print(f"Updated feature layer name: {trapline_app_layer.name}.... will now add it to the contents pane")

        #NOTE Removed this check of the layer name, if you want to use it, it needs to be moved to the bottom after we rename the layer
        # # Verify the updated TB layer name
        # layer_updated = False
        # for lyr in map_obj.listLayers():
        #     if lyr.name == new_boundaries_layer_name:
        #         layer_updated = True
        #         print(f"{lyr.name} successfully added to contents pane -:) ")
        # if not layer_updated:
        #     print(f"Layer name update failed. Current layers: {[lyr.name for lyr in map_obj.listLayers()]}")

        ##############################################################################################################
        #
        # Step 4 - Clip the "Trapline Cabins" layer based on the feature layer
        #
        ##############################################################################################################

        # Specify the "Trapline Cabins" layer name
        #NOTE used the all cabins layer in contents pane instead of the path

        # all_trapline_cabins_layer = os.path.join(data_dir, 'TRAPLINE_CABINS_Polygon.shp')
        clipped_cabins_output = os.path.join(data_subdir, f'{feature_name}_Cabins.shp')
        print
        print(f"Output feature path: {clipped_cabins_output}")
        # Clip the "Trapline Cabins" layer based on the feature layer\

        arcpy.analysis.Clip("All Trapline Cabins", application_trapline_boundary, clipped_cabins_output)  # After the clip, the clipped cabins output path has now become a layer

        print(f"Clipping of trapline boundary to Crown Lands layer completed. --> {feature_name}_Cabins.shp was added to contents pane") 

        #Check to see if CROWN_LAND field exists and if it is None or empty; if it does exist create new_crown_cabins_str as a feature layer.  
        crown_land_field = "CROWN_LAND"  # Assuming this field exists
        if crown_land_field not in [f.name for f in arcpy.ListFields(clipped_cabins_output)]:
            print(f"{crown_land_field} not found in attribute table.")

        else:
                # Update the names of the resulting clipped cabins based on the CROWN_LAND field
            Crown_Num_Values =[]
            with arcpy.da.UpdateCursor(clipped_cabins_output, [crown_land_field]) as cursor:
                for row in cursor:
                    #Capture the crown land # from crown_land_field
                    crown_land_field =row[0]

                    #Adding all crown land #'s to the Crown_Num_Values variable - to be used later to rename trapline cabin feature
                    Crown_Num_Values.append(crown_land_field)

                    #Create a new name for the Trapline Cabin Feature layer naming by appending the joined Crown_Num_Values
                    Crown_Num_Values_String ="_".join(map(str, Crown_Num_Values))

        #NOTE could change this to crown_num
        new_crown_cabins_str = f"{feature_name}_Cabins_{Crown_Num_Values_String}" # Here we are only assigning a string of letter and numbers to a variable 
        print(f"New Crown cabins variable: {new_crown_cabins_str} created.")

        # #NOTE added the .shp to the end of the file name
        # NOTE add a check before this so create feature class doesnt occur when the query returns nothing
        # arcpy.management.CreateFeatureclass(data_subdir, f"{new_crown_cabins_str}.shp")  #Here we are creating feature class for new_crown_cabins_str

        #NOTE could assign create feature class to new variable so its not called .str  not really necessary though
        print (f"{new_crown_cabins_str} feature class created in {data_subdir}. ")

        # Find the text element (Map Title) and update its text to the feature name
        text_updated = False
        #NOTE no need to search for the layout again as we already assigned it to a python object at the top
        #NOTE changed text6 to Map Title to match the name of the element in the layout
        # for lyt in aprx.listLayouts("Layout"):
        for elm in layout.listElements("TEXT_ELEMENT"):
            if elm.name == "Map Title":
                elm.text = f"Trapline {feature_name}"
                print("\tProponent text element changed")

        #NOTE We dont want to save the Aprx file because we will be changing the names of the layers in the contents and then it wont work next time
        # Save the project to preserve changes
        # aprx.saveACopy(os.path.join(aprx_subdir, f'{feature_name}.aprx'))

        # #NOTE
        # # Exported all Trapline Boundaries to a shapefile, changed the name of the BCGW layer to "Trapline Boundaries BCGW" so there arent two layers with the same name


        #remove temp_layer from contents pane
        temp_layer_obj = map_obj.listLayers("temp_Layer")[0]
        arcpy.management.Delete(temp_layer_obj)
        print("temp_layer has been removed from contents pane")


        # NOTE, we are using the layer now so we don't want to remove if from the contents pane, we will rename it instead
        #remove trapline boundary "feature name" layer from contents pane

        # #remove trapline boundary "feature name" layer from contents pane
        # feature_name_obj = map_obj.listLayers(f'{feature_name}')[0]
        # arcpy.management.Delete(feature_name_obj)
        # print("Feature name layer has been removed from contents pane")



        # feature_name_cabins_obj = map_obj.listLayers(f'{feature_name}_Cabins')[0]
        # arcpy.management.Delete(feature_name_cabins_obj)
        # print("clipped cabins output layer has been removed from contents pane")

        #TODO NOTE to Chris could we change the name when it was created?
        #DELETE
        # new_cabins_layer_name = f'Trapline Cabin {feature_name}'
        # all_trapline_cabins_obj.name = new_cabins_layer_name

        #NOTE Need to find the new layers we added in the contents pane
        new_boundaries_layer_name = f"{feature_name} ({formatted_area})"

        for layer in map_obj.listLayers():
            if layer.name == f"{feature_name}_Cabins":
                print(f"(Changing {layer.name} name to {new_crown_cabins_str})")
                layer.name = new_crown_cabins_str
                
            
            if layer.name == f"{feature_name}":
                print(f"(Changing {layer.name} name to {new_boundaries_layer_name}")
                layer.name = new_boundaries_layer_name
                
        # all_trapline_boundaries_obj.name = new_boundaries_layer_name

        # Turn off unused layers
        all_trapline_boundaries_obj.visible = False
        all_trapline_cabins_obj.visible = False

        #attempt to zoom to Trapline Boundary feature and round scale to nearest 5,000

        #set the name of the zoom layer created earlier in the script "new_boundaries_layer_name = f"{feature_name} ({formatted_area})""
        zoom_feature_layer = new_boundaries_layer_name 

        #get the layout and map frame 
        lyt = aprx.listLayouts()[0]

        mf = lyt.listElements('mapframe_element', 'Map Frame')[0]

        # use the zoom_feature_layer for zooming
        arcpy.SelectLayerByAttribute_management(zoom_feature_layer, "NEW_SELECTION", "1=1")

        # zoom to selected featues within the map fram using the zoom feature layer = new_boundaries_layer_name
        mf.zoomToAllLayers(True)
        #clear selection
        arcpy.SelectLayerByAttribute_management(zoom_feature_layer, "CLEAR_SELECTION")

        mf.camera.scale = 250000
        print(f"Zoomed to feature in {zoom_feature_layer} and set scale to : {mf.camera.scale}")

        #Script runs successfully to this point.
        '''
        #Not Exporting Correctly
        pdf_file_name = (f"{feature_name}.pdf")
        # Global exportToPdf function
        def exportToPdf(layout, pdf_dir, pdf_file_name):
            out_pdf = f"{pdf_dir}\\{pdf_file_name}"
            layout.exportToPDF(
                out_pdf=out_pdf,
                resolution=300,  # DPI
                image_quality="BETTER",
                jpeg_compression_quality=80  # Quality (0 to 100)        
            )

        #Then Call the function


        # Export the layout to PDF
        layout.exportToPDF(f"{pdf_dir}\\{pdf_file_name}")    

        print("Trapline Boundary Map Automation has been completed")

        mf.exportToPDF(os.path.join(r'\\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\ebreton\WLRS\Projects\Trapline_Territories\pdf', f'{feature_name}.pdf'))

        print(f"{feature_name}.pdf has been exported to --> \\spatialfiles.bcgov\work\srm\nel\Local\Geomatics\Workarea\ebreton\WLRS\Projects\Trapline_Territories\pdf" )



        #function to round scale to nearest 250,000
        def round_scale_to_nearest_250000(scale):
            return round(scale / 250000) * 250000

        current_scale = mf.camera.scale

        rounded_scale = round_scale_to_nearest_250000(current_scale)
        mf.camera.scale = rounded_scale
        print(f"Zoomed to feature in {zoom_feature_layer} and set scale to : {rounded_scale}")

        #clear selection
        arcpy.SelectLayerByAttribute_management(zoom_feature_layer, "CLEAR_SELECTION")

        print("Zoom completed")'''
                
                
                

        # Assign your parameters to variables

        # proponent_name = parameters[1].valueAsText
        # cp_ID = parameters[2].valueAsText
        # esf_id = parameters[3].valueAsText
        # rp_ID = parameters[4].valueAsText
        # rp_amendment = parameters[5].valueAsText
        # rp_sections = parameters[6].valueAsText
        # sup_ID = parameters[7].valueAsText
        
        # Now write your script
        arcpy.AddMessage(f"Hello World, the file number is {file_num}, the proponent name is {proponent_name}, and the cutting permit number is {cp_ID}")
        
        return

    def postExecute(self, parameters):
        """This method takes place after outputs are processed and
        added to the display."""
        return
    
#NOTE add another tool to the toolbox if you want
class MySecondTool(object):
    def __init__(self):
        
        """This tool will prep all required data for an individual crown tenure - to be used to add/subtract amendment - it will create the 
        amendment polygon, centroid, and splitline, and calculate geometries for centroid and splitline"""
        
        self.label = "Descriptive Name of your Second Tool"
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
        
        parameters = [amend_file_num] #Each parameter name needs to be in here, separated by a comma

        return parameters

    
    def execute(self,parameters,messages):
        
        # Bring in parameters to the function to be used as variables 
        amend_file_num = parameters[0].valueAsText
        
        arcpy.addMessage(f"Hello World, the file number is {amend_file_num}")
        
        ############################################################################################################################################################################
        #
        # Create the shapefile polygon layer to be used for the Amendment.
        #
        ############################################################################################################################################################################
