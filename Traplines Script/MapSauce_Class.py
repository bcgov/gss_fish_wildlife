import arcpy
import os
import datetime
from dotenv import load_dotenv

aprx = arcpy.mp.ArcGISProject("CURRENT")
workspace = arcpy.env.workspace

# Get the hidden file path from the .env file
output_path = os.getenv('WORKSPACE_PATH') # File path 


# Variables to be changed each time you run the script
application_id = "A94309" #Used to be proponentName



# Create the permit_str
global permit_str
application_str = f"{application_id}" # this f string can be used to build more descriptive strings

# Set up the datetime module to be used in naming folders
now = datetime.datetime.now()
day = now.strftime("%d")
month = now.strftime("%M")
year = now.strftime("%Y")

# root self.dir_path for all proponents--from which to build
global dir_path
# dir_path = os.path.join(r"\\spatialfiles2.bcgov\archive\FOR\RNI\DMK\Library\ExA_FN_Folders", year, "Licensee_NRFL_TESTING")
dir_path = output_path


class MapInfo:
    def __init__ (self, map_name, map_frame, layout, application_st, dir_path):
        self.map_name = map_name
        self.map_frame = map_frame
        self.layout = layout
        self.application_str = application_str
        self.dir_path = dir_path
        self.application_dir = os.path.join(dir_path, application_str)
    
    # Function to create a directory path if it doesn't exist
    def create_directory(self):
        if not os.path.exists(self.dir_path):
            os.makedirs(self.dir_path)
            arcpy.AddMessage(f"Directory created: {self.dir_path}")
        else:
            arcpy.AddMessage(f"Directory already exists: {self.dir_path}")
    
        
    def show_map_obj(self):
        '''
        A sample method I created for learning and testing
        '''
        map_obj = aprx.listMaps(self.map_name)[0] #self.map name means go into the collected params of the instance and pull out the one called map_name
        print(map_obj)
    
       
    def update_layer_connection(self, layer_name):
        '''This function will update the layer connection of the Application layer
        to use it, call the instance (ie. fn_map.update_layer_connection and then the layer name that you want to update)'''
        map_obj = aprx.listMaps(self.map_name)[0]
        layers = map_obj.listLayers(layer_name)
        if not layers:
            print(f"No layer named '{layer_name}' found in {map_obj.name}.")
            return

        target_lyr = layers[0]
        new_shapefile_path = os.path.join(self.application_dir, f"{self.application_str}.shp")
        new_conn_props = {
            'connection_info': {'database': os.path.dirname(new_shapefile_path)},
            'dataset': os.path.basename(new_shapefile_path).replace('.shp', ''),
            'workspace_factory': 'Shape File'
        }
        target_lyr.updateConnectionProperties(target_lyr.connectionProperties, new_conn_props)
        print(f"Data source updated for layer: {layer_name} in map: {self.map_name}")

        
    def round_scale(self):
        '''This function will round the scale of the map frame to the nearest 10,000'''
        # map_obj = aprx.listMaps(self.map_name)[0]
        
        # Next, find the layout object by name
        layout_obj = None
        for layout in aprx.listLayouts():
            if layout.name == self.layout:
                layout_obj = layout
                break

        if layout_obj is None:
            arcpy.AddMessage(f"No layout named '{self.layout}' found.")
            return 
        
        mf_list = layout_obj.listElements("MAPFRAME_ELEMENT", self.map_frame)
        if mf_list:  # Check if the list is not empty
            mf = mf_list[0]  # Access the first map frame in the list
            current_scale = mf.camera.scale  # Access the scale property directly
            arcpy.AddMessage(f"Current scale is: {current_scale}")

            # Round the scale to the nearest 10,000
            rounded_scale = round(current_scale / 10000) * 10000
            arcpy.AddMessage(f"Rounded scale is: {rounded_scale}")

            # Set the map frame to the new rounded scale
            mf.camera.scale = rounded_scale  # Set the scale property directly
            arcpy.AddMessage(f'Zooming the map to the new scale of {rounded_scale}')
        else:
            arcpy.AddMessage("No map frame found with the given name.")  

    def create_mapsheet_centroid(self, target_layer):
        '''
        This function will calculate feature to point and put a point into each of the pending cut blocks on FTEN Cut Block SVW (Pending)
        It will then intersect the points with the 20k BCGS Grid to find the mapsheet number for each cut block. Next it creates a set()
        of unique mapsheet numbers and makes them available to the rest of the script. The output of this function is the mapsheet_centroid, 
        and ften pending centroid which is saved to the permit self.dir_path. Future modification could either make these temporary layers or delete
        them after the script is complete.
        Inputs: 
        target_layer - the layer you want to find the BCGS grid for
        path - the path where you want the centroid to be save to
        BCGS_Grid - the BCGS grid (entered as the layer name (str) from your contents) you want to intersect with the target_layer
        '''
        map_name_obj = aprx.listMaps(self.map_name)[0]
        mapsheet = map_name_obj.listLayers("20k BCGS Grid")[0]
        # Set overwrite to True
        arcpy.env.overwriteOutput = True
        arcpy.AddMessage("Calculating Mapsheet Centroid...")
        print('Calculating Mapsheet Centroid...')

        # Check that target_layer exists and has records
        if not arcpy.Exists(target_layer):
            raise FileNotFoundError(f"Create Centroid Error: Feature class {target_layer} does not exist.")
        else:
            arcpy.AddMessage(f"Create Centroid Function: Feature class {target_layer} Exists.") 
            print(f"Create Centroid Function: Feature class {target_layer} Exists.") 
        # Check to see that the target layer has records
        count = arcpy.management.GetCount(target_layer)
        if count == 0:
            arcpy.AddMessage(f"Centroid Function Error: Feature class {target_layer} has no records.")
            print(f"Centroid Function Error: Feature class {target_layer} has no records.")
            arcpy.AddError(f"Centroid Function Error: Feature class {target_layer} has no records.")
            raise ValueError(f"Centroid Function Error: Feature class {target_layer} has no records.")
        else:
            arcpy.AddMessage(f"Centroid Function: Feature class {target_layer} has {count} records, creating centroid.")
            print(f"Centroid Function: Feature class {target_layer} has {count} records, creating centroid.")



            # Create a centroid from the FTEN Cut Block SVW (Pending) layer
            target_layer_centroid = os.path.join(self.dir_path, "target_layer_centroid.shp")
            arcpy.management.FeatureToPoint(in_features=target_layer, out_feature_class=target_layer_centroid, point_location="INSIDE")

        # Check to see if target_layer_centroid has records
        count = arcpy.management.GetCount(target_layer_centroid)
        if count == 0:
            arcpy.AddMessage(f"Feature class {target_layer_centroid} in memory has no records.")
            print(f"Feature class {target_layer_centroid} in memory has no records.")
            arcpy.AddError(f"Feature class {target_layer_centroid} in memory has no records.")
            raise ValueError(f"Feature class {target_layer_centroid} in memory has no records.")
        else:
            arcpy.AddMessage(f"Feature class {target_layer_centroid} in memory has {count} records, intersecting with 20k BCGS Grid.")
            print(f"Feature class {target_layer_centroid} in memory has {count} records, intersecting with 20k BCGS Grid.")

        # Intersect the ften pending Centroid with the 20k BCGS Grid
        mapsheet_centroid = os.path.join(self.dir_path, "mapsheet_centroid.shp")

        arcpy.analysis.Intersect(in_features=[target_layer_centroid, mapsheet], out_feature_class=mapsheet_centroid)


        # Check to see if the mapsheet_centroid feature class exist
        if arcpy.Exists(mapsheet_centroid):
            print(f" {mapsheet_centroid} exists.")


        # Check to see if the mapsheet_centroid feature has records, if it has no records, skip the search cursor and print an error message.
        count = arcpy.management.GetCount(mapsheet_centroid)
        if count == 0:
            arcpy.AddMessage(f"Feature class {mapsheet_centroid} has no records.")
            arcpy.AddError(f"Feature class {mapsheet_centroid} has no records. You will have manually enter mapsheet numbers.")
            raise ValueError(f"Feature class {mapsheet_centroid} has no records.")

        # Use a search cursor to read each row and create a list of mapsheets. Print only each unique mapsheet number to messages.
        arcpy.AddMessage(f"Feature class {mapsheet_centroid} has {count} records, adding Mapsheet Number.")
        global unique_mapsheet_set
        unique_mapsheet_set = set()
        with arcpy.da.SearchCursor(mapsheet_centroid, "MAP_TILE_D") as cursor:
            for row in cursor:
                unique_mapsheet_set.add(row[0])
                arcpy.AddMessage(f"Unique Mapsheet Numbers: {unique_mapsheet_set}")
        mapsheet_centroid_obj = map_name_obj.listLayers("mapsheet_centroid")[0]
        target_layer_centroid_obj = map_name_obj.listLayers("target_layer_centroid")[0]

        map_name_obj.removeLayer(mapsheet_centroid_obj)
        map_name_obj.removeLayer(target_layer_centroid_obj)

        return unique_mapsheet_set
        print('Centroid Function complete')
        arcpy.AddMessage("Centroid Function Finished")


        print(unique_mapsheet_set)
        
        

fn_map = MapInfo("FN Consult Map", "Layers Map Frame", "1_PORTRAIT_FN_legal", permit_str, dir_path)  # Load up the instance with all params      
ex_A_map = MapInfo("7_LANDSCAPE_EXA_AnsiE_Updated2023", "Main Map Frame", "1_PORTRAIT_ExA_legal", permit_str, dir_path)        


# Start using the methods

fn_map.show_map_obj()      
# MapInfo.show_map_obj(fn_map) # Says go ahead and find the variable from this instance, it contains all you need
# MapInfo.update_layer_connection(fn_map)        

fn_map.round_scale()
ex_A_map.round_scale()
fn_map.create_mapsheet_centroid("FN Selected Features")
