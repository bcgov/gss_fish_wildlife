
# Author: Chris Sostad
# Ministry of Forests
# Created Date: January 30th, 2024
# Updated Date: 
# Description:
#   This script will recreate the Exhibit A text for the FTA Clearance Report.

# --------------------------------------------------------------------------------
# * SUMMARY

# - INPUTS:
#   


# - OUTPUTS

# --------------------------------------------------------------------------------
# * IMPROVEMENTS
# * Suggestions...
# --------------------------------------------------------------------------------
# * HISTORY

# Feb 05th, 2024
    # This script is used to generate the Exhibit A text for the FTA Clearance Report.

# FEB 06th, 2024
    # Added the referral section for BCTS Operating Areasr

# Feb 07, 2024
    # Added some details to fields
    # Added handling for RESULTS - Forest Cover Reserve - 

# Feb 09, 2024
    # Combined original Exhibit A script with roads script
    # Added the feature that adds the output.txt to the exhibit a layout

# Feb 10, 2024
    # Added the feature that turns on the roads layer
    # Added the feature that adds the output.txt to the exhibit a layout
    # Added conditional logic for roads layer that conflict with the cutblock
    # Changed the handling of text file. Now it reads the file, adds the new records and rewrites the file with the new content. 
    # Need to redo other text handling, whereby it starts by creating all of the headings and then for each section, it reads the file, 
    # adds the new records and rewrites the file with the new content.
    
# Feb 12, 2024
    # Finished the handling of the roads layer. It now processes the roads are within 10m (adjacent) and processes
    # roads in conflict and then subtracts the conflicts from adjacent to give you roads that are adjacent but not in conflict
    # and the roads that are in conflict with the cutblock

import arcpy
import os

# Get the client name to check for roads conflicts. This can later be obtained from FTEN Cut Block SVW (Pending) layer
# client = input("What is the client name?")
client = "CANADIAN FOREST PRODUCTS LTD."



print("Setting up the script.....")
arcpy.AddMessage

# Set the workspace (to the default project's geodatabase, for example)
arcpy.env.workspace = r'W:\for\RNI\DMK\General_User_Data\csostad_Clearances\Arc_Pro_FTA_Clearances_JF\Arc_Pro_FTA_Clearances.gdb'

aprx = arcpy.mp.ArcGISProject("CURRENT")

# Set the map object to the SCSP Conflicts Map
map_obj = aprx.listMaps("SNSCS Conflicts Map")[0] 

# Set the output folder for the text file
output_folder = r'W:\for\RNI\RNI\General_User_Data\CSostad'

# Path to save the .txt file
output_txt_path = os.path.join(output_folder, 'Exhibit_A_output.txt')

# Set up the text file to write to, create all of the headings

# Function to write initial headings to the text file
def create_initial_headings():
    headings = [
        "Save and Excepts:",
        "Exclude (by notation on report):",
        "Referral:",
        "Comments:",
        "Restricted Notes:",
        "Cut and Paste Phrases:"
    ]
    with open(output_txt_path, 'w') as file:
        for heading in headings:
            file.write(f"{heading}\n\n")  # Adding extra newline for spacing

create_initial_headings()

# Function to append data under specific heading
# This function assumes the heading already exists in the file
def append_data_under_heading(heading, data):
    with open(output_txt_path, 'r') as file:
        contents = file.readlines()
    
    # Find the index for the heading
    try:
        index = contents.index(f"{heading}\n") + 1  # Position to insert data
        while contents[index].strip():  # Move index past existing entries under this heading
            index += 1
    except ValueError:
        print(f"Heading '{heading}' not found. Adding it to the file.")
        contents.append(f"{heading}\n")
        index = len(contents)
    
    # Insert new data
    contents.insert(index, f"{data}\n")
    
    # Write back to the file
    with open(output_txt_path, 'w') as file:
        file.writelines(contents)




###############################################################################################################################################
#
# Standard Save and Excepts
#
###############################################################################################################################################
print("Running Save and Excepts.....")

# Layer information is a list of dictionaries that contains the layer to use for the selection and the fields that need to be copied to the txt file for the Exhibit A
# These are all the layers that are a basic save and except with no conditionals
layer_info = [
    {# MTA - Mineral and Placer Claims and Leases
        'layer': r"Clearance Layers\Mineral GROUP\All Placer and Mineral Claims",
        'fields': ['Tenure_Number_ID', 'Tenure_Type_Description']
    },
    {# MTA - Spatial - Handles 329585 MANSON CREEK DPLA - SCHEDULE K and the 330209 AREA #3 - DESIGNATED PLACER AREA
        'layer': r"Clearance Layers\Mineral GROUP\MTA Mineral Reserve Sites Spatial View",
        'fields': ['Site_Number_ID', 'Site_Name']
    },
    {# Handles Range and Tenure
        'layer': r"Clearance Layers\Range GROUP\FTEN Range",
        'fields': ['Forest_File_ID', 'MAP_BLOCK_ID', 'Client_Name']
    },
    {# Handles Critical Wildlife Habitat
        'layer': r"Clearance Layers\Critical Habitat\Critical Habitat for Federally-Listed Species at Risk - Posted - Colour Themed",
        'fields': ['Common_Name_English']
    },
    {# Handles Old Growth Technical Advisory Panel TAP - Priority Deferral Area
        'layer': r"Clearance Layers\Old Growth Strategic Review Deferral Area\OGSR Priority Deferral Area - TAP Classification Label - Outlined",
        'fields': ['PRIORITY_DEFERRAL_ID', 'TAP_CLASSIFICATION_LABEL']
    },
    {# Handles Traplines
        'layer': r"Clearance Layers\No Display GROUP\Trappers-Guides\Traplines LRDW",
        'fields': ['Trapline_Area_Identifier']
    }
]

select_features = r"Clearance Layers\Pending Tenures GROUP\FTEN Cut Block SVW (Pending)"

for info in layer_info:
    input_layer = info['layer']
    fields_to_read = info['fields']
    
    # Perform the selection
    arcpy.management.SelectLayerByLocation(
        in_layer=input_layer,
        overlap_type="INTERSECT",
        select_features=select_features,
        search_distance=None,
        selection_type="NEW_SELECTION",
        invert_spatial_relationship="NOT_INVERT"
    )

    # Read selected features with SearchCursor and append to the text file under "Save and Excepts"
    with arcpy.da.SearchCursor(input_layer, fields_to_read) as cursor:
        for row in cursor:
            row_data = ' '.join(str(item) for item in row)
            # Call the  function to append data under the specific heading
            append_data_under_heading("Save and Excepts:", f"- {row_data}")

print(f'Standard Save and Excepts were processed.')


##############################################################################################################
#
# Conditional Save and Excepts - • Ungulate Winter Range (UWRs)  Only if CONDITIONAL HARVEST ZONE (otherwise Exclude)  
#
##############################################################################################################
print("Running conditional Save and Excepts ie. UWR")


# Handle Ungulate Winter Range with conditional logic
uwr_layer = r"Clearance Layers\Wildlife\Ungulate Winter Range"
uwr_fields = ['UWR_Number', 'SPECIES_1', 'Timber_Harvest_Code']

# Perform the selection for Ungulate Winter Range
arcpy.management.SelectLayerByLocation(
    in_layer=uwr_layer,
    overlap_type="INTERSECT",
    select_features=select_features,
    search_distance=None,
    selection_type="NEW_SELECTION",
    invert_spatial_relationship="NOT_INVERT"
)

# Process Ungulate Winter Range records
with arcpy.da.SearchCursor(uwr_layer, uwr_fields) as cursor:
    for row in cursor:
        uwr_number, species_1, timber_harvest_code = row
        row_data = f'{uwr_number} {species_1} {timber_harvest_code}'

        if timber_harvest_code != "No Harvest Zone" and timber_harvest_code is not None:
            # Append to the "Save and Excepts:" heading
            append_data_under_heading("Save and Excepts:", f'- {row_data}')
        else:
            # Append to the "Exclude (by notation on report):" heading
            # If no record meets the condition to be excluded, no action is needed here based on your clarification
            append_data_under_heading("Exclude (by notation on report):", f'UWR- {row_data}')

print(f'Conditional Save and Excepts ie. UWR, were processed.')

##############################################################################################################
#
# Referral - • BCTS Operating Areas (resolve right away if it’s conflict with a BCTS submitted permit application) 
#		    #Comment name of business area (from source), and 
#           #Jeremy Greenfield – Timber Sales Manager (for PG District) 
#
##############################################################################################################
print("Running possible Referrals - BCTS Operating Areas")

# Referral - BCTS Operating Areas
bcts_layer = r"Clearance Layers\BCTS Operating Areas"

# Check if the BCTS Operating Areas layer exists
if not arcpy.Exists(bcts_layer):
    raise Exception(f'Layer {bcts_layer} does not exist in the map')

# Fields to retrieve from the BCTS Operating Areas layer
bcts_fields = ['OPERATING_AREA_NAME', 'TIMBER_SALES_OFFICE_NAME']

# Check if the specified fields exist in the layer
for field in bcts_fields:
    if field not in [f.name for f in arcpy.ListFields(bcts_layer)]:
        raise Exception(f'Field {field} does not exist in layer {bcts_layer}')

# Name of the Timber Sales Manager for PG District
manager_name = 'Jeremy Greenfield'

# Select BCTS Operating Areas that intersect with the specified features
arcpy.management.SelectLayerByLocation(
    in_layer=bcts_layer,
    overlap_type="INTERSECT",
    select_features=select_features,
    selection_type="NEW_SELECTION",
)

# Use a SearchCursor to iterate through the selected BCTS Operating Areas
with arcpy.da.SearchCursor(bcts_layer, bcts_fields) as cursor:
    for row in cursor:
        operating_area_name, timber_sales_office_name = row
        # Format the row data to include the manager name
        row_data = f"- {operating_area_name} {timber_sales_office_name} Manager {manager_name}"
        # Append the formatted data under the "Referral:" heading using the append_data_under_heading function
        append_data_under_heading("Referral:", row_data)

print(f'Referrals processed and appended to {output_txt_path}')


###################################################################################################################################
#
#  ROADS
#
###################################################################################################################################
'''
The roads are processed in three passes. The first pass will find all the roads that conflict with the pending tenure. 
If a road conflicts with the pending tenure and the client is the same, it will be added to the roads_exclude list,
which is then printed to the output file as an exclude by notation as "{source_name} {road_section_id} is intended to provide
access to {cut block id} and the client is the same".

If the client is different, it will be added to the roads_referral list, which is then printed to the output file as a referral.
"{source name} {road section id} belongs to {client name} and is in conflict with {cut block id}".

The second pass will find all the roads that are adjacent to the cutblock within a given buffer distance (currently 10 m).
The client name is irrelevant in this case and the the road will be added to comments list along with the text,
"{source name} {road section id} is adjacent to and intended to provide access to {cut block id}".

The third pass will subtract the roads that are intersecting the cutblock from the adjacent roads selection in order to isolate
the roads that are only adjacent to the cutblock. These roads will be added to the comments list along with the text,

'''
       

print("Running Roads....")

# Set the layer name to the roads layer
roads_layer_name = "FTEN Road Sections SVW (All)"
roads_layer = map_obj.listLayers(roads_layer_name)[0] if map_obj.listLayers(roads_layer_name) else None

if roads_layer and not roads_layer.visible:
    roads_layer.visible = True
    print(f"Layer '{roads_layer_name}' has been turned on.")
elif not roads_layer:
    print(f"Layer '{roads_layer_name}' not found in the map '{map_obj.name}'.")

# Variables for the roads layer and fields
roads_fields = ['Forest_File_ID', 'Road_Section_ID', 'Client_Name']
select_features = r"Clearance Layers\Pending Tenures GROUP\FTEN Cut Block SVW (Pending)"

####
#
# FIND CONLICTS IN ROADS
#
####


# Process adjacent roads within 11 meters
print("Processing Adjacent Roads Within 11 Meters...")
arcpy.management.SelectLayerByLocation(
    in_layer=roads_layer,
    overlap_type="WITHIN_A_DISTANCE",
    select_features=select_features,
    search_distance="11 Meters",
    selection_type="NEW_SELECTION"
)



# Run select by location on the selected roads to find only the roads that intersect (not adjacent) and select those as a subset
arcpy.management.SelectLayerByLocation(
    in_layer=roads_layer,
    overlap_type="INTERSECT",
    select_features=select_features,
    selection_type="SUBSET_SELECTION"
)


# Read the selected conflicting roads and write them to the output file as either exclude or referral
with arcpy.da.SearchCursor(roads_layer, roads_fields) as cursor:
    for row in cursor:
        forest_file_id, road_section_id, client_name = row
        record_string = f"- {forest_file_id} {road_section_id} {client_name}"
        if client_name == client:
            append_data_under_heading("Exclude (by notation on report):", record_string)
        else:
            append_data_under_heading("Referral:", record_string)

####
#
# FIND ADJACENT ROADS
#
####

# Process adjacent roads within 11 meters
print("Processing Adjacent Roads Within 11 Meters...")
arcpy.management.SelectLayerByLocation(
    in_layer=roads_layer,
    overlap_type="WITHIN_A_DISTANCE",
    select_features=select_features,
    search_distance="11 Meters",
    selection_type="NEW_SELECTION"
)

# Run select by location on the selected roads to find only the roads that intersect (not adjacent) and remove those from the selection 
# leaving only adjacent roads
arcpy.management.SelectLayerByLocation(
    in_layer=roads_layer,
    overlap_type="INTERSECT",
    select_features=select_features,
    selection_type="REMOVE_FROM_SELECTION"
)

# Write the adjacent roads to the output file as comments
with arcpy.da.SearchCursor(roads_layer, roads_fields) as cursor:
    for row in cursor:
        forest_file_id, road_section_id, client_name = row
        record_string = f"- {forest_file_id} {road_section_id} {client_name}"
        comment_string = f"- {forest_file_id} {road_section_id} {client_name} is adjacent to and does not conflict with {select_features}"
        append_data_under_heading("Comments:", record_string)
        append_data_under_heading("Cut and Paste Phrases:", comment_string)



# # Process adjacent roads within 10 meters, excluding intersecting roads
# print("Processing Adjacent Roads Within 10 Meters...")
# arcpy.management.SelectLayerByLocation(
#     in_layer=roads_layer,
#     overlap_type="WITHIN_A_DISTANCE",
#     select_features=select_features,
#     search_distance="10 Meters",
#     selection_type="NEW_SELECTION"
# )

# # Reapply intersect selection to subtract these from the adjacent selection
# arcpy.management.SelectLayerByLocation(
#     in_layer=roads_layer,
#     overlap_type="INTERSECT",
#     select_features=select_features,
#     selection_type="REMOVE_FROM_SELECTION"
# )


print("Roads Script complete")
       
       
###################################################################################################################################
#
# Write the outputs to the Ex A layout
#
###################################################################################################################################

# Read the output text file
with open(output_txt_path, 'r') as file:
    output_text = file.read()

# Access the layout by its name
layout = aprx.listLayouts("1_PORTRAIT_ExA_legal")[0]  

# Find the text element by its name and update its text
for elem in layout.listElements("TEXT_ELEMENT"):
    if elem.name == "SaveAndExcepts":  
        elem.text = output_text
        break
else:
    print("Text element 'SaveAndExcepts' not found in the layout 'Ex_A_2024'.")



print("The text element 'SaveAndExcepts' in the layout 'Ex_A_2024' has been updated.")
arcpy.AddMessage("The text element 'SaveAndExcepts' in the layout 'Ex_A_2024' has been updated.")
print("Total script finished!")