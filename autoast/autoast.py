# autoast is a script for batch processing the automated status tool
# author: wburt
# copyrite Governent of British Columbia
# Copyright 2019 Province of British Columbia

# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at 

#    http://www.apache.org/licenses/LICENSE-2.0

# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.


import sys
import os
import shutil
from openpyxl import Workbook,load_workbook
from dotenv import load_dotenv
import multiprocessing
import geopandas
import arcpy
import fiona
import datetime
            
            


print("Starting Script")

# Load the default environment
load_dotenv()

# Get the toolbox path from environment variables
ast_toolbox = os.getenv('TOOLBOX')
if ast_toolbox is None:
    print("Unable to find the toolbox. Check the path in .env file")
    exit() 


# Import the toolbox
try:
    arcpy.ImportToolbox(ast_toolbox)
    print(f"AST Toolbox imported successfully.")
except Exception as e:
    print(f"Error importing toolbox: {e}")
    exit()

###############################################################################################################################################################################
#
# Set up the database connection
#
###############################################################################################################################################################################

# Get the secret file containing the database credentials
SECRET_FILE = os.getenv('SECRET_FILE')

# If secret file found, load the secret file and display a print message, if not found display an error message
if SECRET_FILE:
    load_dotenv(SECRET_FILE)
    print(f"Secret file {SECRET_FILE} found")
else:
    print("Secret file not found")

# Assign secret file data to variables    
DB_USER = os.getenv('BCGW_USER')
DB_PASS = os.getenv('BCGW_PASS')


# If DB_USER and DB_PASS found display a print message, if not found display an error message
if DB_USER and DB_PASS:
    print(f"Database user {DB_USER} and password found")
else:
    print("Database user and password not found")

# Define current path of the executing script
current_path = os.path.dirname(os.path.realpath(__file__))

# Create the connection folder
connection_folder = 'connection'
connection_folder= os.path.join(current_path,connection_folder)
    
# Check for the existance of the connection folder and if it doesn't exist, print an error and create a new connection folder
if not os.path.exists(connection_folder):
    print("Connection folder not found, creating new connection folder")
    os.mkdir(connection_folder)

# Check for an existing bcgw connection, if there is one, remove it
if os.path.exists(os.path.join(connection_folder,'bcgw.sde')):
    os.remove(os.path.join(connection_folder,'bcgw.sde'))

# Create a bcgw connection
bcgw_con = arcpy.management.CreateDatabaseConnection(connection_folder,
                                                    'bcgw.sde',
                                                    'ORACLE',
                                                    'bcgw.bcgov/idwprod1.bcgov',
                                                    'DATABASE_AUTH',
                                                    DB_USER,
                                                    DB_PASS,
                                                    'DO_NOT_SAVE_USERNAME')

print("new db connection created")

################################################################################################################################################################################

arcpy.env.workspace = bcgw_con.getOutput(0)

print("workspace set to bcgw connection")
class AST_FACTORY:
    ''' AST_FACTORY creates and manages status tool runs '''
    XLSX_SHEET_NAME = 'ast_config'
    AST_PARAMETERS = {
        0: 'region',
        1: 'feature_layer',
        2: 'crown_file_number',
        3: 'disposition_number',
        4: 'parcel_number',
        5: 'output_directory',
        6: 'output_directory_same_as_input',
        7: 'dont_overwrite_outputs',
        8: 'skip_conflicts_and_constraints',
        9: 'suppress_map_creation',
        10: 'add_maps_to_current',
        11: 'run_as_fcbc',
        12: 'ast_condition'
    }
    
    ADDITIONAL_PARAMETERS = {
        13: 'file_number'
    }
    
    AST_CONDITION_COLUMN = 'ast_condition'
    AST_SCRIPT = ''

    def __init__(self, queuefile, db_user, db_pass) -> None:
        self.user = db_user
        self.user_cred = db_pass
        self.queuefile = queuefile
        self.jobs = []

    def load_jobs(self):
        '''loads jobs from the jobqueue'''
        print("Loading jobs")
        self.jobs = []
        assert os.path.exists(self.queuefile)
        if os.path.exists(self.queuefile):
            wb = load_workbook(filename=self.queuefile)
            ws = wb[self.XLSX_SHEET_NAME]
            header = list([row for row in ws.iter_rows(min_row=1, max_col=None, values_only=True)][0])
            data = [row for row in ws.iter_rows(min_row=2, max_col=None, values_only=True)]
            for d in data:
                job = dict()
                job_condition = None
                for k, v in zip(header, d):
                    if k is not None and k.lower() == self.AST_CONDITION_COLUMN.lower():
                        if v is not None:
                            job_condition = v
                        elif v is None:
                            job_condition = ""
                            v = ""
                        else:
                            job_condition = 'Queued'
                    else:
                        if v is None:
                            v = ""
                    if k is not None:
                        job[k] = v

                if job_condition and job_condition.upper() != 'Complete':
                    self.jobs.append(job)

                # Check if there is a file path in Feature Layer
                if job.get('feature_layer'):
                    print(f'Feature layer found: {job["feature_layer"]}')
                    feature_layer_path = job['feature_layer']
                    print(f"Processing feature layer: {feature_layer_path}")

                    if feature_layer_path.lower().endswith('.kml'):
                        print(f'Kml found, building AOI from KML')
                        job['feature_layer'] = self.build_aoi_from_kml(feature_layer_path)
                    elif feature_layer_path.lower().endswith('.shp'):
                        if job.get('file_number'):
                            print(f"File number found, running FW setup on shapefile: {feature_layer_path}")
                            new_feature_layer_path = self.build_aoi_from_shp(job, feature_layer_path)
                            job['feature_layer'] = new_feature_layer_path
                        else:
                            print(f'No FW File Number provided for the shapefile, using original shapefile path')
                    else:
                        print(f"Unsupported feature layer format: {feature_layer_path}")

                            
            return self.jobs


    def classify_input_type(self, input):
        print("Classifying input type")
        input_type = None
        file_name, extension = os.path.basename(input).split()

    def start_ast_tb(self, jobs):
        '''Starts an AST toolbox from job params. It will check the capitalization of the True or False inputs and 
        change them to appropriate booleans as the script was failing before implementing this.
        It will also create the output directory if it does not exist based on the job number. Currently this is being created in the T: drive.
        but should be updated once on the server. It checks to make a sure a region has been input on the excel sheet as this is a required parameter.
        It will also catch any errors that are thrown and print them to the console.'''
        try:
            print("Starting AST Toolbox")

            # Loop over the jobs in the spreadsheet
            for job in jobs:
                params = []
                try:
                    # Convert 'true'/'false' strings to booleans (For some reason the script was reading them all as lowercase strings)
                    for param in self.AST_PARAMETERS.values():
                        value = job[param]
                        if isinstance(value, str) and value.lower() in ['true', 'false']:
                            value = True if value.lower() == 'true' else False
                        params.append(value)

                    # Ensure output_directory is set correctly
                    output_directory = job.get('output_directory')

                    # Create a folder path on the T:\ drive (to be changed later) if one doesn't exist
                    if not output_directory:
                        # In case the user didn't fill in an output path on the excel sheet.
                        # Arcpy will throw an error but the folder will still be created and the job still runs
                        
                        #TODO need to update from the T:\ drive to the server
                        job_number = jobs.index(job) + 1
                        output_directory = os.path.join('T:', f'job{job_number}')
                        job['output_directory'] = output_directory

                    # Create the output directory if it does not exist
                    if not os.path.exists(output_directory):
                        try:
                            os.makedirs(output_directory)
                            print(f"Output directory '{output_directory}' created.")
                        except OSError as e:
                            raise RuntimeError(f"Failed to create the output directory. Check your permissions '{output_directory}': {e}")

                    # Ensure that region has been entered otherwise job will fail
                    if not job.get('region'):
                        raise ValueError("Region is required and was not provided.")

                    # Run the tool and send the result to "rslt"
                    print(f"Job Parameters are: {params}")
                    rslt = arcpy.MakeAutomatedStatusSpreadsheet_ast(*params)
                    
                    #TODO change this print to logging
                    print(f"Result: {rslt}")
                except KeyError as e:
                    print(f"Error: Missing parameter in the excel queuefile: {e}")
                except ValueError as e:
                    print(f"Error: {e}")
                except arcpy.ExecuteError as e:
                    print(f"Arcpy error: {arcpy.GetMessages(2)}")
                except Exception as e:
                    print(f"Unexpected error processing job: {e}")

        except ImportError as e:
            print(f"Error importing arcpy toolbox. Check file path in .env file: {e}")
        except arcpy.ExecuteError as e:
            print(f"Arcpy error: {arcpy.GetMessages(2)}")
            
        except Exception as e:
            print(f"Unexpected error: {e}")

    def batch_ast(self):
        ''' Executes the loaded jobs'''
        print("Batching AST")
        counter = 1
        for job in self.jobs:
            self.start_ast_tb([job])
            print(f"Job {counter} Complete")
            counter += 1

    def add_job_result(self, job):
        ''' adds result information to job'''
        # TODO: Create a routine to add status/results to job
        pass

    def create_new_queuefile(self):
        '''write a new queuefile with preset header'''
        print("Creating new queuefile")
        wb = Workbook()
        ws = wb.active
        ws.title = self.XLSX_SHEET_NAME
        headers = list(self.AST_PARAMETERS.values())
        headers.append(self.AST_CONDITION_COLUMN)
        for h in headers:
            c = headers.index(h) + 1
            ws.cell(row=1, column=c).value = h
        wb.save(self.queuefile)

    def build_aoi_from_kml(self, aoi):
        "Write shp file for temporary use"

        # Ensure the KML file exists
        if not os.path.exists(aoi):
            raise FileNotFoundError(f"The KML file '{aoi}' does not exist.")
        
        print("Building AOI from KML")
        from fiona.drvsupport import supported_drivers
        supported_drivers['LIBKML'] = 'rw'
        tmp = os.getenv('TEMP')
        if not tmp:
            raise EnvironmentError("TEMP environment variable is not set.")
        bname = os.path.basename(aoi).split('.')[0]
        fc = bname.replace(' ', '_')
        out_name = os.path.join(tmp, bname + '.gdb')
        if os.path.exists(out_name):
            shutil.rmtree(out_name, ignore_errors=True)
        df = geopandas.read_file(aoi)
        df.to_file(out_name, layer=fc, driver='OpenFileGDB')
        
        #DELETE
        print(f' kml ouput is {out_name} / {fc}')
        return out_name + '/' + fc

    def build_aoi_from_shp(self, job, feature_layer_path):
        """This is snippets of Mike Eastwoods FW Setup Script, if run FW Setup is set to true **Not sure if we need this
        as an option or just make it standard.
        This function will take the raw un-appended shapefile and run it through the FW Setup Script"""

        # Mike Eastwoods FW Setup Script
        print("Processing shapefile using FW Setup Script")
        
        arcpy.env.workspace = r"\\spatialfiles\work\lwbc\nsr\Workarea\fcbc_fsj\Wildlife"
        arcpy.env.overwriteOutput = False

        # Check if there is a file path in Feature Layer
        if feature_layer_path:
            print(f"Processing feature layer: {feature_layer_path}")

        # Check to see if a file number was entered in the excel sheet, if so, use it to name the output directory and authorize the build_aoi_from_shp function to run
        file_number = job.get('file_number')
        
        if not file_number:
            raise ValueError("Error: File Number is required if you are putting in a shapefile that has not been processed in the FW Setup Tool.")
        else:
            print(f"Running FW Setup on File Number: {file_number}")
            
        # Convert file_number to string and make it uppercase
        file_number_str = str(file_number).upper()
        
        # Calculate date variables
        date = datetime.date.today()
        year = str(date.year)

        # Set variables
        base = arcpy.env.workspace
        baseYear = os.path.join(base, year)
        outName = file_number_str
        geometry = "POLYGON"
        template = r"\\spatialfiles.bcgov\Work\lwbc\nsr\Workarea\fcbc_fsj\Templates\BLANK_polygon.shp"
        m = "SAME_AS_TEMPLATE"
        z = "SAME_AS_TEMPLATE"
        spatialReference = arcpy.Describe(template).spatialReference

        # ===========================================================================
        # Create Folders
        # ===========================================================================

        print("Creating FW Setup folders . . .")
        outName = file_number_str
        
        # Create path to folder location
        fileFolder = os.path.join(baseYear, outName)
        shapeFolder = fileFolder
        outPath = shapeFolder
        if os.path.exists(fileFolder):
            print(outName + " folder already exists.")
        else:
            os.mkdir(fileFolder)

        # ===========================================================================
        # Create Shapefile(s) and add them to the current map
        # ===========================================================================

        print("Creating Shapefiles using FW Setup . . .")
        if os.path.isfile(os.path.join(outPath, outName + ".shp")):
            print(os.path.join(outPath, outName + ".shp") + " already exists")
            print("Exiting without creating files")
            return os.path.join(outPath, outName + ".shp")
        else:
            # Creating template shapefile
            create_shp = arcpy.management.CreateFeatureclass(outPath, outName, geometry, template, m, z, spatialReference)
            # Append the newly created shapefile with area of interest
            append_shp = arcpy.management.Append(feature_layer_path, create_shp, "NO_TEST")
            print("Append Successful")
            # Making filename for kml
            create_kml = os.path.join(outPath, outName + ".kml")
            # Make layer for kml to be converted from 
            layer_shp = arcpy.management.MakeFeatureLayer(append_shp, outName)
            # Populate the shapefile                          
            arcpy.conversion.LayerToKML(layer_shp, create_kml)
            # Send message to user that kml has been created
            print("kml created: " + create_kml)

            print(f"FW Setup complete, returned shapefile is {os.path.join(outPath, outName + '.shp')}")
            return os.path.join(outPath, outName + ".shp")


if __name__ == '__main__':
    current_path = os.path.dirname(os.path.realpath(__file__))
      
    # Contains 2 jobs with raw shapefile, with file number
    qf = os.path.join(current_path, '4_jobs_Sunny.xlsx')
    #qf = os.path.join(current_path, '4jobs.xlsx')
    ast = AST_FACTORY(qf, DB_USER, DB_PASS)

    if not os.path.exists(qf):
        print("Queuefile not found, creating new queuefile")
        ast.create_new_queuefile()
    jobs = ast.load_jobs()
    ast.batch_ast()
    ast.start_ast_tb(jobs)

    print("AST Factory Complete")
