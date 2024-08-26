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
from openpyxl import Workbook, load_workbook
from dotenv import load_dotenv
import multiprocessing
import geopandas
import arcpy
import fiona
import datetime
import logging

## *** INPUT YOUR EXCEL FILE NAME HERE ***
excel_file = '2_quick_jobs.xlsx'



###############################################################################################################################################################################
# Set up logging

def setup_logging():
    ''' Set up logging for the script '''
    # Create the log folder filename
    log_folder = f'autoast_logs_{datetime.datetime.now().strftime("%Y%m%d")}'

    # Create the log folder in the current directory if it doesnt exits
    if not os.path.exists(log_folder):
        os.mkdir(log_folder)

    # Create the log file path with the date and time appended
    log_file = os.path.join(log_folder, f'ast_log_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.log')



    # Set up logging config to DEBUG level
    logging.basicConfig(filename=log_file, 
                        level=logging.DEBUG, 
                        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    # Create the logger object and set to the current file name
    logger = logging.getLogger(__name__)

    print("Logging set up")
    logger.info("Logging set up")

    print("Starting Script")
    logger.info("Starting Script")
    
    return logger
###############################################################################################################################################################################




def import_ast():
    # Get the toolbox path from environment variables
    ast_toolbox = os.getenv('TOOLBOX') # File path 

    if ast_toolbox is None:
        print("Unable to find the toolbox. Check the path in .env file")
        logger.error("Unable to find the toolbox. Check the path in .env file")
        exit() 

    # Import the toolbox
    try:
        arcpy.ImportToolbox(ast_toolbox)
        print(f"AST Toolbox imported successfully.")
        logger.info(f"AST Toolbox imported successfully.")
    except Exception as e:
        print(f"Error importing toolbox: {e}")
        logger.error(f"Error importing toolbox: {e}")
        exit()

    # Assign the shapefile template for FW Setup to a variable
    template = os.getenv('TEMPLATE') # File path in .env
    if template is None:
        print("Unable to find the template. Check the path in .env file")
        logger.error("Unable to find the template. Check the path in .env file")
        
    return template
    

###############################################################################################################################################################################
#
# Set up the database connection
#
###############################################################################################################################################################################
def setup_bcgw():
    # Get the secret file containing the database credentials
    SECRET_FILE = os.getenv('SECRET_FILE')

    # If secret file found, load the secret file and display a print message, if not found display an error message
    if SECRET_FILE:
        load_dotenv(SECRET_FILE)
        print(f"Secret file {SECRET_FILE} found")
        logger.info(f"Secret file {SECRET_FILE} found")
    else:
        print("Secret file not found")
        logger.error("Secret file not found")

    # Assign secret file data to variables    
    DB_USER = os.getenv('BCGW_USER')
    DB_PASS = os.getenv('BCGW_PASS')

    # If DB_USER and DB_PASS found display a print message, if not found display an error message
    if DB_USER and DB_PASS:
        print(f"Database user {DB_USER} and password found")
        logger.info(f"Database user {DB_USER} and password found")
    else:
        print("Database user and password not found")
        logger.error("Database user and password not found")

    # Define current path of the executing script
    current_path = os.path.dirname(os.path.realpath(__file__))

    # Create the connection folder
    connection_folder = 'connection'
    connection_folder = os.path.join(current_path, connection_folder)

    # Check for the existance of the connection folder and if it doesn't exist, print an error and create a new connection folder
    if not os.path.exists(connection_folder):
        print("Connection folder not found, creating new connection folder")
        logger.info("Connection folder not found, creating new connection folder")
        os.mkdir(connection_folder)

    # Check for an existing bcgw connection, if there is one, remove it
    if os.path.exists(os.path.join(connection_folder, 'bcgw.sde')):
        os.remove(os.path.join(connection_folder, 'bcgw.sde'))

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
    logger.info("new db connection created")


    arcpy.env.workspace = bcgw_con.getOutput(0)

    print("workspace set to bcgw connection")
    logger.info("workspace set to bcgw connection")
    
    secrets = [DB_USER, DB_PASS]
    
    return secrets
###############################################################################################################################################################################

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
        logger.info("Loading jobs")
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
                        # If the value is not None, assign it to job_condition 
                        if v is not None:
                            job_condition = v
                        #if it is None, assign an empty string to job_condition
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
                    logger.info(f'Feature layer found: {job["feature_layer"]}')
                    feature_layer_path = job['feature_layer']
                    print(f"Processing feature layer: {feature_layer_path}")
                    logger.info(f"Processing feature layer: {feature_layer_path}")

                    if feature_layer_path.lower().endswith('.kml'):
                        print(f'Kml found, building AOI from KML')
                        logger.info(f'Kml found, building AOI from KML')
                        job['feature_layer'] = self.build_aoi_from_kml(feature_layer_path)
                    elif feature_layer_path.lower().endswith('.shp'):
                        if job.get('file_number'):
                            print(f"File number found, running FW setup on shapefile: {feature_layer_path}")
                            logger.info(f"File number found, running FW setup on shapefile: {feature_layer_path}")
                            new_feature_layer_path = self.build_aoi_from_shp(job, feature_layer_path)
                            job['feature_layer'] = new_feature_layer_path
                        else:
                            print(f'No FW File Number provided for the shapefile, loading original shapefile path')
                            logger.info(f'No FW File Number provided for the shapefile, loading original shapefile path')
                    else:
                        print(f"Unsupported feature layer format: {feature_layer_path}")
                        logger.warning(f"Unsupported feature layer format: {feature_layer_path}")

            return self.jobs

    def classify_input_type(self, input):
        print("Classifying input type")
        logger.info("Classifying input type")
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
            logging.info("Starting AST Toolbox")

            # Loop over the jobs in the spreadsheet
            for job in jobs:
                params = []
                
                # Apply a separator line between each job in the log file
                
                logger.info(f"===================================================================")
                logger.info(f"======================= Starting Job #: {job} ======================")
                logger.info(f"====================================================================")
                try:
                    # Convert 'true'/'false' strings to booleans (For some reason the script was reading them all as lowercase strings)
                    for param in self.AST_PARAMETERS.values():
                        value = job[param]
                        if isinstance(value, str) and value.lower() in ['true', 'false']:
                            value = True if value.lower() == 'true' else False
                        params.append(value)

                    # Ensure that region has been entered otherwise job will fail
                    if not job.get('region'):
                        raise ValueError("Region is required and was not provided.")

                    # Run the ast tool 
                    print(f"Job Parameters are: {params}")
                    logger.info(f"Job Parameters are: {params}")
                    arcpy.MakeAutomatedStatusSpreadsheet_ast(*params)
                    
                    self.capture_arcpy_messages()
                    
                    #TODO
                    #Update the ast_condition column in the excel sheet to 'Complete' or Failed
                    # After Jared has completed the function add some sort of job index, 
                    # so it marks the result for each job
                    self.add_job_result(job)
                    
                except KeyError as e:
                    print(f"Error: Missing parameter in the excel queuefile: {e}")
                    logger.error(f"Error: Missing parameter in the excel queuefile: {e}")
                except ValueError as e:
                    print(f"Error: {e}")
                    logger.error(f"Error: {e}")
                except arcpy.ExecuteError as e:
                    print(f"Arcpy error: {arcpy.GetMessages(2)}")
                    logger.error(f"Arcpy error: {arcpy.GetMessages(2)}")
                except Exception as e:
                    print(f"Unexpected error processing job: {e}")
                    logger.error(f"Unexpected error processing job: {e}")

        except ImportError as e:
            print(f"Error importing arcpy toolbox. Check file path in .env file: {e}")
            logger.error(f"Error importing arcpy toolbox. Check file path in .env file: {e}")
        except arcpy.ExecuteError as e:
            print(f"Arcpy error: {arcpy.GetMessages(2)}")
            logger.error(f"Arcpy error: {arcpy.GetMessages(2)}")
        except Exception as e:
            print(f"Unexpected error: {e}")
            logger.error(f"Unexpected error: {e}")

    def batch_ast(self):
        global counter
        ''' Executes the loaded jobs'''
        print("Batching AST")
        logger.info("Batching AST")
        counter = 1
        
        # iterate through the jobs and run the start_ast_tb function on each row of the excel sheet
        for job in self.jobs:
            try:
                self.start_ast_tb([job])
                print(f"Job {counter} Complete")
                logger.info(f"Job {counter} Complete")

            except Exception as e:
                # Log the exception and the job that caused it
                print(f"Error encountered with job {counter}: {e}")
                logger.error(f"Error encountered with job {counter}: {e}")
            finally:
                counter += 1

    def add_job_result(self, job):
        ''' Jared to complete this.
        Function adds result information to job. If the job result is successful, it will update the ast_condition column to "Complete",
        if the job result is failed, it will update the ast_condition column to "Failed" '''
        # TODO: Create a routine to add status/results to job  #Jared
 
        pass
    
    def rerun_failed_jobs(self):
        '''After all jobs have run, this function will scan the excel sheet job result for jobs entered as "failed", 
        if a job is found as failed it will change 'dont_overwrite_outputs' and rerun the job'''
        
        # iterate through the excel sheet of jobs and look for "failed" jobs in the ast_condition column, if the job condition is failed,
        # change the 'dont_overwrite_outputs' to True and rerun the job
        for job in self.jobs:
            # Check the ast_condition column for failed jobs
            if job.get('ast_condition') == 'Failed':
                
                # If the job is failed, change the 'dont_overwrite_outputs' to True and rerun the job
                job['dont_overwrite_outputs'] = True
                
                # Rerun the job
                self.start_ast_tb([job])
                print(f"Job # {counter} rerun")
                logger.info(f"Job # {counter}/ {job} rerun")


    def create_new_queuefile(self):
        '''write a new queuefile with preset header'''
        print("Creating new queuefile")
        logger.info("Creating new queuefile")
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
        logger.info("Building AOI from KML")
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
        logger.info(f' kml ouput is {out_name} / {fc}')
        return out_name + '/' + fc

    def build_aoi_from_shp(self, job, feature_layer_path):
        """This is snippets of Mike Eastwoods FW Setup Script, if run FW Setup is set to true **Not sure if we need this
        as an option or just make it standard.
        This function will take the raw un-appended shapefile and run it through the FW Setup Script"""

        # Mike Eastwoods FW Setup Script
        print("Processing shapefile using FW Setup Script")
        logger.info("Processing shapefile using FW Setup Script")
        
        fsj_workspace = os.getenv('FSJ_WORKSPACE')
        arcpy.env.workspace = fsj_workspace
        arcpy.env.overwriteOutput = False

        # Check if there is a file path in Feature Layer
        if feature_layer_path:
            print(f"Processing feature layer: {feature_layer_path}")
            logger.info(f"Processing feature layer: {feature_layer_path}")

        # Check to see if a file number was entered in the excel sheet, if so, use it to name the output directory and authorize the build_aoi_from_shp function to run
        file_number = job.get('file_number')

        if not file_number:
            raise ValueError("Error: File Number is required if you are putting in a shapefile that has not been processed in the FW Setup Tool.")
        else:
            print(f"Running FW Setup on File Number: {file_number}")
            logger.info(f"Running FW Setup on File Number: {file_number}")

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
    
        m = "SAME_AS_TEMPLATE"
        z = "SAME_AS_TEMPLATE"
        spatialReference = arcpy.Describe(template).spatialReference

        # ===========================================================================
        # Create Folders
        # ===========================================================================

        print("Creating FW Setup folders . . .")
        logger.info("Creating FW Setup folders . . .")
        outName = file_number_str

        # Create path to folder location
        fileFolder = os.path.join(baseYear, outName)
        shapeFolder = fileFolder
        outPath = shapeFolder
        if os.path.exists(fileFolder):
            print(outName + " folder already exists.")
            logger.info(outName + " folder already exists.")
        else:
            os.mkdir(fileFolder)

        # ===========================================================================
        # Create Shapefile(s) and add them to the current map
        # ===========================================================================

        print("Creating Shapefiles using FW Setup . . .")
        logger.info("Creating Shapefiles using FW Setup . . .")
        if os.path.isfile(os.path.join(outPath, outName + ".shp")):
            print(os.path.join(outPath, outName + ".shp") + " already exists")
            logger.info(os.path.join(outPath, outName + ".shp") + " already exists")
            print("Exiting without creating files")
            logger.info("Exiting without creating files")
            return os.path.join(outPath, outName + ".shp")
        else:
            # Creating template shapefile
            create_shp = arcpy.management.CreateFeatureclass(outPath, outName, geometry, template, m, z, spatialReference)
            # Append the newly created shapefile with area of interest
            append_shp = arcpy.management.Append(feature_layer_path, create_shp, "NO_TEST")
            print("Append Successful")
            logger.info("Append Successful")
            # Making filename for kml
            create_kml = os.path.join(outPath, outName + ".kml")
            # Make layer for kml to be converted from 
            layer_shp = arcpy.management.MakeFeatureLayer(append_shp, outName)
            # Populate the shapefile                          
            arcpy.conversion.LayerToKML(layer_shp, create_kml)
            # Send message to user that kml has been created
            print("kml created: " + create_kml)
            logger.info("kml created: " + create_kml)

            print(f"FW Setup complete, returned shapefile is {os.path.join(outPath, outName + '.shp')}")
            logger.info(f"FW Setup complete, returned shapefile is {os.path.join(outPath, outName + '.shp')}")

            return os.path.join(outPath, outName + ".shp")
        

    def capture_arcpy_messages(self):
        ''' Re assigns the arcpy messages  (0 for all messages, 1 for warnings, and 2 for errors) to variables and passes them to the logger'''
        
        arcpy_messages = arcpy.GetMessages(0) # Gets all messages
        arcpy_warnings = arcpy.GetMessages(1) # Gets all warnings only
        arcpy_errors = arcpy.GetMessages(2) # Gets all errors only
        
        if arcpy_messages:
            logger.info(f'ast_toobox arcpy messages: {arcpy_messages}')
        if arcpy_warnings:
            logger.warning(f'ast_toobox arcpy warnings: {arcpy_warnings}')
        if arcpy_errors:
            logger.error(f'ast_toobox arcpy errors: {arcpy_errors}')   

if __name__ == '__main__':
    current_path = os.path.dirname(os.path.realpath(__file__))
    
    # Call the setup_logging function to log the messages
    logger = setup_logging()
    
    # Load the default environment
    load_dotenv()
    
    # Call the import_ast function to import the AST toolbox
    template = import_ast()
    
    # Call the setup_bcgw function to set up the database connection
    secrets = setup_bcgw()
    
    # Create the path for the queuefile
    qf = os.path.join(current_path, excel_file)

    # Create and instance of the Ast Factory class, assign the quefile path and the bcgw username and passwords to the instance
    ast = AST_FACTORY(qf, secrets[0], secrets[1])

    if not os.path.exists(qf):
        print("Queuefile not found, creating new queuefile")
        logger.info("Queuefile not found, creating new queuefile")
        ast.create_new_queuefile()
        
    # load the jobs using the load jobs method. This will scan the excel sheet and assign to "jobs"    
    jobs = ast.load_jobs()
    

    ast.batch_ast()


    print("AST Factory Complete")
    logger.info("AST Factory Complete")
