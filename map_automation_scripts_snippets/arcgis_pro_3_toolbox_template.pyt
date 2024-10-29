# ArcGIS Pro 3 Toolbox Template

#===========================================================================
# Script name: Arcgis Pro 3 Toolbox Template
# Author: https://pro.arcgis.com/en/pro-app/latest/arcpy/geoprocessing_and_python/a-template-for-python-toolboxes.htm

# Created on: 10/21/2024

#New sample changes
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
        self.tools = [NameOfYourTool, MySecondTool]  
        # Insert the name of each tool in your toolbox if you have more than one. 
        # i.e. self.tools = [FullSiteOverviewMaps, ExportSiteAndImageryLayout, Amendment]


class NameOfYourTool(object):
    def __init__(self):
        """Define the tool (tool name is the name of the class)."""
        self.label = "Tool"
        self.description = ""

    def getParameterInfo(self):
        """This function assigns parameter information for tool""" 
        
        ef setup_logging():
    ''' Set up logging for the script '''
    # Create the log folder filename
    log_folder = f'autoast_logs_{datetime.datetime.now().strftime("%Y%m%d")}'

    # Create the log folder in the current directory if it doesn't exits
    if not os.path.exists(log_folder):
        os.mkdir(log_folder)
    
    # Check if the log folder was created successfully
    assert os.path.exists(log_folder), "Error creating log folder, check permissions and path"

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



def import_ast(logger):
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
def setup_bcgw(logger):
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

    }
    
    ADDITIONAL_PARAMETERS = {
        12: 'ast_condition',
        13: 'file_number'
    }
    
    AST_CONDITION_COLUMN = 'ast_condition'
    DONT_OVERWRITE_OUTPUTS = 'dont_overwrite_outputs'
    AST_SCRIPT = ''
    job_index = None  # Initialize job_index as a global variable
    
    def __init__(self, queuefile, db_user, db_pass, logger=None, current_path=None) -> None:
            self.user = db_user
            self.user_cred = db_pass
            self.queuefile = queuefile
            self.jobs = []
            self.logger = logger or logging.getLogger(__name__)
            self.current_path = current_path  

    def load_jobs(self):
        '''
        load jobs will check for the existence of the queuefile, if it exists it will load the jobs from the queuefile. Checking if they 
        are Complete and if not, it will add them to the jobs  as Queued
        '''

        #global job_index
        self.logger.info("##########################################################################################################################")
        self.logger.info("#")
        self.logger.info("Loading Jobs...")
        self.logger.info("#")
        self.logger.info("##########################################################################################################################")

        # Initialize the jobs list to store jobs
        self.jobs = []

        #This parameter is the file number of the application
        file_num = arcpy.Parameter(
            displayName = "Lands File Number",
            name="file_num",
            datatype="String",
            parameterType="Required",
            direction="Input")
        
        
        # This parameter is the proponent name for the cutting permit
        proponent_name = arcpy.Parameter(
            displayName="Proponent Name",
            name="Proponent Name",
            datatype="String",
            parameterType="Optional",
            direction="Input"
            )

        # This parameter is the cutting permit or TSL number
        cp_ID = arcpy.Parameter(
            displayName="Cuttng Permit/TSL",
            name="Cutting Permit",
            datatype="String",
            parameterType="Optional",
            direction="Input"
            ) 

        parameters = [file_num, proponent_name, cp_ID]  # Each parameter name needs to be in here, separated by a comma

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
        
        """The source code of the tool."""
        
        
        # Assign your parameters to variables
        file_num = parameters[0].valueAsText  # Take the first parameter **0 indexed!!** and assign in the the variable file_num
        proponent_name = parameters[1].valueAsText
        cp_ID = parameters[2].valueAsText
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
