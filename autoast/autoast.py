# autoast is a script for batch processing the automated status tool
# author: wburt


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

load_dotenv()
SECRET_FILE = os.getenv('SECRET_FILE')
load_dotenv(SECRET_FILE)
DB_USER = os.getenv('BCGW_USER')
DB_PASS = os.getenv('BCGW_PASS')
connection_folder = 'connection'
current_path = os.path.dirname(os.path.realpath(__file__))
if not os.path.exists(connection_folder):
    os.mkdir(connection_folder)
connection_folder= os.path.join(current_path,connection_folder)
if os.path.exists(os.path.join(connection_folder,'bcgw.sde')):
    os.remove(os.path.join(connection_folder,'bcgw.sde'))

bcgw_con = arcpy.management.CreateDatabaseConnection(connection_folder,
                                                    'bcgw.sde',
                                                    'ORACLE',
                                                    'bcgw.bcgov/idwprod1.bcgov',
                                                    'DATABASE_AUTH',
                                                    DB_USER,
                                                    DB_PASS,
                                                    'DO_NOT_SAVE_USERNAME')

#ast_script = r'P:\corp\script_whse\python\Utility_Misc\Ready\statusing_tools_arcpro\scripts\automated_status_sheet_call_routine_arcpro.py'
ast_script = 'automated_status_sheet_call_routine_arcpro.py'
arcpy.env.workspace = bcgw_con.getOutput(0)
class AST_FACTORY:
    ''' AST_FACTORY creates and manages status tool runs '''
    XLSX_SHEET_NAME = 'ast_config'
    AST_PARAMETERS = {
            0:'region',
            1:'feature_layer',
            2:'crown_file_number',
            3:'disposition_number',
            4:'parcel_number',
            5:'output_directory',
            6:'output_directory_same_as_input',
            7:'dont_overwrite_outputs',
            8:'skip_conflicts_and_constraints',
            9:'suppress_map_creation',
            10:'add_maps_to_current',
            11:'run_as_fcbc'
            }
    AST_CONDITION_COLUMN = 'ast_condition'
    AST_SCRIPT = ''

    def __init__(self,queuefile,db_user, db_pass) -> None:
        self.user = db_user
        self.user_cred = db_pass
        self.queuefile = queuefile

    def load_jobs(self):
        '''loads jobs from the jobqueue'''
        self.jobs = []
        assert os.path.exists(self.queuefile)
        if os.path.exists(self.queuefile):
            wb = load_workbook(filename=self.queuefile)
            ws = wb[self.XLSX_SHEET_NAME]
            header = list([row for row in ws.iter_rows(min_row=1, max_col=None,values_only=True)][0])
            data = [row for row in ws.iter_rows(min_row=2, max_col=None,values_only=True)]
            for d in data:
                job = dict()
                for k,v in zip(header,d):
                    if k.lower() ==self.AST_CONDITION_COLUMN.lower():
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
                    job[k]=v
                if job_condition.upper() != 'Complete':
                    self.jobs.append(job)
        
    def start_ast(self,job):
        '''starts a ast process from job params'''
        # TODO: Need a routine to execute and manage ast errors. Ideas:
        #   a. resolve ast toolbox import errors and import toolbox
        #   b. modify ast call routine to allow for os.system calls
        #   c. modify ast call routine to allow for import and exection as a functions

        pass

    def batch_ast(self):
        ''' Executes the loaded jobs'''
        for job in self.jobs:
            self.start_ast(job)
    def add_job_result(self,job):
        ''' adds result information to job'''
        #TODO: Create a routine to add status/results to job
        pass
    def create_new_queuefile(self):
        '''write a new queuefile with preset header'''
        wb = Workbook()
        ws = wb.active
        ws.title = self.XLSX_SHEET_NAME
        headers = list(self.AST_PARAMETERS.values())
        headers.append(self.AST_CONDITION_COLUMN)
        for h in headers:
            c = headers.index(h)+1
            ws.cell(row=1,column=c).value = h
        wb.save(self.queuefile)
    def build_aoi_from_kml(self,aoi):
        "Write shp file for temporary use"
        from fiona.drvsupport import supported_drivers
        supported_drivers['LIBKML'] = 'rw'
        tmp = os.getenv('TEMP')
        bname = os.path.basename(aoi).split('.')[0]
        fc = bname.replace(' ','_')
        out_name = os.path.join(tmp,bname+'.gdb')
        if os.path.exists(out_name):
            shutil.rmtree(out_name,ignore_errors=True)
        df = geopandas.read_file(aoi)
        df.to_file(out_name,layer=fc,driver='OpenFileGDB')
        return out_name + '/' + fc

if __name__=='__main__':
    qf = 'test.xlsx'
    ast = AST_FACTORY(qf,DB_USER,DB_PASS)
    aoi = ast.build_aoi_from_kml('aoi.kml')
    if not os.path.exists(qf):
        ast.create_new_queuefile()
    ast.load_jobs()
    ast.batch_ast()
    

