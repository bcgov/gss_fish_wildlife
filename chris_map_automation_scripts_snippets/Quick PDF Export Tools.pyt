 #Author: Chris Sostad
# chris.sostad@gov.bc.ca
# Ministry of WLRS
# Created Date: February 21st, 2024
# Updated Date: August 23, 2024
# Description:
#   This script is a tool box for exporting either a sinlge layout to the desired directory or exporting a single layout from a list of layouts to the desired directory.

# --------------------------------------------------------------------------------

#import all necessary libraries 
import arcpy 
import traceback 
import sys 

############################################################################################################################################################################################
#
#
# Global Functions used in all of the Export Scripts
#
#
############################################################################################################################################################################################

# Global exportToPdf function
def exportToPdf(layout, workSpace_path, pdf_file_name):
    out_pdf = f"{workSpace_path}\\{pdf_file_name}"
    layout.exportToPDF(
        out_pdf=out_pdf,
        resolution=300,  # DPI
        image_quality="BETTER",
        jpeg_compression_quality=80  # Quality (0 to 100)
    )



class Toolbox(object):
    def __init__(self):
        """Define the toolbox (name of toolbox is name of the file)"""
        self.label = "Toolbox"
        self.alias = ""

        #List of tool classes associated with this toolbox
        self.tools = [ExportSingleLayout, FromMultipleExportSingleLayout]




class ExportSingleLayout(object):
    def __init__(self):
        """This tool will export a selected layout to a chosen directory."""
        self.label = "Export Layout From Project With Only One Layout"
        self.description = "Export a selected layout from your project to a PDF."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """This function assigns parameter information for tool"""

        # Parameter for the directory where the PDF will be stored
        workSpace = arcpy.Parameter(
            displayName="Navigate to the folder where you want to save your pdf (Warning: Overwrite set to true!)",
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )

        # Parameter for the name of the exported PDF
        pdfName = arcpy.Parameter(
            displayName="File name you want for your pdf",
            name="pdfName",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )

        # List of parameters
        parameters = [workSpace, pdfName]

        return parameters

    def updateParameters(self, parameters):
        """This method is called whenever a parameter has been changed."""
        if not parameters[1].altered:
            aprx = arcpy.mp.ArcGISProject("CURRENT")
            layouts = aprx.listLayouts()
            if len(layouts) == 1:
                parameters[1].value = f"{layouts[0].name}.pdf"
        return

    def execute(self, parameters, messages):
        try:
            workSpace_path = parameters[0].valueAsText
            pdf_file_name = parameters[1].valueAsText

            # Set the project variable
            aprx = arcpy.mp.ArcGISProject("CURRENT")

            # Check if '.pdf' extension is already included in the name, if not add it
            if not pdf_file_name.endswith('.pdf'):
                pdf_file_name += '.pdf'

            # Access the selected layout in the project
            layout = aprx.listLayouts()[0]

            # Set overwrite to True
            arcpy.env.overwriteOutput = True

            # Export the layout to PDF
            layout.exportToPDF(f"{workSpace_path}\\{pdf_file_name}")

        except arcpy.ExecuteError:
            msgs = arcpy.GetMessages(2)
            arcpy.AddError(msgs)
        except:
            tb = sys.exc_info()[2]
            tbinfo = traceback.format_tb(tb)[0]
            pymsg = ("PYTHON ERRORS:\nTraceback info:\n" + tbinfo +
                     "\nError Info:\n" + str(sys.exc_info()[1]))
            msgs = "ArcPy ERRORS:\n" + arcpy.GetMessages(2) + "\n"
            arcpy.AddError(pymsg)
            arcpy.AddError(msgs)

        
class FromMultipleExportSingleLayout(object):
    def __init__(self):
        """This tool will allow you to choose from multiple layouts to export a selected layout to a chosen directory."""
        self.label = "From Multiple Layouts Export Single Layout"
        self.description = "Choose from multiple layouts and export a selected layout from your project to a PDF."
        self.canRunInBackground = False

    def getParameterInfo(self):
        """This function assigns parameter information for the tool."""

        # Parameter for selecting the layout
        layoutList = arcpy.Parameter(
            displayName="Select a layout from a list of layouts",
            name="layoutList",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        layoutList.filter.type = "ValueList"

        # Initialize the layout list
        aprx = arcpy.mp.ArcGISProject("CURRENT")
        layouts = aprx.listLayouts()
        layoutList.filter.list = [layout.name for layout in layouts]

        # Parameter for the name of the exported PDF
        pdfName = arcpy.Parameter(
            displayName="File name you want for your pdf",
            name="pdfName",
            datatype="GPString",
            parameterType="Required",
            direction="Input"
        )
        pdfName.value = "ChangeMe.pdf"

        # Parameter for the directory where the PDF will be stored
        workSpace = arcpy.Parameter(
            displayName="Navigate to the folder where you want to save your pdf (Warning: Overwrite set to true!)",
            name="workSpace",
            datatype="DEWorkspace",
            parameterType="Required",
            direction="Input"
        )
        workSpace.filter.list = ["Local Database", "File System"]

        # List of parameters in the desired order
        parameters = [layoutList, pdfName, workSpace]

        return parameters

    def updateParameters(self, parameters):
        """This method is called whenever a parameter has been changed."""
        # Update pdfName to match the selected layout name
        if parameters[0].altered:  # Check if the layout has been altered
            parameters[1].value = parameters[0].value + ".pdf"

        return

    def execute(self, parameters, messages):
        try:
            workSpace_path = parameters[2].valueAsText
            pdf_file_name = parameters[1].valueAsText
            selected_layout_name = parameters[0].valueAsText

            # Check if '.pdf' extension is already included in the name, if not add it
            if not pdf_file_name.endswith('.pdf'):
                pdf_file_name += '.pdf'

            # Set the path to the current project
            aprx = arcpy.mp.ArcGISProject("CURRENT")

            # Access the selected layout in the project
            layout = aprx.listLayouts(selected_layout_name)[0]
            
            # Set overwrite to True
            arcpy.env.overwriteOutput = True
            
            # Export the layout to a PDF with the specified settings
            layout.exportToPDF(
                out_pdf=workSpace_path + "\\" + pdf_file_name,
                resolution=300,  # DPI
                image_quality="BETTER",
                jpeg_compression_quality=80  # Quality (0 to 100)
            )

            arcpy.AddMessage(f"Layout '{selected_layout_name}' exported to {workSpace_path}")

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
            #



