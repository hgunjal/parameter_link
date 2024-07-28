# -*- coding: utf-8 -*-

__title__ = "Excel-Datei laden"

__doc__ = """Version = 1.0
Date    = 22.04.2024
_____________________________________________________________________
Beschreibung:

Skript zum Speichern des Excel-Dateipfads in einer .JSON-Datei
_____________________________________________________________________"""

# IMPORTS
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms
from Autodesk.Revit.Exceptions import InvalidOperationException

# VARIABLES
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
active_view = doc.ActiveView

from pyrevit import HOST_APP
from pyrevit import DB, revit, script
import sys
import json
import os
import codecs

# Choose the Excel file
filterXcl = 'Excel workbooks|*.xlsm'

# relative Path
path_xcl = forms.pick_file(files_filter=filterXcl, title="Choose excel file")

# Check if a file is selected
if path_xcl:
    # Determine the directory of the current script (plugin path)
    script_path = os.path.dirname(__file__)

    # Define the path for the JSON file in the same directory as the script
    json_file_path = os.path.join(script_path, "file_path.json")

    # Create a dictionary to store the file path
    file_info = {"file_path": path_xcl}

    try:
        # Use codecs.open() to open the JSON file with UTF-8 encoding
        with codecs.open(json_file_path, 'w', 'utf-8') as json_file:
            json.dump(file_info, json_file, ensure_ascii=False)

        # Use .format() to print the success message
        print("File path saved to {}".format(json_file_path))

    except Exception as e:
        # Use .format() to print the error message
        print("An error occurred: {}".format(e))

else:
    # Use .format() for the message indicating no file was selected
    print("No file selected.")

sPFile = app.OpenSharedParameterFile()
if sPFile is None:
    output = script.get_output()
    output.log_warning("No Shared Parameter File found! Please add a Shared Parameter File.")
    raise Exception(
        "No Shared Parameter File found! Please add a Shared Parameter File.")  # Raise an exception if no shared parameter file is found