# -*- coding: utf-8 -*-

__title__ = "Objekttyp auswählen"

__doc__ = """Version = 1.0
Date    = 02.04.2024
_____________________________________________________________________
Beschreibung:

Skript zum Auswählen des Objekttyps.
_____________________________________________________________________
Anleiting:

-> Klick auf "Objekttyp auswählen" Plug-in Button
-> Wähl ein oder mehrere Objekte aus, die zum gleichen Objekttyp gehören.
-> Klick auf "Fertig stellen".
-> Wähl den Objekttyp.
-> Objekttyp wird Als Wert zugeordnet.
_____________________________________________________________________"""

# IMPORTS
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms

import json
import os
import codecs

# VARIABLES
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

# FUNCTIONS

def get_selected_elements(uidoc):
    """This function will prompt the user to select elements in Revit UI and return their element IDs
    :param uidoc:   uidoc where elements are selected.
    :return:        List of selected element IDs"""
    selected_elements = []
    try:
        reference = uidoc.Selection.PickObjects(ObjectType.Element, "Wählen Sie die Objekte aus, zu denen Allgemeine Attribute hinzugefügt werden sollen.")
        for ref in reference:
            selected_elements.append(ref.ElementId)
    except Exception as ex:
        print("Fehler beim Auswählen von Objekten:", ex)
    return selected_elements


with forms.WarningBar(title='Wähl die Objekte aus, um den Objekttyp zu definieren.:'):
    selected_element_ids = get_selected_elements(uidoc)
    print('Ausgewählte Element-Namen:')
    for element_id in selected_element_ids:
        element = doc.GetElement(element_id)
        print(element.Name)  # Assuming the selected elements have a 'Name' property

# filterXcl = 'Excel workbooks|*.xlsm'

# # relative Path
# path_xcl = forms.pick_file(files_filter=filterXcl, title="Choose excel file")

# # fixed path
# # path_xcl = "C:/Users/Harshal.Gunjal/Desktop/Attributprüfung_13350_2 2.xlsm"

# Define the path to the JSON file
script_directory = os.path.dirname(__file__)
parent_directory = os.path.join(script_directory, '..')
json_folder_path = os.path.join(parent_directory, 'Excel-Datei laden.pushbutton')
json_file_path = os.path.join(json_folder_path, 'file_path.json')

# Read the JSON file
try:
    # Use codecs.open() to open the JSON file with UTF-8 encoding
    with codecs.open(json_file_path, 'r', 'utf-8') as json_file:
        # Load the JSON file contents into a Python object (dictionary)
        data = json.load(json_file)

    # Extract the file path from the JSON data
    windows_path = data["file_path"]

    # Convert the Windows-style file path to a Python-style file path
    python_path = windows_path.replace("\\", "/")

    # Print the converted file path
    print("\n\nExcel File path für Objektliste:")
    print('path_xcl = "{}"'.format(python_path))

except Exception as e:
    # Print the error message if an error occurs
    print("An error occurred: {}".format(e))

from guRoo_xclUtils import *
xcl = xclUtils([], python_path)
all_obj = xcl.xclUtils_import("Objekteliste", 1, 0)

# Extract objects from excel
all_obj_list = []
for row in all_obj[0]:
    all_obj_list.append(row[0])

selected_option = forms.CommandSwitchWindow.show(
    all_obj_list,
     message='Select Option:',
)
print(" ")
print("\n\nAusgewähltes Objekt Name: {}".format(selected_option))

if selected_option:
   t = Transaction(doc, 'Selection of Objekttyp')
   t.Start()
   for element_id in selected_element_ids:
        element = doc.GetElement(element_id)
        for p in element.Parameters:
            # GET PARAMETER
            obj = element.LookupParameter('Objekt')
            # SET PARAMETER
            new_value = selected_option
            obj.Set(new_value)
   t.Commit()


output = script.get_output()
output.log_success("Objekttyp ist ausgewählt")