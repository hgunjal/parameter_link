# -*- coding: utf-8 -*-

__title__ = "Zuordnung Value (X)"

__doc__ = """Version = 1.0
Date    = 16.04.2024
_____________________________________________________________________
Beschreibung:

Skript zum Zuweisung von Werten der weiteren Attribute (Geometrische Eigenschaften, Materialeigenschaften, Bauteileigenschaften).
_____________________________________________________________________
Anleiting:

-> Öffne die erforderliche Ansicht, in der die Werten zugeordnet werden.
-> Klick auf "Zuordnung Value (X)" Plug-in Button
-> Die Werte werden zu den Attributen zugeordnet.
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
from pyrevit import DB, revit
import sys

import json
import os
import codecs

# Function to process elements in batches
def process_elements_in_batches(elements, batch_size, process_function):
    for i in range(0, len(elements), batch_size):
        batch = elements[i:i + batch_size]
        process_function(batch)

# Function to set parameter values for a batch of elements
# def set_parameter_values(elements):
    # for element in elements:
        # try:
            # # Start a transaction to modify the element
            # t = Transaction(doc, 'Adding Platzhalter')
            # t.Start()

            # # Retrieve the Objekt parameter value
            # objekt_typ = element.LookupParameter("Objekt")
            # if objekt_typ:
                # objekt_typ_value = objekt_typ.AsString()

                # # Search for keys in objekt_param_mapping using objekt_typ_value
                # if objekt_typ_value in objekt_param_mapping:
                    # # Retrieve the list of parameters associated with the key
                    # required_params = objekt_param_mapping[objekt_typ_value]

                    # # Set the required parameters as "X" for the element
                    # for target_param_name in required_params:
                        # param = element.LookupParameter(target_param_name)
                        # if param:
                            # if not param.IsReadOnly:
                                # param.Set("X")
                            # # else:
                                # # print("Error: The parameter '{}' is read-only on element {}".format(target_param_name, element.Id))
                
            # # Commit the transaction
            # t.Commit()
        
        # except:
            # print("Error processing element {}".format(element.Id))


def set_parameter_values(elements):
    for element in elements:
        try:
            # Start a transaction to modify the element
            t = Transaction(doc, 'Adding Platzhalter')
            t.Start()

            # Retrieve the Objekt parameter value
            objekt_typ = element.LookupParameter("Objekt")
            if objekt_typ:
                objekt_typ_value = objekt_typ.AsString()
                print("Element ID: {} - Objekt parameter value: {}".format(element.Id, objekt_typ_value))

                # Search for keys in objekt_param_mapping using objekt_typ_value
                if objekt_typ_value in objekt_param_mapping:
                    print("Found matching key in objekt_param_mapping for value: {}".format(objekt_typ_value))
                    
                    # Retrieve the list of parameters associated with the key
                    required_params = objekt_param_mapping[objekt_typ_value]

                    # Set the required parameters for the element
                    for target_param_name in required_params:
                        param = element.LookupParameter(target_param_name)
                        if param:
                            print("Setting parameter '{}' for element {}".format(target_param_name, element.Id))
                            
                            if not param.IsReadOnly:
                                storage_type = param.StorageType
                                
                                # Set the parameter value based on its storage type
                                if storage_type == StorageType.String:
                                    if target_param_name == "TypGelaender":
                                        param.Set("Fuellstabgelaender")
                                    elif target_param_name == "TypUeberbau":
                                        param.Set("Vollrahmen")
                                    elif target_param_name == "TypLaermschutzwandelement":
                                        param.Set("Aluminium")
                                    elif target_param_name == "Herstellungsort":
                                        param.Set("Ortbeton")
                                    elif target_param_name == "TypFundament":
                                        param.Set("Plattenfundament")
                                    elif target_param_name == "TypDichtungsbahn":
                                        param.Set("Bitumen-Dichtungsbahn")
                                    else:
                                        param.Set("X")  # Set default "X" for other target_param_name
                                elif storage_type == StorageType.Integer or storage_type == StorageType.Double:
                                    param.Set(-9999)
                                else:
                                    print("Unhandled storage type for parameter '{}'".format(target_param_name))
                            else:
                                print("Error: The parameter '{}' is read-only on element {}".format(target_param_name, element.Id))
                else:
                    print("No matching key found in objekt_param_mapping for value: {}".format(objekt_typ_value))
            
            # Commit the transaction
            t.Commit()
        
        except Exception as e:
            print("Error processing element {}: {}".format(element.Id, str(e)))



# Get all categories
cats = doc.Settings.Categories

# Check if categories allow bound parameters
allow_bound = [i.AllowsBoundParameters for i in cats]

# Filter categories that allow bound parameters
filtered_categories = [i.Name for i, j in zip(cats, allow_bound) if j is True]

# Sort the categories alphabetically
sorted_categories = sorted(filtered_categories)

# for c in sorted_categories:
    # print(c)

# Create a filtered element collector for the active view
view_collector = FilteredElementCollector(doc, active_view.Id)

# # fixed path
path_xcl = "C:/Users/Harshal.Gunjal/OneDrive - ILF Group Holding GmbH/Desktop/Attributprüfung_13350_2 2.xlsm"

# Define the path to the JSON file
script_directory = os.path.dirname(__file__)
parent_directory = os.path.dirname(os.path.dirname(script_directory))
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
    print("Excel File path:")
    print('path_xcl = "{}"'.format(python_path))

except Exception as e:
    # Print the error message if an error occurs
    print("An error occurred: {}".format(e))

from guRoo_xclUtils import *
xcl = xclUtils([], python_path)

dat = xcl.xclUtils_import("Sheet1", 5, 0) ###CHANGE SHEET NAME HERE###

# Extract data from Excel into lists
targets_params, target_bipgs, fam_inst, fam_wert = [],[],[],[]
for row in dat[0][1:]:
    targets_params.append(row[0])
    target_bipgs.append(row[2])
    fam_inst.append(row[3] == "Ja")
    fam_wert.append(row[4])

dat_1 = xcl.xclUtils_import("andere Eigenschaften", 5, 0) ###CHANGE SHEET NAME HERE###

objekt_name, target_params = [],[]
for row in dat_1[0][1:]:
    objekt_name.append(row[0])
    target_params.append(row[1])
    
# Initialize an empty dictionary to store the mapping
objekt_param_mapping = {}

# Iterate through the lists objekt_name and target_params simultaneously
for objekt, param_name in zip(objekt_name, target_params):
    # Check if the objekt already exists in the dictionary
    if objekt in objekt_param_mapping:
        # If objekt exists, append the param_name to its list of parameters
        objekt_param_mapping[objekt].append(param_name)
    else:
        # If objekt does not exist, create a new entry with the objekt and its corresponding param_name
        objekt_param_mapping[objekt] = [param_name]

# Print the dictionary
print("Objekt Parameter Mapping:")
for objekt, params in objekt_param_mapping.items():
    print("{0}: {1}".format(objekt, params))

objekt_list = list(objekt_param_mapping.keys())

# Get categories present in the active view
categories_in_view = set()
for element in view_collector:
    if element.Category is not None:
        categories_in_view.add(element.Category.Name)

# Filter categories from filtered_categories that are present in the active view
categories_in_active_view = [category for category in filtered_categories if category in categories_in_view]

# Create a list to store elements from all categories
all_elements = []

# Iterate over each category name in 'categories_in_active_view'
for category_name in categories_in_active_view:
    category = doc.Settings.Categories.get_Item(category_name)
    if category:
        category_filter = ElementCategoryFilter(category.Id)
        category_elements = FilteredElementCollector(doc, active_view.Id).WherePasses(category_filter).ToElements()
        all_elements.extend(category_elements)

# Print total number of elements
total_element_count = len(all_elements)
print("Total number of elements:", total_element_count)


# Process elements in batches
batch_size = 10  # Adjust batch size based on performance
process_elements_in_batches(all_elements, batch_size, set_parameter_values)

print("Added 'X' as Placeholder")
