# -*- coding: utf-8 -*-

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
def set_parameter_values(elements, selected_parameter, selected_parameter_value):
    for element in elements:
        try:
            t = Transaction(doc, 'Adding Value to AA')
            t.Start()

            for target_param_name in targets_params:
                param = element.LookupParameter(target_param_name)
                if param:
                    if not param.IsReadOnly:
                        if target_param_name == selected_parameter:
                            param.Set(selected_parameter_value)
                    else:
                        print("Error: The parameter '{}' is read-only on element {}".format(target_param_name, element.Id))
                else:
                    print("Parameter '{}' not found on element {}".format(target_param_name, element.Id))

            t.Commit()
        except Exception as e:
            print("Error processing elements:", str(e))

def get_selected_elements(uidoc):
    """This function will prompt the user to select elements in Revit UI and return their element IDs
    :param uidoc:   uidoc where elements are selected.
    :return:        List of selected element IDs"""
    selected_elements = []
    try:
        reference = uidoc.Selection.PickObjects(ObjectType.Element, "Wähl die Objekte aus, um die Werten festzulegen.")
        for ref in reference:
            selected_elements.append(ref.ElementId)
    except Exception as ex:
        print("Fehler beim Auswählen von Objekten:", ex)
    return selected_elements


with forms.WarningBar(title='Wähl die Objekte aus, um die Werten festzulegen.'):
    selected_element_ids = get_selected_elements(uidoc)
    print('Ausgewählte Element-Namen:')
    for element_id in selected_element_ids:
        element = doc.GetElement(element_id)
        print(element.Name)  # Assuming the selected elements have a 'Name' property


# Assuming 'categories_in_active_view' contains the category names present in the active view

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
# path_xcl = "C:/Users/Harshal.Gunjal/Desktop/Attributprüfung_13350_2 2.xlsm"

# Define the path to the JSON file
# Navigate two levels up from the current script's directory
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
dat = xcl.xclUtils_import("Allgemeine Attribute", 5, 0) ###CHANGE SHEET NAME HERE###

# Extract data from Excel into lists
targets_params, target_bipgs, par_inst, par_formulae = [],[],[],[]
for row in dat[0][1:]:
    targets_params.append(row[0])

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
# print("Total number of elements:", total_element_count)

selected_parameter = forms.ask_for_one_item(
    targets_params,
    default='4D-Vorgangs-ID',
    prompt='Bitte wähl ein Attribut, dessen Wert du festlegen möchtest',
    title='Allgemeine Attribute auswählen'
)

selected_parameter_value = forms.ask_for_string(
    default='z.B. DB InfraGO',
    prompt='Bitte gib den Wert ein',
    title='Zuweisung von Werten'
)

# Process elements in batches
batch_size = 100  # Adjust batch size based on performance

if selected_parameter_value:
    for element_id in selected_element_ids:
        try:
            t = Transaction(doc, 'Adding Value to AA')
            t.Start()

            element = doc.GetElement(element_id)
            for p in element.Parameters:
                # GET PARAMETER
                obj = element.LookupParameter(selected_parameter)
                # SET PARAMETER
                new_value = selected_parameter_value
                obj.Set(new_value)

            t.Commit()
        except Exception as e:
            print("Error processing parameter:", str(e))


# print("Attributes Updated")
