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
def set_parameter_values(elements):
    for element in elements:
        try:
            t = Transaction(doc, 'Updating Values of Parameters')
            t.Start()

            param = element.LookupParameter(selected_parameter)
            # if param and target_param_name != "Objekt":
            if not param.IsReadOnly:
                param.Set(selected_parameter_value)
                print("The Value '{}' of Parameter '{}' was set on element '{}' with element ID: {}".format(selected_parameter_value, selected_parameter, element.Name, element.Id))
            else:
                print("Error: The parameter '{}' is read-only on element {}".format(selected_parameter, element.Id))

            t.Commit()
        except:
            print("Error processing elements")


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
path_xcl = "N:/B-BI/Team/Gunjal/Pyrevit Plugin/Attributliste_ABS48.xlsm"


# # Define the path to the JSON file
# # Navigate two levels up from the current script's directory
# script_directory = os.path.dirname(__file__)
# parent_directory = os.path.dirname(os.path.dirname(script_directory))
# json_folder_path = os.path.join(parent_directory, 'Excel-Datei laden.pushbutton')
# json_file_path = os.path.join(json_folder_path, 'file_path.json')

# # Read the JSON file
# try:
    # # Use codecs.open() to open the JSON file with UTF-8 encoding
    # with codecs.open(json_file_path, 'r', 'utf-8') as json_file:
        # # Load the JSON file contents into a Python object (dictionary)
        # data = json.load(json_file)

    # # Extract the file path from the JSON data
    # windows_path = data["file_path"]

    # # Convert the Windows-style file path to a Python-style file path
    # python_path = windows_path.replace("\\", "/")

    # # Print the converted file path
    # print("Excel File path:")
    # print('path_xcl = "{}"'.format(python_path))

# except Exception as e:
    # # Print the error message if an error occurs
    # print("An error occurred: {}".format(e))

from guRoo_xclUtils import *
xcl = xclUtils([], path_xcl)
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

# fixed path
path_xcl = "N:/B-BI/Team/Gunjal/Pyrevit Plugin/Attributliste_ABS48.xlsm"

from guRoo_xclUtils import *
xcl = xclUtils([], path_xcl)
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
print("Total number of elements:", total_element_count)


# Process elements in batches
batch_size = 100  # Adjust batch size based on performance
process_elements_in_batches(all_elements, batch_size, set_parameter_values)