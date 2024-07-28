# -*- coding: utf-8 -*-

__title__ = "Hinzufügung"

__doc__ = """Version = 1.0
Date    = 15.04.2024
_____________________________________________________________________
Beschreibung:

Skript zum Hinzufügen von Projektparametern (Geometrische Eigenschaften, Materialeigenschaften, Bauteileigenschaften) zu Revit-Elementen basierend auf Daten aus einer Excel-Datei.
_____________________________________________________________________
Anleitung:

-> Hinzufügen der "Shared Parameter" .txt Datei zum Revit-Modell.
-> Öffne die Ansicht, in der Parameter hinzufügen werden.
-> Klick auf "Hinzufügung" Plug-in Button
-> Shared Parameters als Projektparameter erstellen.

Hinweis:
Für eine Instanz werden viele Projektparameter hinzugefügt, da beim Hinzufügen von Projektparametern eine Bindung an eine Kategorie und nicht an eine einzelne Instanz erfolgt.
_____________________________________________________________________
"""

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

# FUNCTIONS

def insert_shared_parameter(app, name, element_cat, group, inst):
    """Inserts a shared parameter into the project.
    
    Args:
        app (Application): The Revit application instance.
        name (str): The name of the shared parameter.
        element_cat (CategorySet): The categories to which the parameter should be added.
        group (BuiltInParameterGroup): The parameter group.
        inst (bool): Whether the parameter should be instance-based.
    """
    # Open the shared parameter file associated with the application
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception("No Shared Parameter File found! Please add a Shared Parameter File.")  # Raise an exception if no shared parameter file is found

    # Find the definition of the shared parameter with the specified name
    definitions = [d for dg in def_file.Groups 
                   for d in dg.Definitions 
                   if d.Name == name]
    if not definitions:
        return  # Skip the rest of the function and return to the caller
        # raise Exception("Invalid Name Input!")  # Raise an exception if the specified parameter name is not found

    # Get the first definition (assuming there's only one parameter with the specified name)
    def_param = definitions[0]
    

    # Create a binding for either type or instance based on the 'inst' parameter
    binding = app.Create.NewTypeBinding(element_cat)
    if inst:
        # print("Selected object is Instance")
        binding = InstanceBinding(element_cat)  # New Binding
    
    t = Transaction(doc, "Attach SP as PP")
    t.Start()
    
    # Get the parameter bindings of the active document
    map = doc.ParameterBindings
    
    # Insert the shared parameter definition and its binding into the parameter bindings of the document
    map.Insert(def_param, binding, group)
    
    t.Commit()
    
    # Print the name of the parameter added


def reinsert_shared_parameter(app, name, element_cat, group, inst):
    """Reinserts a shared parameter into the project.
    
    Args:
        app (Application): The Revit application instance.
        name (str): The name of the shared parameter.
        element_cat (CategorySet): The categories to which the parameter should be added.
        group (BuiltInParameterGroup): The parameter group.
        inst (bool): Whether the parameter should be instance-based.
    """
    # Open the shared parameter file associated with the application
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception("No Shared Parameter File found! Please add a Shared Parameter File.")  # Raise an exception if no shared parameter file is found

    # Find the definition of the shared parameter with the specified name
    definitions = [d for dg in def_file.Groups 
                   for d in dg.Definitions 
                   if d.Name == name]
    if not definitions:
        raise Exception("Invalid Name Input!")  # Raise an exception if the specified parameter name is not found

    # Get the first definition (assuming there's only one parameter with the specified name)
    def_param = definitions[0]
    

    # Create a binding for either type or instance based on the 'inst' parameter
    binding = app.Create.NewTypeBinding(element_cat)
    if inst:
        # print("Selected object is Instance")
        binding = InstanceBinding(element_cat)
    
    t = Transaction(doc, "Attach SP as PP")
    t.Start()
    
    # Get the parameter bindings of the active document
    map = doc.ParameterBindings
    
    # Insert the shared parameter definition and its binding into the parameter bindings of the document
    map.ReInsert(def_param, binding, group)
    
    t.Commit()
    

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

path_xcl = "C:/Users/Harshal.Gunjal/OneDrive - ILF Group Holding GmbH/Desktop/Attributprüfung_13350_2 2.xlsm"

    
# Import Excel data using custom utility
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

# # Print the dictionary
# print("Objekt Parameter Mapping:")
# for objekt, params in objekt_param_mapping.items():
    # print("{0}: {1}".format(objekt, params))

objekt_list = list(objekt_param_mapping.keys())
# print(objekt_list)
    
# Parameters to be added
req_param = targets_params

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

# Get categories present in the active view
categories_in_view = set()
for element in view_collector:
    if element.Category is not None:
        categories_in_view.add(element.Category.Name)

# Filter categories from filtered_categories that are present in the active view
categories_in_active_view = [category for category in filtered_categories if category in categories_in_view]

# Print categories in the active view
for category in categories_in_active_view:
    print(category)

# Create a list to store elements from all categories
all_elements = []

for category_name in categories_in_active_view:
    try:
        # Get the category by name
        category = doc.Settings.Categories.get_Item(category_name)
        if category:
            # Create a filter for the category
            category_filter = ElementCategoryFilter(category.Id)
            # Collect elements of the category in the active view
            category_elements = FilteredElementCollector(doc, active_view.Id).WherePasses(category_filter).ToElements()
            # Extend the list with collected elements
            all_elements.extend(category_elements)
    except Exception as e:
        # Handle exception and skip the category
        print("Skipping category '{}' due to exception: {}".format(category_name, e))

target_param_name = "Objekt"

# for e in all_elements:
    # param = e.LookupParameter(target_param_name)
    # print("The parameter '{}' has value as '{}' on element {}".format(target_param_name, param.AsString(), e.Id))
    
    
# # Initialize an empty dictionary to store the key-value pairs
# param_to_category_dict = {}
#
# # Iterate through the list of all elements
# for e in all_elements:
#     # Look up the specified parameter in the element
#     param = e.LookupParameter(target_param_name)
#
#     # Check if the parameter exists and is not None
#     if param:
#         # Get the parameter value as a string
#         param_value = param.AsString()
#
#         # Get the category name of the element
#         category_name = e.Category.Name
#
#         # Use the parameter value as the key and the category name as the value
#         param_to_category_dict[param_value] = category_name

# Initialize an empty dictionary to store parameter values and corresponding categories
param_to_category_dict = {}

# Iterate through the list of all elements
for e in all_elements:
    # Look up the specified parameter in the element
    param = e.LookupParameter(target_param_name)

    # Check if the parameter exists and is not None
    if param:
        # Get the parameter value as a string
        param_value = param.AsString()

        # Get the category name of the element
        category_name = e.Category.Name

        # Check if the parameter value already exists in the dictionary
        if param_value in param_to_category_dict:
            # Add the category name to the set of categories for this parameter value
            param_to_category_dict[param_value].add(category_name)
        else:
            # Create a new entry with the parameter value as the key and a set with the category name as the value
            param_to_category_dict[param_value] = {category_name}

# Convert sets back to lists and print the result using .format
for param_value, categories in param_to_category_dict.items():
    print("{} = {}".format(param_value, list(categories)))

# Print the dictionary to see the key-value pairs
print("[object_name] = category_name")
print(param_to_category_dict)

# Convert sets back to lists for further processing
param_to_category_dict = {k: list(v) for k, v in param_to_category_dict.items()}

# Initialize an empty list to store the filtered keys
filtered_objects = []

# Iterate through the key-value pairs in the dictionary
for key, value in param_to_category_dict.items():
    # Check if the value is not equal to "X"
    if key != "X":
        # Add the key to the filtered list
        filtered_objects.append(key)

# # Print the list of filtered keys
# print("List of Objects (excluding 'X'):")
# print(filtered_objects)

# Initialize an empty dictionary to store parameter values and corresponding categories
param_to_category_dict = {}

# Iterate through the list of all elements
for e in all_elements:
    # Look up the specified parameter in the element
    param = e.LookupParameter(target_param_name)

    # Check if the parameter exists and is not None
    if param:
        # Get the parameter value as a string
        param_value = param.AsString()

        # Get the category name of the element
        category_name = e.Category.Name

        # Check if the parameter value already exists in the dictionary
        if param_value in param_to_category_dict:
            # Add the category name to the set of categories for this parameter value
            param_to_category_dict[param_value].add(category_name)
        else:
            # Create a new entry with the parameter value as the key and a set with the category name as the value
            param_to_category_dict[param_value] = {category_name}

# Convert sets back to lists for further processing
param_to_category_dict = {k: list(v) for k, v in param_to_category_dict.items()}

# Initialize an empty list to store the filtered keys
filtered_objects = []

# Iterate through the key-value pairs in the dictionary
for key, value in param_to_category_dict.items():
    # Check if the key is not equal to "X"
    if key != "X":
        # Add the key to the filtered list
        filtered_objects.append(key)

# Print the list of filtered keys
print("List of Objects (excluding 'X'):")
print(filtered_objects)

# Initialize an empty dictionary to store the final mappings
category_to_objekt_params_dict = {}

# Iterate through the list of filtered objects
for obj in filtered_objects:
    # Check if the object name exists as a key in the objekt_param_mapping dictionary
    if obj in objekt_param_mapping:
        # Get the list of parameters associated with the object
        params = objekt_param_mapping[obj]

        # Check if the object name exists as a key in the param_to_category_dict
        if obj in param_to_category_dict:
            # Get the category names of the object
            categories = param_to_category_dict[obj]

            # Iterate through the categories and update the final dictionary
            for category in categories:
                # Check if the category already exists in the final dictionary
                if category in category_to_objekt_params_dict:
                    # If category exists, extend the list of parameters
                    category_to_objekt_params_dict[category].extend(params)
                else:
                    # If category does not exist, create a new entry
                    category_to_objekt_params_dict[category] = params

# Print the final dictionary to see the mappings
print("Category to Objekt Parameters Mapping:")
for category, params in category_to_objekt_params_dict.items():
    print("{0}: {1}".format(category, params))

# Initialize an empty dictionary to store the mapping from parameters to categories
param_to_categories_dict = {}

# Iterate through the final dictionary category_to_objekt_params_dict
for category, params in category_to_objekt_params_dict.items():
    # Iterate through the list of parameters for the current category
    for param in params:
        # Check if the parameter already exists in the dictionary
        if param not in param_to_categories_dict:
            # If the parameter does not exist, create a new entry with the parameter as the key
            # and a list containing the current category as the value
            param_to_categories_dict[param] = [category]
        else:
            # If the parameter already exists, add the current category to its list of categories
            param_to_categories_dict[param].append(category)

# Print the new dictionary to see the mappings from parameters to categories
print("Parameter to Categories Mapping:")
for param, categories in param_to_categories_dict.items():
    print("{0}: {1}".format(param, categories))



# Iterate through the dictionary param_to_categories_dict
for param, categories in param_to_categories_dict.items():
    # Create a new CategorySet in each iteration
    cats1 = app.Create.NewCategorySet()
    
    # Iterate through the list of categories associated with the current parameter
    for category_name in categories:
        print("Category: {}".format(category_name))
        
        # Retrieve the category from the document settings
        category = doc.Settings.Categories.get_Item(category_name)
        
        # Perform the operations with the category and parameter list
        if category is not None:
            # Insert the category into the CategorySet
            cats1.Insert(category)
    
    # Now that the CategorySet has been populated, insert the shared parameter
    print("Inserting parameter: {}".format(param))
    insert_shared_parameter(app, param, cats1, BuiltInParameterGroup.PG_DATA, True)
    
    cats1.Clear()
    # The CategorySet will be cleared when a new set is created in each iteration
    print("*" * 50)
