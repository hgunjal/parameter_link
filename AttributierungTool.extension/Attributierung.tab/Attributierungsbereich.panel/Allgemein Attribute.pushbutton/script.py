# -*- coding: utf-8 -*-
__title__ = "Allgemein Attribute"

__doc__ = """Version = 1.0
Date    = 02.04.2024
_____________________________________________________________________
Beschreibung:

Skript zum Hinzufügen von Projektparametern (Allgemeine Attribute) zu Revit-Elementen basierend auf Daten aus einer Excel-Datei.
_____________________________________________________________________
Anleitung:

-> Hinzufügen der "Shared Parameter" .txt Datei zum Revit-Modell.
-> Öffne die Ansicht, in der Parameter hinzufügen werden.
-> Klick auf "Allgemein Attribute" Plug-in Button
-> Shared Parameters als Projektparameter erstellen.
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
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception("No Shared Parameter File found! Please add a Shared Parameter File.")
    
    definitions = [d for dg in def_file.Groups 
                   for d in dg.Definitions 
                   if d.Name == name]

    if not definitions:
        raise Exception("Parameter '{}' not found in the Shared Parameter File!".format(name))

    def_param = definitions[0]

    binding = app.Create.NewTypeBinding(element_cat) if not inst else InstanceBinding(element_cat)
    
    t = Transaction(doc, "Attach SP as PP")
    t.Start()
    map = doc.ParameterBindings
    map.Insert(def_param, binding, group)
    t.Commit()

def reinsert_shared_parameter(app, name, element_cat, group, inst):
    """Reinserts a shared parameter into the project.
    
    Args:
        app (Application): The Revit application instance.
        name (str): The name of the shared parameter.
        element_cat (CategorySet): The categories to which the parameter should be added.
        group (BuiltInParameterGroup): The parameter group.
        inst (bool): Whether the parameter should be instance-based.
    """
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception("No Shared Parameter File found! Please add a Shared Parameter File.")
    
    definitions = [d for dg in def_file.Groups 
                   for d in dg.Definitions 
                   if d.Name == name]

    if not definitions:
        raise Exception("Parameter '{}' not found in the Shared Parameter File!".format(name))

    def_param = definitions[0]

    binding = app.Create.NewTypeBinding(element_cat) if not inst else InstanceBinding(element_cat)
    
    t = Transaction(doc, "Attach SP as PP")
    t.Start()
    map = doc.ParameterBindings
    map.ReInsert(def_param, binding, group)
    t.Commit()

def check_loaded_params(list_p_names):
    """Checks if parameters from a provided list are missing in the project.
    
    Args:
        list_p_names (list): List of parameter names.
        
    Returns:
        list: List of missing parameters.
    """
    bm = doc.ParameterBindings
    itor = bm.ForwardIterator()
    itor.Reset()
    loaded_parameters = []
    while itor.MoveNext():
        d = itor.Key
        loaded_parameters.append(d.Name)
    missing_params = [p_name for p_name in list_p_names if p_name not in loaded_parameters]
    if not missing_params:
        print("No missing parameters in Project")
    else:
        print("Missing parameters in Project:", missing_params)
    return missing_params

def check_loaded_params_in_category(list_p_names, category_set):
    """Checks if parameters from a provided list are missing in a specified category.
    
    Args:
        list_p_names (list): List of parameter names.
        category_set (CategorySet): Set of categories to check against.
        
    Returns:
        list: List of missing parameters.
    """
    # Access the parameter bindings of the document
    bm = doc.ParameterBindings
    itor = bm.ForwardIterator()
    itor.Reset()
    # Initialize an empty list to store the names of loaded parameters
    loaded_parameters = []
    # Iterate through the parameter bindings
    while itor.MoveNext():
        # Get the parameter definition from the binding
        parameter_definition = itor.Key
        b = bm[parameter_definition]
        # Check if the parameter is applicable to any of the specified categories
        for cat in b.Categories:
            if cat in category_set:
                # Append the name of the parameter to the loaded_parameters list
                loaded_parameters.append(parameter_definition.Name)
                break  # Break the loop as we found the category
    # Create a list of missing parameters by comparing the provided list with loaded_parameters
    missing_params = [p_name for p_name in list_p_names if p_name not in loaded_parameters]
    if not missing_params:
        print("No missing parameters in the specified category")
    else:
        print("Missing parameters in the specified category:", missing_params)
    # Return the list of missing parameters
    return missing_params

def get_selected_elements(uidoc):
    """Prompts the user to select elements in Revit UI and returns their element IDs.
    
    Args:
        uidoc (UIDocument): The active UIDocument.
        
    Returns:
        list: List of selected element IDs.
    """
    selected_elements = []
    try:
        reference = uidoc.Selection.PickObjects(ObjectType.Element, "Wähl die Objekte aus, zu denen Allgemeine Attribute hinzugefügt werden sollen.")
        for ref in reference:
            selected_elements.append(ref.ElementId)
    except Exception as ex:
        print("Fehler beim Auswählen von Objekten:", ex)
    return selected_elements

# MAIN

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
    print("Excel File path:")
    print('path_xcl = "{}"'.format(python_path))

except Exception as e:
    # Print the error message if an error occurs
    print("An error occurred: {}".format(e))

# # Choose the Excel file
# filterXcl = 'Excel workbooks|*.xlsm'

# # relative Path
# path_xcl = forms.pick_file(files_filter=filterXcl, title="Choose excel file")

# fixed path
# path_xcl = "C:/Users/Harshal.Gunjal/Desktop/Attributprüfung_13350_2 2.xlsm"

# Import Excel data using custom utility
from guRoo_xclUtils import *
xcl = xclUtils([], python_path)
dat = xcl.xclUtils_import("Allgemeine Attribute", 5, 0) ###CHANGE SHEET NAME HERE###

# Extract data from Excel into lists
targets_params, target_bipgs, par_inst, par_formulae = [],[],[],[]
for row in dat[0][1:]:
    targets_params.append(row[0])
    target_bipgs.append(row[2])
    par_inst.append(row[3] == "Ja")
    par_formulae.append(row[4])

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
# for category in categories_in_active_view:
    # print(category)

# Create a new CategorySet
cats1 = app.Create.NewCategorySet()

# Insert categories from categories_in_active_view into the CategorySet
for category_name in categories_in_active_view:
    category = doc.Settings.Categories.get_Item(category_name)
    # print(category)
    if category is not None:
        cats1.Insert(category)

# Now cats1 contains all the categories present in the active view from filtered_categories

AM = doc.Settings.Categories.get_Item(BuiltInCategory.OST_GenericModel) # BIPG: Allgemeines Modell
cats1.Insert(AM)

# Create an iterator for cats1
itor = iter(cats1)

# Initialize an empty list to store loaded categories
loaded_categories = []

# Iterate over the set using the iterator
for item in itor:
    # Append the current item to the loaded_categories list
    loaded_categories.append(item.Name)

# Print the loaded categories
print("\n\nDie Projektparameter sind den folgenden Kategorien zugeordnet: ")
print(loaded_categories)

# INSERT NEW PARAMETERS
for p in req_param:
    try:
        insert_shared_parameter(app, p, cats1, BuiltInParameterGroup.PG_GENERAL, True)
    except Exception as e:
        print("Error inserting parameter '{}': {}".format(p, e))

# REINSERT BINDINGS
for p in req_param:
    try:
        reinsert_shared_parameter(app, p, cats1, BuiltInParameterGroup.PG_GENERAL, True)
    except Exception as e:
        print("Error reinserting parameter '{}': {}".format(p, e))
    
print("\n\nAllgemeine Attribute sind hinzugefügt")

output = script.get_output()
output.log_success("Allgemeine Attribute sind hinzugefügt")