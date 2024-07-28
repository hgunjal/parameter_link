# -*- coding: utf-8 -*-
__title__ = "Projektinformation"

__doc__ = """Version = 1.0
Date    = 06.05.2024
_____________________________________________________________________
Beschreibung:

Skript zum Hinzufügen von Projektinformationen
_____________________________________________________________________
Anleitung:

-> Klick auf die Button.
-> Falls ein Parameter fehlt, wird er automatisch hinzugefügt.
-> Füll die Werte der erforderlichen Projektparameter aus.
_____________________________________________________________________
Hinweis:
nur folgende Parameters werden angepasst.
'Projektnummer' -> IfcProject, 
'SiteName' -> IfcSite, 
'Gebäudebezeichnung' -> IfcBuilding.
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

from System.IO import File, Path

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
        raise Exception("Invalid Name Input!")  # Raise an exception if the specified parameter name is not found

    # Get the first definition (assuming there's only one parameter with the specified name)
    def_param = definitions[0]
    

    # Create a binding for either type or instance based on the 'inst' parameter
    binding = app.Create.NewTypeBinding(element_cat)
    if inst:
        # print("Selected object is Instance")
        binding = InstanceBinding(element_cat)  # New Binding
    
    t = Transaction(doc, "Add missing Projectinformation Parameter")
    t.Start()
    
    # Get the parameter bindings of the active document
    map = doc.ParameterBindings
    
    # Insert the shared parameter definition and its binding into the parameter bindings of the document
    map.Insert(def_param, binding, group)
    
    t.Commit()
    
    print("\nThe Parameter '{}' was added.".format(name))
    print("")

# MAIN

# projectInfo = doc.ProjectInformation

# projectName = projectInfo.Name
# ProjectNumber = projectInfo.Number
# # Projektadresse = projectInfo.Projektadresse

# # Print the project information using .format()
# print("Project Information:\n"
      # "Project Name: {}\n"
      # "Project Number: {}\n".format(projectName, ProjectNumber))
      
      
projectInfo = doc.ProjectInformation

project_info_dict = {}

# Print a header
# print("Project Information:")

# Loop through each parameter of the project information object
for param in projectInfo.Parameters:
    # Retrieve the parameter's name
    param_name = param.Definition.Name
    
    # Retrieve the parameter's value based on its storage type
    if param.StorageType == StorageType.String:
        param_value = param.AsString()
    elif param.StorageType == StorageType.Integer:
        param_value = param.AsInteger()
    elif param.StorageType == StorageType.Double:
        param_value = param.AsDouble()
    else:
        param_value = None  # Default for unsupported types
    
    # Add the parameter name and value to the dictionary
    project_info_dict[param_name] = param_value
    
    # Print the parameter name and value using .format()
    # print("{}: {}".format(param_name, param_value))

# Define the attributes you want to check for availability
required_attributes = ['Projektnummer', 'SiteName', 'Gebäudebezeichnung']

# Create an empty list to store the names of not available attributes
not_available_attributes = []

# Check the availability of each required attribute
print("\nAttribute Availability:")
for attribute in required_attributes:
    # Determine if the attribute is available in the dictionary
    is_available = attribute in project_info_dict
    print("{}: {}".format(attribute, "Available" if is_available else "Not Available"))
    
    # If the attribute is not available, add it to the list of not available attributes
    if not is_available:
        not_available_attributes.append(attribute)
    

# Output a completion message
print("\nSetting attributes completed.")
print("\n")

Original_SPF = app.SharedParametersFilename

# Determine the directory of the current script
script_directory = os.path.dirname(__file__)

# Define the relative path to the shared parameters file
temp_shared_parameters_file = os.path.join(script_directory, "IFC Shared Parameters-RevitIFCBuiltIn_ALL.txt")

app.SharedParametersFilename = temp_shared_parameters_file

# Create a new CategorySet
cats1 = app.Create.NewCategorySet()

AM = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ProjectInformation) # BIPG: ProjectInformation
cats1.Insert(AM)

# INSERT NEW PARAMETERS
for p in not_available_attributes:
    x = insert_shared_parameter(app, p, cats1, BuiltInParameterGroup.PG_DATA, True)
    
app.SharedParametersFilename = Original_SPF
    
# Iterate through each required attribute
for attribute in required_attributes:
    # Look up the parameter for the attribute in the project information object
    param = projectInfo.LookupParameter(attribute)
    
    # Check if the parameter exists
    if param is not None:
        # Parameter exists
        if param.StorageType == StorageType.String:
            # Determine the default value based on the attribute name
            if attribute == 'Projektnummer':
                default_value = "ABS 48"
            elif attribute == 'SiteName':
                default_value = "MBUX"
            elif attribute == 'Gebäudebezeichnung':
                default_value = "KIB"
            else:
                default_value = None  # Default for unsupported attributes
            
            # Ask the user for the value with a prompt
            selected_parameter_value = forms.ask_for_string(
                default=default_value,
                prompt='Bitte gib den Wert für {} ein'.format(attribute),
                title='Zuweisung von Projektinformation'
            )
            
            # Set the value of the parameter to the selected value
            t = Transaction(doc, "Adding Projektinformation")
            t.Start()
            param.Set(selected_parameter_value)
            t.Commit()
            print("{}: Set to '{}'".format(attribute, selected_parameter_value))
        else:
            # The parameter exists but is not a string type, print a message
            print("{}: Cannot set, not a string type parameter".format(attribute))
    else:
        # Attribute is not available in the project information object
        print("{}: Not Available".format(attribute))
