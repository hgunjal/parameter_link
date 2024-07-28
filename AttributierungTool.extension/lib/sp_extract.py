# -*- coding: utf-8 -*-

# IMPORTS
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms
from pyrevit import HOST_APP
from pyrevit import DB, revit
import sys

# VARIABLES
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document

# FUNCTIONS

def insert_shared_parameter(app, name, element_cat, group, inst):
    # Open the shared parameter file associated with the application
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception("No SharedParameter File!")  # Raise an exception if no shared parameter file is found

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
    
    t = Transaction(doc, "Attach SP as PP")
    t.Start()
    
    # Get the parameter bindings of the active document
    map = doc.ParameterBindings
    
    # Insert the shared parameter definition and its binding into the parameter bindings of the document
    map.Insert(def_param, binding, group)
    
    t.Commit()
    
    # Print the name of the parameter added


def reinsert_shared_parameter(app, name, element_cat, group, inst):
    # Open the shared parameter file associated with the application
    def_file = app.OpenSharedParameterFile()
    if def_file is None:
        raise Exception("No SharedParameter File!")  # Raise an exception if no shared parameter file is found

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
    
    # Print the name of the parameter and the category it is added to

def check_loaded_params(list_p_names):
    """Check if any parameters from provided list are missing in the project"""
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
    """Check if any parameters from provided list are missing in the specified category"""
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
    """This function will prompt the user to select elements in Revit UI and return their element IDs
    :param uidoc:   uidoc where elements are selected.
    :return:        List of selected element IDs"""
    selected_elements = []
    try:
        reference = uidoc.Selection.PickObjects(ObjectType.Element, "Wähl die Objekte aus, zu denen Allgemeine Attribute hinzugefügt werden sollen.")
        for ref in reference:
            selected_elements.append(ref.ElementId)
    except Exception as ex:
        print("Fehler beim Auswählen von Objekten:", ex)
    return selected_elements