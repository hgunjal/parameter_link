# -*- coding: utf-8 -*-

__title__ = "Objekttyp auswählen und weitere Attribute hinzufügen"

__doc__ = """Version = 1.0
Date    = 02.04.2024
_____________________________________________________________________
Beschreibung:

Skript zum Auswählen des Objekttyps.
_____________________________________________________________________
Anleiting:

-> Wähl die Excel-Datei mit der Liste der Objekte.
-> Wähl ein oder mehrere Objekte aus, die zum gleichen Objekttyp gehören.
-> Klick auf "Fertig stellen".
-> Wähl den Objekttyp.
_____________________________________________________________________"""

# IMPORTS
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms

# Import necessary libraries from pyRevit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms

# VARIABLES
uidoc = __revit__.ActiveUIDocument
doc   = __revit__.ActiveUIDocument.Document
app = __revit__.Application

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

def famDoc_setParameter(famDoc,p,v):
    try:
        t = Transaction(famDoc, "ObjectTyp auswählen")
        t.Start()
        famMan = famDoc.FamilyManager
        famMan.Set(p,v)
        t.Commit()
        return True
    except:
        return False

# with forms.WarningBar(title='Wähl die Objekte aus, um den Objekttyp zu definieren.:'):
    # selected_element_ids = get_selected_elements(uidoc)
    # print('Ausgewählte Element-Namen:')
    # for element_id in selected_element_ids:
        # element = doc.GetElement(element_id)
        # print(element.Name)  # Assuming the selected elements have a 'Name' property

filterXcl = 'Excel workbooks|*.xlsm'

# relative Path
# path_xcl = forms.pick_file(files_filter=filterXcl, title="Choose excel file")

# fixed path
path_xcl = "C:/Users/Harshal.Gunjal/Desktop/Attributprüfung_13350_2 2.xlsm"

from guRoo_xclUtils import *
xcl = xclUtils([], path_xcl)
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
print("Ausgewähltes Objekt Name: {}".format(selected_option))

import enum
import System
from System import Enum
famDoc = revit.doc
famMan = revit.doc.FamilyManager
famParamsList =[]
famParamNamesList =[]

params = famMan.GetParameters()
paramNames = [p.Definition.Name for p in params]
famParamsList.append(params)
famParamNamesList.append(paramNames)

# Initialize variables to store parameters and parameter names
objekt_parameter = None
objekt_parameter_name = None

# Iterate through the parameters
for param in params:
    # Get the parameter name
    param_name = param.Definition.Name
    
    # Check if the parameter name is "Objekt"
    if param_name == "Objekt":
        # Store the parameter and its name
        objekt_parameter = param
        objekt_parameter_name = param_name
        # Break the loop since we found the parameter
        break

famDoc_setParameter(famDoc,objekt_parameter,selected_option)

# Import necessary libraries from pyRevit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms

# Prompt user to specify Excel file path
filterXcl = 'Excel workbooks|*.xlsm'
# filterRfa = 'Family Files|*.rfa'
# path_xcl = forms.pick_file(files_filter=filterXcl, title="Choose excel file")

# fixed path
path_xcl = "C:/Users/Harshal.Gunjal/Desktop/Attributprüfung_13350_2 2.xlsm"


# Exit script if no file is selected
if not path_xcl:
    script.exit()

# Get the active document in Revit
doc = __revit__.ActiveUIDocument.Document

# Exit script if no active document is found
if doc is None:
    forms.alert("No active document found. Please open a family document.", title="Script cancelled")
    script.exit()
    
# Import Excel data using custom utility
from guRoo_xclUtils import *
xcl = xclUtils([], path_xcl)
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
# print("Objekt Parameter Mapping:")
# for objekt, params in objekt_param_mapping.items():
    # print("{0}: {1}".format(objekt, params))

objekt_list = list(objekt_param_mapping.keys())
# print(objekt_list)
    
# Get shared parameters from Revit
app = __revit__.Application
fam_defs, fam_bipgs = [],[]

# Retrieve shared parameter definitions and names
spFile   = app.OpenSharedParameterFile()
spGroups = spFile.Groups
sp_defs, sp_nams = [],[]
for g in spGroups:
    for d in g.Definitions:
        sp_defs.append(d)
        sp_nams.append(d.Name)
        
# Get target parameter definitions
for t in targets_params:
    if t in sp_nams:
        ind = sp_nams.index(t)
        fam_defs.append(sp_defs[ind])
        
# Exit script if not all parameters are found
if len(fam_defs) != len(targets_params):
    forms.alert("Some parameters not found, refer to report for details.", title="Script cancelled")
    print("NOT FOUND IN SHARED PARAMETERS FILE:")
    print("---")
    for t in targets_params:
        if t not in sp_nams:
            print(t)
    script.exit()
    
# Import enum
import System
from System import Enum

# Get all built-in parameter groups for checking
bipgs = [a for a in System.Enum.GetValues(DB.BuiltInParameterGroup)]
bipg_names = [str(a) for a in bipgs]

# Retrieve built-in parameter groups for target bipgs
for t in target_bipgs:
    if t in bipg_names:
        ind = bipg_names.index(t)
        fam_bipgs.append(bipgs[ind])

# Exit script if not all BIPGs are found
if len(fam_bipgs) != len(target_bipgs):
    forms.alert("Some groups not found, refer to report for details.", title="Script cancelled")
    print("NOT A VALID PARAMETER GROUP NAME:")
    print("---")
    for t in target_bipgs:
        if t not in bipg_names:
            print(t)
    script.exit()

# Function to open family document
def famDoc_open(filePath, app):
    try:
        famDoc = app.OpenDocumentFile(filePath)
        return famDoc
    except:
        return None

# Function to close and save family document
def famDoc_close(famDoc, saveOpt=True):
    try:
        famDoc.Close(saveOpt)
        return 1
    except:
        return 0

# Functions to add parameters
from Autodesk.Revit.DB import Transaction

def famDoc_addSharedParams(famDoc, famDefs, famBipgs, famInst, famValue, objekt=None):
    if famDoc.IsFamilyDocument and objekt is not None:
        famMan = famDoc.FamilyManager
        parNames = [p.Definition.Name for p in famMan.Parameters]
        t = Transaction(famDoc, 'Add parameters')
        t.Start()
        params = []
        
        # Check if the objekt exists in the dictionary
        if objekt in objekt_param_mapping:
            objekt_params = objekt_param_mapping[objekt]
            for d, b, i, f in zip(famDefs, famBipgs, famInst, famValue):
                # Check if the parameter definition is in the list of parameters corresponding to the objekt
                if d.Name in objekt_params and d.Name not in parNames:
                    p = famMan.AddParameter(d, b, i)
                    params.append(p)
                    if f is not None:
                        try:
                            famMan.Set(p, f)
                        except:
                            pass
        t.Commit()
        return params
    else:
        return None




# Undertake the process with a progress bar
with forms.ProgressBar(step=1, title="Updating families", cancellable=True) as pb:
    pbCount = 1
    pbTotal = 1  
    passCount = 0
    famDoc = doc  
    if famDoc != None:
        
        params_added = famDoc_addSharedParams(famDoc, fam_defs, fam_bipgs, fam_inst, fam_wert, objekt=selected_option)
        if pb.cancelled or len(params_added) == 0:
            forms.alert("Failed to update parameters in the active family document.", title="Script cancelled")
        else:
            passCount += 1
    pb.update_progress(pbCount, pbTotal)

# Display final message to user
form_message = str(passCount) + "/" + str(pbTotal) + " family updated."
forms.alert(form_message, title="Script completed", warn_icon=False)

