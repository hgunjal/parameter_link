# Import necessary libraries from pyRevit
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import ObjectType
from pyrevit import DB, revit, script, forms

# Prompt user to specify Excel file path
filterXcl = 'Excel workbooks|*.xlsm'
# filterRfa = 'Family Files|*.rfa'
path_xcl = forms.pick_file(files_filter=filterXcl, title="Choose excel file")


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
dat = xcl.xclUtils_import("Allgemeine Attribute", 5, 0) ###CHANGE SHEET NAME HERE###

# Extract data from Excel into lists
targets_params, target_bipgs, fam_inst, fam_wert = [],[],[],[]
for row in dat[0][1:]:
    targets_params.append(row[0])
    target_bipgs.append(row[2])
    fam_inst.append(row[3] == "Ja")
    fam_wert.append(row[4])
    
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

def famDoc_addSharedParams(famDoc, famDefs, famBipgs, famInst, famValue):
    if famDoc.IsFamilyDocument is not None:
        famMan = famDoc.FamilyManager
        parNames = [p.Definition.Name for p in famMan.Parameters]
        t = Transaction(famDoc, 'Add parameters')
        t.Start()
        params = []
        
        # Check if the objekt exists in the dictionary

        for d, b, i, f in zip(famDefs, famBipgs, famInst, famValue):
            # Check if the parameter definition is in the list of parameters corresponding to the objekt
            if d.Name not in parNames:
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
        
        params_added = famDoc_addSharedParams(famDoc, fam_defs, fam_bipgs, fam_inst, fam_wert)
        if pb.cancelled or len(params_added) == 0:
            forms.alert("Failed to update parameters in the active family document.", title="Script cancelled")
        else:
            passCount += 1
    pb.update_progress(pbCount, pbTotal)

# Display final message to user
form_message = str(passCount) + "/" + str(pbTotal) + " family updated."
forms.alert(form_message, title="Script completed", warn_icon=False)