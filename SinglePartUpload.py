from tkinter import filedialog, messagebox
import openpyxl
import os
from EVOUtil import abreviateWords, isNone, createNewPart


file = filedialog.askopenfilename()

if not file:
    print("No file selected")
    os._exit(0)
    
print("File selected: ", file)

ws = openpyxl.load_workbook(file).active

"""
    info that i need
    
    partNumber
    description
    partClass
    partType
    mfg
    mfgNumber
    vendor
    vendorNumber
    specs
        
"""

partNumber = ws["A2"].value
description = ws["B2"].value
partClass = ws["C2"].value
partType = ws["D2"].value
mfg = ws["E2"].value
mfgNumber = ws["F2"].value
vendor = ws["G2"].value
vendorNumber = ws["H2"].value
specs = ws["I2"].value

if not partNumber or not description:
    messagebox.showerror("Error", "Part Number and Description are required")
    os._exit(0)
    

description = abreviateWords(description)
specs = abreviateWords(specs)

if isNone(specs):
    specs = ""
if isNone(vendor):
    vendor = ""
else:
    vendor = vendor[:4]+"50"
    
if isNone(vendorNumber):
    vendorPartNumber = ""
if isNone(mfg):
    mfg = ""
if isNone(mfgNumber):
    mfgPartNumber = ""
    
if isNone(partClass):
    partClass = ""
if isNone(partType):
    partType = ""





        partNumber = partNumber.strip()
        description = description.strip()
        partClass = partClass.strip()
        partType = partType.strip()
        mfg = mfg.strip()
        mfgNumber = mfgNumber.strip()
        vendor = findVendor(vendor.strip()[:4])
        vendorNumber = vendorNumber.strip()
        
        
        
        specs = specs.strip()
        
        middle2Digits = partNumber[5:7]
        print("middle2Digits", middle2Digits)
        print("partNumber", partNumber)
        
        classTypes = {
            "1": "FINM",
            "2": "ASSL",
            "3": "MECH",
            "4": "MSCE",
            "5": "MSCE",
            "6": "MECH",
            "7": "MECH",
            "8": "MSCE",
            "9": "ELEC",
        }

        if mfg == "" and mfgNumber == "" and vendor == "" and vendorNumber == "":
            partType = "A"
        elif mfg != "" and mfgNumber != "" and vendor != "" and vendorNumber != "": 
            partType = "R"
            
        else:
                
            
            missingInfo = []
            if mfg == "":
                missingInfo.append("Manufacturer")
            if mfgNumber == "":
                missingInfo.append("Manufacturer Number")
            if vendor == "":
                missingInfo.append("Vendor")
            if vendorNumber == "":
                missingInfo.append("Vendor Number")
                
            raise_error_message = f"Error: Incomplete information for part creation, part: {partNumber}.\n Missing information: {", ".join(missingInfo)}"
            
            
            root = tk.Tk()
            root.withdraw()
            tk.messagebox.showerror("Error", raise_error_message)
            
            
            if vendor == "" and vendor != "":
                tk.messagebox.showerror("Error", "Vendor not found in vendor list.")
                
            
            
            root.destroy()
            
            partType = "R"
        
        if partClass == "":
            print(middle2Digits)
            partClass = classTypes[middle2Digits[0]]


createNewPart(partNumber, description, partClass, partType, mfg, mfgNumber, vendor, vendorNumber, specs)

messagebox.showinfo("Success", "Part has been created")