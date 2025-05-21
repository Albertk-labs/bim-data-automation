# Revit BIM Automation Toolkit

This repository contains Python scripts for Dynamo that automate key BIM workflows in Autodesk Revit.  
The scripts allow you to detect and tag elements with missing parameter values and to update model data based on Excel files.

Designed for infrastructure and prefab workflows where data completeness and coordination are critical.

---

##  Included Scripts

### 1. `FilterAndTagWalls.py`

Filters walls and assemblies based on empty custom parameters and tags them with identifiers for later use.

####  Purpose:
- Identify walls with missing values in a given parameter
- Tag elements to prepare them for data population via external Excel

####  Inputs:
- `IN[0]`: List of selected walls
- `IN[1]`: (Optional) List of selected assemblies

####  Process:
- Finds walls inside assemblies with empty `Custom_Param_1`
- Tags these walls with the Assembly ID (`Assembly_Tag_1`)
- Finds stand-alone walls with empty `Custom_Param_1`
- Tags these walls with their own ID (`Wall_Tag_1`)

####  Output:
- Two lists:
  - Filtered walls
  - Filtered assemblies

---

### 2. `UpdateWallsFromExcel.py`

Updates Revit walls and assemblies with section numbers and reinforcement coefficients based on an Excel file.

####  Purpose:
- Automatically populate model parameters from Excel metadata
- Standardize reinforcement values and section labels

####  Inputs:
- `IN[0]`: File path to Excel document (via `File.FromPath` in Dynamo)

####  Process:
- Reads section number, coefficient, and wall IDs from the Excel file
- Finds matching elements in the Revit model
- Writes to parameters:
  - `Custom_Param_1`
  - `Custom_Param_2`
  - `Mark`, `Drawing_Label` (if available and empty)

####  Output:
- Operation status
- Log of all changes made

---

##  Recommended Workflow

1. **Use `FilterAndTagWalls.py`**  
   Tag elements with missing data to prepare them for Excel export.

2. **Export to Excel**  
   (outside the script — using a Revit schedule or Dynamo script)

3. **Use `UpdateWallsFromExcel.py`**  
   Read Excel data and populate parameters back into Revit.

---

##  Repository Structure
/scripts
├── FilterAndTagWalls.py
├── UpdateWallsFromExcel.py

/sample_data
├── example_input.xlsx

README.md
.gitignore
---

## ⚙ Technologies

- Revit API
- Dynamo for Revit
- Microsoft Excel Interop (COM)
- Python 2 (IronPython for Dynamo)

---

##  Author

**Albert Kłoczewiak**  
GitHub: [@albertk-labs](https://github.com/albertk-labs)  
Email: akkloczewiak@gmail.com
