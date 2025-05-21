import clr

# Revit and Dynamo API references
clr.AddReference('RevitServices')
clr.AddReference('RevitAPI')
clr.AddReference('RevitNodes')
clr.AddReference('ProtoGeometry')
clr.AddReference('System')

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, ElementId, Transaction
import System
from System.IO import File, FileInfo

# Required for Excel Interop
clr.AddReference('Microsoft.Office.Interop.Excel')
import Microsoft.Office.Interop.Excel as Excel

# Get the active Revit document
doc = DocumentManager.Instance.CurrentDBDocument

# Input: Excel file reference (via File.FromPath in Dynamo)
file_info = IN[0]
file_path = file_info.FullName

# --------------------------------------------
# Function to read data from Excel spreadsheet
# --------------------------------------------
def process_excel_data():
    data = []
    excel_app = None
    workbook = None
    try:
        excel_app = Excel.ApplicationClass()
        excel_app.Visible = False
        workbook = excel_app.Workbooks.Open(file_path)
        worksheet = workbook.Sheets[1]

        row = 2  # Start reading from row 2
        while True:
            section_number = worksheet.Cells[row, 1].Value2
            coefficient = worksheet.Cells[row, 6].Value2
            
            # End of data check
            if section_number is None and coefficient is None:
                break
            
            if section_number and coefficient is not None:
                wall_ids = []
                for col in range(7, 13):  # Columns G to L = wall IDs
                    cell_value = worksheet.Cells[row, col].Value2
                    if cell_value:
                        try:
                            ids = [int(id.strip()) for id in str(cell_value).split() if id.strip()]
                            wall_ids.extend(ids)
                        except ValueError:
                            continue
                
                if wall_ids:
                    data.append([str(section_number), float(coefficient), wall_ids])
            row += 1
    except Exception as e:
        return [], "Error while reading Excel: " + str(e)
    finally:
        if workbook:
            workbook.Close(False)
        if excel_app:
            excel_app.Quit()
    return data, ""

# ----------------------------------------------------
# Function to apply the data from Excel to Revit model
# ----------------------------------------------------
def update_revit_model(data):
    walls = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()
    assemblies = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Assemblies).WhereElementIsNotElementType().ToElements()

    wall_dict = {wall.Id.IntegerValue: wall for wall in walls}
    assembly_dict = {assembly.Id.IntegerValue: assembly for assembly in assemblies}

    t = Transaction(doc, "Update wall and assembly parameters")
    t.Start()
    try:
        log = []
        for row_index, row in enumerate(data):
            section_number, coefficient, wall_ids = row
            log.append("Processing row {0}: {1}".format(row_index + 2, wall_ids))
            
            for i, wall_id in enumerate(wall_ids):
                if wall_id in wall_dict:
                    wall = wall_dict[wall_id]

                    # Parameter 1 (e.g. reinforcement ratio)
                    param1 = wall.LookupParameter("Custom_Param_1")
                    # Parameter 2 (e.g. reinforcement ratio alt)
                    param2 = wall.LookupParameter("Custom_Param_2")

                    if param1 and not param1.IsReadOnly:
                        if i == 0:
                            formatted_value = "{0:.2f} kg/m³".format(coefficient)
                            param1.Set(formatted_value)
                            log.append("Wall {0}: set Param 1 to {1}".format(wall_id, formatted_value))
                        else:
                            param1.Set("-")
                            log.append("Wall {0}: set Param 1 to dash".format(wall_id))

                    if param2 and not param2.IsReadOnly:
                        formatted_value = "{0:.2f} kg/m³".format(coefficient)
                        param2.Set(formatted_value)
                        log.append("Wall {0}: set Param 2 to {1}".format(wall_id, formatted_value))
                    else:
                        log.append("Wall {0}: Param 2 is missing or read-only".format(wall_id))

                    # Set section number in other ID-related parameters if empty
                    for param_name in ["Mark", "Drawing_Label"]:
                        param = wall.LookupParameter(param_name)
                        if param and not param.IsReadOnly and not param.HasValue:
                            param.Set(section_number)
                            log.append("Wall {0}: set {1} to {2}".format(wall_id, param_name, section_number))

                elif wall_id in assembly_dict:
                    assembly = assembly_dict[wall_id]
                    param2 = assembly.LookupParameter("Custom_Param_2")
                    if param2 and not param2.IsReadOnly:
                        formatted_value = "{0:.2f} kg/m³".format(coefficient)
                        param2.Set(formatted_value)
                        log.append("Assembly {0}: set Param 2 to {1}".format(wall_id, formatted_value))
                    else
