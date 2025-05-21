import clr

# Load necessary Revit and Dynamo API references
clr.AddReference('RevitServices')
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

# Get current Revit document
doc = DocumentManager.Instance.CurrentDBDocument

# Check if wall input exists
if not IN[0]:
    OUT = "No walls selected."
else:
    # Unwrap inputs from Dynamo
    selected_walls = UnwrapElement(IN[0])
    selected_assemblies = UnwrapElement(IN[1]) if IN[1] else []

    # Prepare result containers
    filtered_walls = []
    filtered_assemblies = []
    walls_in_assemblies = set()

    # Start transaction
    TransactionManager.Instance.EnsureInTransaction(doc)

    try:
        # Process selected assemblies
        for assembly in selected_assemblies:
            if assembly:
                # Get all wall elements that belong to this assembly
                assembly_walls = [
                    doc.GetElement(member_id)
                    for member_id in assembly.GetMemberIds()
                    if doc.GetElement(member_id).Category.Id.IntegerValue == int(BuiltInCategory.OST_Walls)
                ]

                # Check if any wall in the assembly has an empty required parameter
                if any(wall.LookupParameter("Custom_Param_1").AsString() == "" for wall in assembly_walls):
                    filtered_assemblies.append(assembly)

                    for wall in assembly_walls:
                        walls_in_assemblies.add(wall.Id)

                        # Assign assembly ID to each wall in a tag parameter
                        param_tag_assembly = wall.LookupParameter("Assembly_Tag_1")
                        if param_tag_assembly and not param_tag_assembly.IsReadOnly:
                            param_tag_assembly.Set(str(assembly.Id))

        # Process individual selected walls
        for wall in selected_walls:
            if wall and wall.Id not in walls_in_assemblies:
                # Check if wall is missing the key parameter
                param_main = wall.LookupParameter("Custom_Param_1")
                if param_main and (not param_main.AsString() or param_main.AsString().strip() == ""):
                    filtered_walls.append(wall)

                    # Assign wall ID to wall's own tag parameter
                    param_tag = wall.LookupParameter("Wall_Tag_1")
                    if param_tag and not param_tag.IsReadOnly:
                        param_tag.Set(str(wall.Id))

        # Commit transaction
        TransactionManager.Instance.TransactionTaskDone()

    except Exception as e:
        # Roll back in case of error
        TransactionManager.Instance.ForceCloseTransaction()

    # Output filtered elements
    OUT = (filtered_walls, filtered_assemblies)
