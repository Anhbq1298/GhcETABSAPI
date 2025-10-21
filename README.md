# GhcETABSAPI

GhcETABSAPI is a Grasshopper plug-in that wraps parts of the ETABS API so Rhino users can automate model setup, unit management, and Excel handoffs without leaving the Rhino/Grasshopper environment.

## Project structure

The solution is organised as Grasshopper component categories that mirror the tabs visible inside Grasshopper:

- `Category0-Utility/GhcWriteGHTreeToExcel.cs` – exports a Grasshopper data tree to Excel using `ExcelHelpers` for COM handling.
- `Category1-IO/` – connection and unit utilities including:
  - `GhcAttachETABSInstance.cs` – attaches to an ETABS session.
  - `GhcGetETABSUnits.cs` and `GhcGetRhinoUnits.cs` – report active unit systems.
  - `GhcUnitScaleFactor .cs` – computes length conversion factors.
- `Category2-AddFrame/` – frame element tools such as section assignment, creation from Rhino curves, grouping, and load extraction.
- `Category3-AddShell/GhcAddShellsFromPolylines.cs` – creates ETABS shell elements from Grasshopper geometry.
- `ExcelHelpers.cs` – shared Excel automation helpers used by export components.
- `Resources/` – embedded icons referenced by the Grasshopper components.

## Build requirements

- Targets .NET Framework 4.8 and produces a `.gha` assembly (`GhcETABSAPI.csproj`).
- References Grasshopper 8.0 and the ETABS COM interop (`ETABSv1.dll`).
- Uses Microsoft Office Excel interop with embedded COM types for deployment.

## Post-build deployment

The `PostBuild` target copies the compiled `.gha` and any ETABS interop DLLs into the user's Grasshopper libraries folder so Rhino can discover the plug-in automatically.

## Development tips

- When running outside of Rhino, ensure the Microsoft Office Excel interop is available locally so the Excel helper functions succeed.
- Keep new components within the existing category folders so the Grasshopper UI remains organised.
