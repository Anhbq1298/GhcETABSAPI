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

## Duplicate code review

- `Category2-AddFrame/GhcSetLoadDistOnFrames_Rel.cs` contains private copies of helpers such as `EnsureModelUnlocked`, `TryGetExistingFrameNames`, `Clamp01`, and `ClampDirCode` even though the same implementations already live in `Category0-Utility/ComponentShared.cs`. Consolidating these would reduce maintenance overhead.
- Both `Category2-AddFrame/GhcSetLoadDistOnFrames_Rel2.cs` and `Category3-AddShell/GhcSetLoadUniformOnAreas2.cs` ship almost identical Excel ingestion pipelines (`ReadExcelSheet` plus `ExcelLoadData` containers) that could be merged into shared utilities.
- `Category3-AddShell/GhcSetLoadUniformOnAreas.cs` reimplements coordinate-system resolution and result summarising logic that overlaps with the Excel-driven area loader, suggesting an opportunity to share result-formatting helpers across the two components.

## Suggested enhancements

- Move the duplicated helper methods noted above into `ComponentShared` (or expand it) so all components reuse a single implementation for model unlocking, direction handling, and pluralisation logic.
- Extract the Excel worksheet parsing routines that currently live inside the frame and area Excel components into `ExcelHelpers` to avoid divergent behaviour when column expectations change.
- Rename files that carry trailing spaces (for example `Category1-IO/GhcUnitScaleFactor .cs` and `Category2-AddFrame/GhcAddFramesToGroup .cs`) to prevent duplicate-name issues on stricter filesystems and to keep future Git diffs clean.
