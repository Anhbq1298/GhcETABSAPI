# MGT – Meinhardt Grasshopper Tool

MGT is a Grasshopper plug-in that wraps key pieces of the ETABS API so structural engineers can automate common modelling and
documentation workflows without leaving the Rhino/Grasshopper environment. The plug-in focuses on bridging ETABS model data
with Rhino geometry and Excel-based handover documents to streamline multidisciplinary coordination.

## Key capabilities

- **Attach to ETABS sessions** – connect to an existing ETABS model or spin up a new instance directly from Grasshopper.
- **Synchronise units** – read the active units from both ETABS and Rhino so geometric inputs stay consistent.
- **Author frame and shell elements** – generate beams, columns, braces, and shell meshes from Rhino curves and polylines.
- **Manage metadata** – assign section properties, groups, and loads as part of the element creation process.
- **Export data to Excel** – capture Grasshopper data trees in Excel spreadsheets for downstream coordination.

## Repository layout

The solution is organised as Grasshopper component categories that mirror the tabs visible inside Grasshopper:

| Folder | Purpose |
| --- | --- |
| `Category0-Utility/` | Helper utilities including `GhcWriteGHTreeToExcel` that serialises Grasshopper trees into Excel using COM interop. |
| `Category1-IO/` | Connection and unit utilities (`GhcAttachETABSInstance`, `GhcGetETABSUnits`, `GhcGetRhinoUnits`, `GhcUnitScaleFactor`). |
| `Category2-AddFrame/` | Frame element authoring components: create members from Rhino curves, assign sections, group elements, and extract loads. |
| `Category3-AddShell/` | Shell element tools, notably `GhcAddShellsFromPolylines` for creating ETABS shell elements from planar polylines. |
| `_Helpers/ExcelHelpers.cs` | Shared Excel automation helpers used by export components. |
| `Resources/` | Embedded icons referenced by the Grasshopper components. |

Additional top-level files include the Visual Studio solution (`MGT.sln`), the class library project (`MGT.csproj`), and the ETABS
interop assembly (`ETABSv1.dll`).

## Build requirements

- Targets **.NET Framework 4.8**.
- References **Grasshopper 8.0** assemblies.
- Uses the **ETABS COM interop** (`ETABSv1.dll`) for API access.
- Relies on **Microsoft Office Excel Interop** with embedded COM types for deployment without requiring local Excel installation.

## Building the plug-in

1. Open `MGT.sln` in Visual Studio 2022 (or newer) with the .NET desktop development workload.
2. Restore any missing references:
   - `Grasshopper.dll` (installed with Rhino/Grasshopper).
   - `ETABSv1.dll` (bundled in the repository).
3. Build the `MGT` project in **Release** mode to produce the `MGT.gha` assembly.

## Post-build deployment

The project includes a `PostBuild` target that copies the compiled `.gha` and supporting ETABS interop DLLs into the current
user's Grasshopper libraries folder. This allows Rhino to discover the plug-in automatically the next time Grasshopper is opened.
If the automatic copy fails, manually place the contents of the `bin/Release` folder into `%AppData%\Grasshopper\Libraries`.

## Using the components

1. Launch Rhino and open Grasshopper.
2. Locate the **MGT** tab within Grasshopper's component ribbon.
3. Drag the desired component onto the canvas:
   - Start with the IO components to attach to ETABS and synchronise units.
   - Use the Add Frame/Shell components to author geometry-driven ETABS elements.
   - Leverage the Utility components to export Grasshopper data trees to Excel for reporting.
4. Monitor the Rhino command line for status messages while the components interact with ETABS or Excel.

## Development guidelines

- Keep new components within the existing category folders so the Grasshopper UI remains organised.
- Follow Grasshopper naming conventions: prefix component classes with `Ghc` and provide clear `Category` and `SubCategory` values.
- Reuse helpers in `_Helpers` for common COM/interop logic to avoid duplicate boilerplate.
- When running outside of Rhino, ensure the Microsoft Office Excel interop is installed locally so the Excel helper functions succeed.

## Troubleshooting

- **Cannot attach to ETABS:** Verify ETABS is running or that the ETABS interop is registered. Launch ETABS manually before using the attach component.
- **Missing references during build:** Confirm that Rhino/Grasshopper is installed and the `Grasshopper.dll` path is correct. Re-add the reference if needed.
- **Excel export failures:** Ensure Excel (or the necessary interop assemblies) is installed and accessible. Running Visual Studio as administrator can resolve COM permission issues.

## Further resources

- [ETABS API documentation](https://docs.csiamerica.com) – official API reference from Computers and Structures, Inc.
- [Grasshopper developer guides](https://developer.rhino3d.com/guides/grasshopper/) – tips for authoring custom components.
- [RhinoCommon SDK](https://developer.rhino3d.com/api/) – underlying Rhino APIs used by Grasshopper components.

