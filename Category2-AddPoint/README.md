# Category2-AddPoint

Point related Grasshopper components for ETABS integration.

- `GhcAddPoints.cs` – Create new point objects inside ETABS.
- `GhcGetPointInfo.cs` – Query point names and XYZ coordinates (tree output aligns with Excel sync component).
- `GhcSyncPointCoordinatesFromExcel.cs` – Diff an Excel sheet (UniqueName/X/Y/Z) against a captured baseline tree, highlight edits, and push renames/coordinate updates back to ETABS by selecting points and driving `EditGeneral.Move`.
- `GhcSetPointInfo_ExcelInteraction.cs` – Legacy component kept for backward compatibility with workflows that require additional label/story metadata.
