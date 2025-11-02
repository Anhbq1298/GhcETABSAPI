using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using ETABSv1;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using static MGT.ComponentShared;
using Excel = Microsoft.Office.Interop.Excel;

namespace MGT
{
    /// <summary>
    /// Reads point information (UniqueName, X, Y, Z) from Excel and applies
    /// changes to ETABS: create (default), rename, move, delete (if removed in Excel
    /// and with no connectivity). Reader streams Excel with accurate 100% progress.
    ///
    /// Notes:
    /// - Create-if-missing is ALWAYS ON (no input pin).
    /// - Excel progress uses candidate rows via End(xlUp) and the reader does the final tick.
    /// - Assignment progress updates after each processed row and final tick hits 100%.
    /// </summary>
    public class GhcSetPointInfo : GH_Component
    {
        // ======= Constants =======
        private const string DefaultSheet = "Point Info";  // Default worksheet name
        private const int StartColumn = 2;                 // Column B
        private const int StartRow = 2;                    // Data starts at row 2 (row 1 = headers)
        private const int ColumnCount = 4;                 // Columns: UniqueName, X, Y, Z
        private const double Tolerance = 1e-6;             // Movement tolerance (in model units)
        private const int ExcelChunkSize = 2000;           // Candidate rows per block read

        // ======= Sticky replay (non-rising) =======
        private bool _lastRun;
        private GH_Structure<GH_ObjectWrapper> _lastValues = new GH_Structure<GH_ObjectWrapper>();
        private readonly List<string> _lastActions = new();
        private readonly List<string> _lastMessages = new() { "No previous run. Toggle 'run' to execute." };

        public GhcSetPointInfo()
            : base(
                "Set Point Info",
                "SetPointInfo",
                "Create/rename/move/delete ETABS point objects from an Excel worksheet.",
                "MGT",
                "2.0 Point Object Modelling")
        { }

        public override Guid ComponentGuid => new Guid("A9B6F07F-7D5E-4A25-AD2A-6F0A7AE12C47");
        protected override Bitmap Icon => null;

        // ======= Inputs =======
        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("run", "run", "Rising edge trigger.", GH_ParamAccess.item, false);
            pManager.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            pManager.AddTextParameter("excelPath", "excelPath", "Full or project relative path to the workbook.", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("sheetName", "sheetName", "Worksheet name containing the data.", GH_ParamAccess.item, DefaultSheet);

            int baselineIndex = pManager.AddGenericParameter(
                "baseline",
                "baseline",
                "Optional tree captured from GhcGetPointInfo. Used for deletes and old names.",
                GH_ParamAccess.tree);
            pManager[baselineIndex].Optional = true;
        }

        // ======= Outputs =======
        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("values", "values", "Column-wise data tree (UniqueName / X / Y / Z).", GH_ParamAccess.tree);
            pManager.AddTextParameter("actions", "actions", "Create/rename/move/delete actions executed on ETABS points.", GH_ParamAccess.list);
            pManager.AddTextParameter("messages", "messages", "Summary and diagnostic messages.", GH_ParamAccess.list);
        }

        // ======= Solve =======
        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sapModel = null;
            string excelPath = null;
            string sheetName = DefaultSheet;

            da.GetData(0, ref run);
            da.GetData(1, ref sapModel);
            da.GetData(2, ref excelPath);
            da.GetData(3, ref sheetName);
            da.GetDataTree(4, out GH_Structure<IGH_Goo> baselineTree);

            // Rising-edge gate
            bool rising = !_lastRun && run;
            if (!rising)
            {
                da.SetDataTree(0, _lastValues.Duplicate());
                da.SetDataList(1, _lastActions.ToArray());
                da.SetDataList(2, _lastMessages.ToArray());
                _lastRun = run;
                return;
            }

            var messages = new List<string>();
            var actions = new List<string>();
            var valueTree = new GH_Structure<GH_ObjectWrapper>();

            try
            {
                // ===== Validate inputs =====
                if (sapModel == null)
                    throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");

                string fullPath = ExcelHelpers.ProjectRelative(excelPath);
                if (string.IsNullOrWhiteSpace(fullPath))
                    throw new InvalidOperationException("excelPath is empty.");

                if (!File.Exists(fullPath))
                    throw new FileNotFoundException("Excel workbook not found.", fullPath);

                if (string.IsNullOrWhiteSpace(sheetName))
                    sheetName = DefaultSheet;

                // ===== Progress UI =====
                UiHelpers.ShowDualProgressBar(
                    "Set Point Info",
                    "Reading Excel...",
                    0,
                    "Updating points...",
                    0);

                // ===== Read Excel (streamed; reader handles the final tick) =====
                PointBaseline baseline = PointBaseline.FromTree(baselineTree);

                List<PointRow> rows = ReadExcelRows_ChunkedStreaming(
                    fullPath,
                    sheetName,
                    (current, total, status) => UiHelpers.UpdateExcelProgressBar(current, total, status));

                valueTree = BuildValueTree(rows);
                // NOTE: do not add an extra final tick here — the reader already did (candidateCount, candidateCount).

                // ===== Apply to ETABS =====
                int totalRows = rows.Count;
                if (totalRows == 0)
                {
                    messages.Add("Excel sheet contained no data rows.");
                    UiHelpers.UpdateAssignmentProgressBar(0, 0, "No rows to update.");
                }
                else
                {
                    EnsureModelUnlocked(sapModel);

                    HashSet<string> existingNames = GetPointNames(sapModel);
                    if (existingNames == null)
                        throw new InvalidOperationException("Failed to read point names from ETABS.");

                    UiHelpers.UpdateAssignmentProgressBar(0, totalRows, BuildAssignmentStatus(0, totalRows));

                    const bool createIfMissing = true; // <== ALWAYS ON
                    ProcessRows_CreateMoveDelete(
                        sapModel,
                        rows,
                        baseline,
                        existingNames,
                        createIfMissing,
                        actions,
                        messages,
                        (current, total, status) => UiHelpers.UpdateAssignmentProgressBar(current, total, status));

                    // Final assignment tick to visually reach 100%
                    UiHelpers.UpdateAssignmentProgressBar(totalRows, totalRows, BuildAssignmentStatus(totalRows, totalRows));
                }

                // Non-fatal view refresh
                try { sapModel.View.RefreshView(0, false); } catch { /* ignore */ }

                int refreshed = TriggerGetPointInfoRefresh();
                if (refreshed > 0)
                    messages.Add($"Scheduled refresh for {refreshed} Get Point Info component(s).");
            }
            catch (Exception ex)
            {
                messages.Clear();
                messages.Add("Error: " + ex.Message);
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
            }
            finally
            {
                UiHelpers.CloseProgressBar();

                da.SetDataTree(0, valueTree);
                da.SetDataList(1, actions);
                da.SetDataList(2, messages);

                _lastValues = valueTree.Duplicate();
                _lastActions.Clear();
                _lastActions.AddRange(actions);
                _lastMessages.Clear();
                _lastMessages.AddRange(messages);
                _lastRun = run;
            }
        }

        /// <summary>
        /// STREAMED EXCEL READER:
        /// - Compute candidateCount via End(xlUp) across B..E.
        /// - Read blocks of rows (object[,]) to reduce COM calls.
        /// - Advance progress for every candidate row (incl. blanks) so the bar
        ///   runs smoothly and exactly hits 100%.
        /// - Performs the final tick: progress(candidateCount, candidateCount).
        /// </summary>
        private static List<PointRow> ReadExcelRows_ChunkedStreaming(
            string workbookPath,
            string sheetName,
            Action<int, int, string> progress)
        {
            var rows = new List<PointRow>();

            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;

            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    UserControl = false
                };

                books = app.Workbooks;
                workbook = books.Open(
                    Filename: workbookPath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                sheet = FindWorksheet(workbook, sheetName);
                if (sheet == null)
                    throw new InvalidOperationException($"Worksheet '{sheetName}' not found.");

                // Validate headers once
                string[] expectedHeaders = { "UniqueName", "X", "Y", "Z" };
                for (int i = 0; i < ColumnCount; i++)
                {
                    Excel.Range h = null;
                    try
                    {
                        h = sheet.Cells[StartRow - 1, StartColumn + i];
                        string got = ReadText(h?.Value2);
                        if (!string.Equals(got, expectedHeaders[i], StringComparison.OrdinalIgnoreCase))
                        {
                            char col = (char)('A' + StartColumn + i - 1);
                            throw new InvalidOperationException($"Expected header '{expectedHeaders[i]}' in column {col}.");
                        }
                    }
                    finally { ExcelHelpers.ReleaseCom(h); }
                }

                // Determine last candidate row by taking max End(xlUp) across B..E
                int lastRowB = sheet.Cells[sheet.Rows.Count, StartColumn + 0].End(Excel.XlDirection.xlUp).Row;
                int lastRowC = sheet.Cells[sheet.Rows.Count, StartColumn + 1].End(Excel.XlDirection.xlUp).Row;
                int lastRowD = sheet.Cells[sheet.Rows.Count, StartColumn + 2].End(Excel.XlDirection.xlUp).Row;
                int lastRowE = sheet.Cells[sheet.Rows.Count, StartColumn + 3].End(Excel.XlDirection.xlUp).Row;
                int lastRow = Math.Max(StartRow, Math.Max(Math.Max(lastRowB, lastRowC), Math.Max(lastRowD, lastRowE)));
                int candidateCount = Math.Max(0, lastRow - StartRow + 1);

                // Initialize progress for candidate rows
                progress?.Invoke(0, candidateCount, BuildExcelStatus(0, candidateCount));

                if (candidateCount == 0)
                {
                    // Still do a completed state for UI clarity
                    progress?.Invoke(0, 0, "Reading Excel (0 rows).");
                    progress?.Invoke(0, 0, "Reading Excel complete.");
                    return rows;
                }

                int processedCandidates = 0;

                // Stream by chunks to keep UI responsive and COM calls low
                for (int chunkStartOffset = 0; chunkStartOffset < candidateCount; chunkStartOffset += ExcelChunkSize)
                {
                    int thisCount = Math.Min(ExcelChunkSize, candidateCount - chunkStartOffset);
                    int r1 = StartRow + chunkStartOffset;
                    int r2 = r1 + thisCount - 1;

                    Excel.Range dataRange = sheet.Range[
                        sheet.Cells[r1, StartColumn],
                        sheet.Cells[r2, StartColumn + ColumnCount - 1]
                    ];

                    object[,] data = (object[,])dataRange.Value2;
                    ExcelHelpers.ReleaseCom(dataRange);

                    // Iterate rows inside this chunk
                    for (int i = 1; i <= thisCount; i++)
                    {
                        string name = ReadText(data[i, 1]);
                        double? x = ReadNullableDouble(data[i, 2]);
                        double? y = ReadNullableDouble(data[i, 3]);
                        double? z = ReadNullableDouble(data[i, 4]);

                        bool allEmpty = string.IsNullOrWhiteSpace(name) && x == null && y == null && z == null;
                        if (!allEmpty)
                        {
                            rows.Add(new PointRow { UniqueName = name, X = x, Y = y, Z = z });
                        }

                        processedCandidates++;
                        progress?.Invoke(processedCandidates, candidateCount, BuildExcelStatus(processedCandidates, candidateCount));
                    }
                }

                // Final tick to ensure the bar reaches 100%
                progress?.Invoke(candidateCount, candidateCount, "Reading Excel complete.");
            }
            finally
            {
                try { workbook?.Close(false); } catch { }
                ExcelHelpers.ReleaseCom(sheet);
                ExcelHelpers.ReleaseCom(workbook);
                ExcelHelpers.ReleaseCom(books);
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    ExcelHelpers.ReleaseCom(app);
                }
            }

            return rows;
        }

        // ======= Worksheet find =======
        private static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null) return null;
            Excel.Sheets sheets = null;
            Excel.Worksheet target = null;

            try
            {
                sheets = workbook.Worksheets;
                int count = sheets?.Count ?? 0;

                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet candidate = null;
                    try
                    {
                        candidate = sheets[i] as Excel.Worksheet;
                        if (candidate != null && string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            target = candidate;
                            candidate = null; // ownership transferred to caller
                            break;
                        }
                    }
                    finally
                    {
                        if (candidate != null) ExcelHelpers.ReleaseCom(candidate);
                    }
                }
            }
            finally
            {
                ExcelHelpers.ReleaseCom(sheets);
            }

            return target;
        }

        // ======= UI status strings =======
        private static string BuildExcelStatus(int processed, int total)
        {
            if (total <= 0) return "Reading Excel (0 rows)";
            int clamped = Math.Max(0, Math.Min(processed, total));
            double percent = total == 0 ? 0.0 : (clamped / (double)total) * 100.0;
            return $"Reading Excel {clamped} of {total} row(s) ({percent:0.##}%).";
        }

        private static string BuildAssignmentStatus(int processed, int total)
        {
            if (total <= 0) return "Updating ETABS points (0 rows)";
            int clamped = Math.Max(0, Math.Min(processed, total));
            double percent = total == 0 ? 0.0 : (clamped / (double)total) * 100.0;
            return $"Updating ETABS points {clamped} of {total} row(s) ({percent:0.##}%).";
        }

        // ======= Build GH tree for preview/debug =======
        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(List<PointRow> rows)
        {
            var tree = new GH_Structure<GH_ObjectWrapper>();
            GH_Path uniquePath = new GH_Path(0);
            GH_Path xPath = new GH_Path(1);
            GH_Path yPath = new GH_Path(2);
            GH_Path zPath = new GH_Path(3);

            foreach (PointRow row in rows)
            {
                tree.Append(new GH_ObjectWrapper(row.UniqueName), uniquePath);
                tree.Append(new GH_ObjectWrapper(row.X), xPath);
                tree.Append(new GH_ObjectWrapper(row.Y), yPath);
                tree.Append(new GH_ObjectWrapper(row.Z), zPath);
            }
            return tree;
        }

        /// <summary>
        /// Create-if-missing (always on), then rename (case), move by delta,
        /// and delete baseline-only points with no connectivity.
        /// </summary>
        private static void ProcessRows_CreateMoveDelete(
            cSapModel sapModel,
            List<PointRow> rows,
            PointBaseline baseline,
            HashSet<string> existingNames,
            bool createIfMissing,
            List<string> actions,
            List<string> messages,
            Action<int, int, string> progress)
        {
            if (sapModel.SelectObj != null) sapModel.SelectObj.ClearSelection();

            int renameCount = 0, moveCount = 0, deleteCount = 0, createCount = 0;
            int processed = 0, total = rows?.Count ?? 0;
            var namesToKeep = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AdvanceProgress()
            {
                processed++;
                progress?.Invoke(processed, total, BuildAssignmentStatus(processed, total));
            }

            for (int index = 0; index < rows.Count; index++)
            {
                PointRow row = rows[index];
                string desiredName = row.UniqueName?.Trim();
                string baselineName = baseline.HasData ? baseline.GetUniqueName(index) : string.Empty;
                string rowLabel = $"Row {StartRow + index}";

                if (string.IsNullOrWhiteSpace(desiredName))
                {
                    MarkNameForKeep(namesToKeep, baselineName);
                    messages.Add($"{rowLabel}: UniqueName is empty. Skipped.");
                    AdvanceProgress();
                    continue;
                }

                // Determine if point exists already (case-insensitive)
                bool exists = existingNames.Contains(desiredName);
                string workingName = exists ? desiredName : baselineName;

                if (!exists && !string.IsNullOrWhiteSpace(workingName))
                    exists = existingNames.Contains(workingName);

                // If the point does not exist → create (always enabled here)
                if (!exists)
                {
                    if (!createIfMissing)
                    {
                        messages.Add($"{rowLabel}: Point '{desiredName}' not found and createIfMissing=false. Skipped.");
                        AdvanceProgress();
                        continue;
                    }

                    // Need absolute coordinates to add a point
                    if (row.X == null || row.Y == null || row.Z == null)
                    {
                        messages.Add($"{rowLabel}: Cannot create '{desiredName}' without X, Y, Z.");
                        AdvanceProgress();
                        continue;
                    }

                    string newName = desiredName; // ETABS may override this
                    int addRet = sapModel.PointObj.AddCartesian(row.X.Value, row.Y.Value, row.Z.Value, ref newName);
                    if (addRet == 0)
                    {
                        createCount++;
                        actions.Add($"{rowLabel}: Created point '{newName}' at ({row.X:0.###}, {row.Y:0.###}, {row.Z:0.###}).");

                        // If ETABS gave an auto name and it's different → try to rename to Excel's desiredName
                        if (!string.Equals(newName, desiredName, StringComparison.Ordinal))
                        {
                            int rn = sapModel.PointObj.ChangeName(newName, desiredName);
                            if (rn == 0)
                            {
                                actions.Add($"{rowLabel}: Renamed '{newName}' -> '{desiredName}'.");
                                existingNames.Add(desiredName);
                                workingName = desiredName;
                            }
                            else
                            {
                                messages.Add($"{rowLabel}: ChangeName failed after create (code {rn}). Kept '{newName}'.");
                                existingNames.Add(newName);
                                workingName = newName;
                            }
                        }
                        else
                        {
                            existingNames.Add(desiredName);
                            workingName = desiredName;
                        }

                        MarkNameForKeep(namesToKeep, workingName);
                        AdvanceProgress();
                        continue; // created at target XYZ already → no move needed
                    }
                    else
                    {
                        messages.Add($"{rowLabel}: AddCartesian failed for '{desiredName}' (code {addRet}).");
                        AdvanceProgress();
                        continue;
                    }
                }

                // At this point the point exists; sync exact casing if needed (baseline→desired)
                if (!string.IsNullOrWhiteSpace(baselineName) &&
                    !string.Equals(desiredName, baselineName, StringComparison.Ordinal) &&
                    existingNames.Contains(baselineName))
                {
                    int ret = sapModel.PointObj.ChangeName(baselineName, desiredName);
                    if (ret == 0)
                    {
                        renameCount++;
                        actions.Add($"{rowLabel}: Renamed '{baselineName}' -> '{desiredName}'.");
                        existingNames.Remove(baselineName);
                        existingNames.Add(desiredName);
                        workingName = desiredName;
                    }
                    else
                    {
                        // Keep workingName as baselineName if rename failed
                        workingName = existingNames.Contains(desiredName) ? desiredName : baselineName;
                        messages.Add($"{rowLabel}: ChangeName failed (code {ret}).");
                    }
                }
                else
                {
                    // Working name preference: use desired if exists, else baseline
                    workingName = existingNames.Contains(desiredName) ? desiredName : baselineName;
                }

                MarkNameForKeep(namesToKeep, workingName);

                // Read current coordinates
                if (!TryGetPoint(sapModel, workingName, out double currentX, out double currentY, out double currentZ))
                {
                    messages.Add($"{rowLabel}: Cannot read current coordinates for '{workingName}'.");
                    AdvanceProgress();
                    continue;
                }

                // Resolve target coordinates (null => keep as-is)
                double targetX = row.X ?? currentX;
                double targetY = row.Y ?? currentY;
                double targetZ = row.Z ?? currentZ;

                double dx = targetX - currentX;
                double dy = targetY - currentY;
                double dz = targetZ - currentZ;

                // Skip move under tolerance
                if (Math.Abs(dx) < Tolerance && Math.Abs(dy) < Tolerance && Math.Abs(dz) < Tolerance)
                {
                    AdvanceProgress();
                    continue;
                }

                int selectRet = sapModel.PointObj.SetSelected(workingName, true);
                if (selectRet != 0)
                {
                    messages.Add($"{rowLabel}: Failed to select '{workingName}' (code {selectRet}).");
                    AdvanceProgress();
                    continue;
                }

                try
                {
                    int moveRet = sapModel.EditGeneral.Move(dx, dy, dz);
                    if (moveRet == 0)
                    {
                        moveCount++;
                        actions.Add($"{rowLabel}: Moved '{workingName}' to ({targetX:0.###}, {targetY:0.###}, {targetZ:0.###}).");
                    }
                    else
                    {
                        messages.Add($"{rowLabel}: Move failed for '{workingName}' (code {moveRet}).");
                    }
                }
                finally
                {
                    sapModel.PointObj.SetSelected(workingName, false);
                }

                AdvanceProgress();
            }

            if (sapModel.SelectObj != null) sapModel.SelectObj.ClearSelection();

            // Delete points present in baseline but not kept by this pass (and with no connectivity)
            int deleted = DeleteMissingPoints(sapModel, baseline, namesToKeep, actions, messages, existingNames);
            deleteCount += deleted;

            // Summary
            messages.Add($"Processed {rows.Count} row(s).");
            if (createCount > 0) messages.Add($"Created {createCount} point(s).");
            if (renameCount > 0) messages.Add($"Renamed {renameCount} point(s).");
            if (moveCount > 0) messages.Add($"Moved {moveCount} point(s).");
            if (deleteCount > 0) messages.Add($"Deleted {deleteCount} point(s) removed from Excel.");

            progress?.Invoke(total, total, BuildAssignmentStatus(total, total));
        }

        private static void MarkNameForKeep(HashSet<string> namesToKeep, string name)
        {
            if (namesToKeep == null) return;
            if (string.IsNullOrWhiteSpace(name)) return;
            namesToKeep.Add(name.Trim());
        }

        private static int DeleteMissingPoints(
            cSapModel sapModel,
            PointBaseline baseline,
            HashSet<string> namesToKeep,
            List<string> actions,
            List<string> messages,
            HashSet<string> existingNames)
        {
            if (sapModel == null || baseline == null || !baseline.HasData) return 0;

            int deleteCount = 0;

            foreach (PointBaseline.Entry entry in baseline.Entries)
            {
                string name = entry.UniqueName;
                if (string.IsNullOrWhiteSpace(name)) continue;

                if (namesToKeep != null && namesToKeep.Contains(name)) continue;

                // Skip if point no longer exists
                if (!TryGetPoint(sapModel, name, out _, out _, out _)) continue;

                // Only delete isolated points (no connectivity)
                if (!TryCheckPointConnectivity(sapModel, name, out bool hasConnectivity))
                {
                    messages.Add($"Unable to determine connectivity for '{name}'. Skipped delete.");
                    MarkNameForKeep(namesToKeep, name);
                    continue;
                }

                if (hasConnectivity)
                {
                    messages.Add($"Skipped deleting '{name}' because it still has connectivity.");
                    MarkNameForKeep(namesToKeep, name);
                    continue;
                }

                int ret = sapModel.PointObj.DeleteSpecialPoint(name);
                if (ret == 0)
                {
                    deleteCount++;
                    actions.Add($"Deleted point '{name}' because it was removed from Excel.");
                    existingNames?.Remove(name);
                }
                else
                {
                    messages.Add($"Failed to delete point '{name}' (code {ret}).");
                    MarkNameForKeep(namesToKeep, name);
                }
            }

            return deleteCount;
        }

        private static bool TryCheckPointConnectivity(cSapModel model, string name, out bool hasConnectivity)
        {
            hasConnectivity = false;
            if (model == null || string.IsNullOrWhiteSpace(name)) return false;

            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int[] pointNumbers = null;

            try
            {
                int ret = model.PointObj.GetConnectivity(name, ref numberItems, ref objectTypes, ref objectNames, ref pointNumbers);
                if (ret != 0) return false;

                hasConnectivity = numberItems > 0;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetPoint(cSapModel model, string name, out double x, out double y, out double z)
        {
            x = 0; y = 0; z = 0;
            if (model == null || string.IsNullOrWhiteSpace(name)) return false;
            int ret = model.PointObj.GetCoordCartesian(name, ref x, ref y, ref z);
            return ret == 0;
        }

        private static HashSet<string> GetPointNames(cSapModel model)
        {
            try
            {
                int count = 0;
                string[] names = null;
                int ret = model.PointObj.GetNameList(ref count, ref names);
                if (ret != 0) return null;

                var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (names == null) return result;

                foreach (string name in names)
                {
                    if (!string.IsNullOrWhiteSpace(name))
                        result.Add(name.Trim());
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        private int TriggerGetPointInfoRefresh()
        {
            GH_Document document = OnPingDocument();
            if (document == null) return 0;

            var targets = new List<GhcGetPointInfo>();
            foreach (IGH_DocumentObject obj in document.Objects)
            {
                if (obj is GhcGetPointInfo component && !component.Locked && !component.Hidden)
                    targets.Add(component);
            }

            if (targets.Count == 0) return 0;

            document.ScheduleSolution(5, _ =>
            {
                foreach (GhcGetPointInfo component in targets)
                {
                    if (!component.Locked && !component.Hidden)
                        component.ExpireSolution(false);
                }
            });

            return targets.Count;
        }

        // ======= Excel value helpers =======
        private static string ReadText(object value)
        {
            if (value == null) return string.Empty;
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return text.Trim();
        }

        private static double? ReadNullableDouble(object value)
        {
            if (value == null) return null;
            if (value is double direct) return direct;

            string text = ReadText(value);
            if (string.IsNullOrEmpty(text)) return null;

            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double result))
                return result;

            return null;
        }

        // ======= Data structures =======
        private class PointRow
        {
            public string UniqueName { get; set; }
            public double? X { get; set; }
            public double? Y { get; set; }
            public double? Z { get; set; }
        }

        private class PointBaseline
        {
            // Preserves capture sequence; lets us fetch old names by index
            private readonly OrderedLookup<string, Entry> _orderedEntries =
                new OrderedLookup<string, Entry>(StringComparer.Ordinal);
            private bool _hasData;

            private PointBaseline() { }

            public static PointBaseline FromTree(GH_Structure<IGH_Goo> tree)
            {
                if (tree == null) return new PointBaseline();

                PointBaseline baseline = new PointBaseline();
                List<string> names = new List<string>(ReadStringBranch(tree, 0));
                List<double?> xValues = new List<double?>(ReadDoubleBranch(tree, 1));
                List<double?> yValues = new List<double?>(ReadDoubleBranch(tree, 2));
                List<double?> zValues = new List<double?>(ReadDoubleBranch(tree, 3));

                int rowCount = Math.Max(Math.Max(names.Count, xValues.Count), Math.Max(yValues.Count, zValues.Count));
                for (int i = 0; i < rowCount; i++)
                {
                    string name = i < names.Count ? names[i] ?? string.Empty : string.Empty;
                    double? x = i < xValues.Count ? xValues[i] : null;
                    double? y = i < yValues.Count ? yValues[i] : null;
                    double? z = i < zValues.Count ? zValues[i] : null;

                    Entry entry = new Entry(name ?? string.Empty, x, y, z);
                    string lookupKey = string.IsNullOrWhiteSpace(entry.UniqueName) ? null : entry.UniqueName;

                    // Instance fields must be qualified with 'baseline.'
                    baseline._orderedEntries.Add(lookupKey, entry);
                }

                baseline._hasData = tree.DataCount > 0;
                return baseline;
            }

            public bool HasData => _hasData;
            public int Count => _orderedEntries.Count;
            public IReadOnlyList<Entry> Entries => _orderedEntries.Entries;

            public string GetUniqueName(int index)
            {
                var entries = _orderedEntries.Entries;
                if (index < 0 || index >= entries.Count) return string.Empty;
                return entries[index].UniqueName ?? string.Empty;
            }

            public bool TryGetEntry(string uniqueName, out Entry entry)
            {
                string lookupKey = string.IsNullOrWhiteSpace(uniqueName) ? null : uniqueName;
                return _orderedEntries.TryGetValue(lookupKey, out entry);
            }

            private static IEnumerable<string> ReadStringBranch(GH_Structure<IGH_Goo> tree, int index)
            {
                if (index < 0 || index >= tree.PathCount) yield break;
                IList branch = tree.get_Branch(index);
                if (branch == null) yield break;

                foreach (object item in branch)
                {
                    IGH_Goo goo = item as IGH_Goo;
                    if (goo == null) yield return string.Empty;
                    else if (GH_Convert.ToString(goo, out string text, GH_Conversion.Both))
                        yield return string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
                    else yield return string.Empty;
                }
            }

            private static IEnumerable<double?> ReadDoubleBranch(GH_Structure<IGH_Goo> tree, int index)
            {
                if (index < 0 || index >= tree.PathCount) yield break;
                IList branch = tree.get_Branch(index);
                if (branch == null) yield break;

                foreach (object item in branch)
                {
                    IGH_Goo goo = item as IGH_Goo;
                    if (goo == null) yield return null;
                    else if (GH_Convert.ToDouble(goo, out double value, GH_Conversion.Both))
                        yield return value;
                    else yield return null;
                }
            }

            internal readonly record struct Entry(string UniqueName, double? X, double? Y, double? Z);
        }
    }
}
