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
    /// Reads point information from Excel and pushes the data back to ETABS.
    /// The class tries to be as small and as clear as possible so new developers
    /// can follow the logic without jumping between many helper methods.
    /// </summary>
    public class GhcSetPointInfo : GH_Component
    {
        private const string DefaultSheet = "Point Info";
        private const int StartColumn = 2;          // Column B
        private const int StartRow = 2;             // Header lives on row 1
        private const int ColumnCount = 4;          // UniqueName, X, Y, Z
        private const double Tolerance = 1e-6;
        private static readonly int Grey = ColorTranslator.ToOle(Color.LightGray);
        private static readonly int White = ColorTranslator.ToOle(Color.White);

        private bool _lastRun;
        private GH_Structure<GH_ObjectWrapper> _lastValues = new GH_Structure<GH_ObjectWrapper>();
        private readonly List<string> _lastActions = new List<string>();
        private readonly List<string> _lastMessages = new List<string> { "No previous run. Toggle 'run' to execute." };

        public GhcSetPointInfo()
            : base(
                "Set Point Info",
                "SetPointInfo",
                "Rename and move ETABS point objects by reading an Excel worksheet. The component also highlights edited cells so users can see what changed.",
                "MGT",
                "2.0 Point Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("A9B6F07F-7D5E-4A25-AD2A-6F0A7AE12C47");

        protected override System.Drawing.Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager pManager)
        {
            pManager.AddBooleanParameter("run", "run", "Rising edge trigger.", GH_ParamAccess.item, false);
            pManager.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            pManager.AddTextParameter("excelPath", "excelPath", "Full or project relative path to the workbook.", GH_ParamAccess.item, string.Empty);
            pManager.AddTextParameter("sheetName", "sheetName", "Worksheet name containing the data.", GH_ParamAccess.item, DefaultSheet);

            int baselineIndex = pManager.AddGenericParameter(
                "baseline",
                "baseline",
                "Optional tree captured from GhcGetPointInfo. It is used to highlight edited cells and to know the old point names.",
                GH_ParamAccess.tree);
            pManager[baselineIndex].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager pManager)
        {
            pManager.AddGenericParameter("values", "values", "Column wise data tree (UniqueName / X / Y / Z).", GH_ParamAccess.tree);
            pManager.AddTextParameter("actions", "actions", "Rename and move actions executed on ETABS points.", GH_ParamAccess.list);
            pManager.AddTextParameter("messages", "messages", "Summary and diagnostic messages.", GH_ParamAccess.list);
        }

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

            bool rising = !_lastRun && run;
            if (!rising)
            {
                da.SetDataTree(0, _lastValues.Duplicate());
                da.SetDataList(1, _lastActions.ToArray());
                da.SetDataList(2, _lastMessages.ToArray());
                _lastRun = run;
                return;
            }

            List<string> messages = new List<string>();
            List<string> actions = new List<string>();
            GH_Structure<GH_ObjectWrapper> valueTree = new GH_Structure<GH_ObjectWrapper>();

            try
            {
                if (sapModel == null)
                {
                    throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");
                }

                string fullPath = ExcelHelpers.ProjectRelative(excelPath);
                if (string.IsNullOrWhiteSpace(fullPath))
                {
                    throw new InvalidOperationException("excelPath is empty.");
                }

                if (!File.Exists(fullPath))
                {
                    throw new FileNotFoundException("Excel workbook not found.", fullPath);
                }

                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    sheetName = DefaultSheet;
                }

                UiHelpers.ShowDualProgressBar(
                    "Set Point Info",
                    "Reading Excel...",
                    0,
                    "Updating points...",
                    0);

                PointBaseline baseline = PointBaseline.FromTree(baselineTree);

                int highlightedCells;
                List<PointRow> rows = ReadExcelRows(
                    fullPath,
                    sheetName,
                    baseline,
                    out highlightedCells,
                    (current, total, status) => UiHelpers.UpdateExcelProgressBar(current, total, status));

                valueTree = BuildValueTree(rows);

                int totalRows = rows.Count;
                if (totalRows == 0)
                {
                    messages.Add("Excel sheet contained no data rows.");
                }
                else
                {
                    EnsureModelUnlocked(sapModel);

                    HashSet<string> existingNames = GetPointNames(sapModel);
                    if (existingNames == null)
                    {
                        throw new InvalidOperationException("Failed to read point names from ETABS.");
                    }

                    UiHelpers.UpdateAssignmentProgressBar(
                        0,
                        totalRows,
                        UiHelpers.FormatProgressStatus(0, totalRows, "Updating ETABS points", "row", "rows"));

                    ProcessRows(
                        sapModel,
                        rows,
                        baseline,
                        existingNames,
                        actions,
                        messages,
                        (current, total, status) => UiHelpers.UpdateAssignmentProgressBar(current, total, status));

                    UiHelpers.UpdateAssignmentProgressBar(
                        totalRows,
                        totalRows,
                        UiHelpers.FormatProgressStatus(totalRows, totalRows, "Updating ETABS points", "row", "rows"));
                }

                if (highlightedCells > 0)
                {
                    messages.Add($"Highlighted {highlightedCells} cell(s) in Excel.");
                }

                try
                {
                    sapModel.View.RefreshView(0, false);
                }
                catch
                {
                    // Not fatal if refresh fails.
                }

                int refreshed = TriggerGetPointInfoRefresh();
                if (refreshed > 0)
                {
                    messages.Add($"Scheduled refresh for {refreshed} Get Point Info component(s).");
                }
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

        private static List<PointRow> ReadExcelRows(
            string workbookPath,
            string sheetName,
            PointBaseline baseline,
            out int highlightedCells,
            Action<int, int, string> progress)
        {
            // Snapshot the worksheet into PointRow records while remembering
            // which cells differ from the baseline capture.
            highlightedCells = 0;
            List<PointRow> rows = new List<PointRow>();

            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            Excel.Range usedRange = null;

            bool workbookModified = false;

            try
            {
                // Run Excel invisibly so we can safely read/annotate the file
                // without flashing windows at the user.
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    UserControl = false
                };

                books = app.Workbooks;
                // Open the workbook directly which lets us recolor cells when
                // highlighting rows that changed from the baseline snapshot.
                workbook = books.Open(
                    Filename: workbookPath,
                    UpdateLinks: 0,
                    ReadOnly: false,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                sheet = ExcelHelpers.FindWorksheet(workbook, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Worksheet '{sheetName}' not found.");
                }

                string[] expectedHeaders = { "UniqueName", "X", "Y", "Z" };
                for (int i = 0; i < ColumnCount; i++)
                {
                    Excel.Range headerCell = null;
                    try
                    {
                        headerCell = sheet.Cells[StartRow - 1, StartColumn + i];
                        string text = ReadText(headerCell?.Value2);
                        if (!string.Equals(text, expectedHeaders[i], StringComparison.OrdinalIgnoreCase))
                        {
                            // Abort immediately if someone renamed or moved a
                            // header; downstream parsing would be unreliable.
                            char columnLetter = (char)('A' + StartColumn + i - 1);
                            throw new InvalidOperationException($"Expected header '{expectedHeaders[i]}' in column {columnLetter}.");
                        }
                    }
                    finally
                    {
                        ExcelHelpers.ReleaseCom(headerCell);
                    }
                }

                usedRange = sheet.UsedRange;
                int lastRow = Math.Max(StartRow, (usedRange?.Row ?? StartRow) + (usedRange?.Rows?.Count ?? 0) - 1);
                int totalRows = Math.Max(0, lastRow - StartRow + 1);

                // Surface progress to the component UI so long-running imports
                // show signs of life.
                progress?.Invoke(
                    0,
                    totalRows,
                    UiHelpers.FormatProgressStatus(0, totalRows, "Reading Excel", "row", "rows"));

                for (int rowIndex = 0; rowIndex < totalRows; rowIndex++)
                {
                    int excelRow = StartRow + rowIndex;
                    progress?.Invoke(
                        rowIndex,
                        totalRows,
                        UiHelpers.FormatProgressStatus(rowIndex, totalRows, "Reading Excel", "row", "rows"));

                    string name = string.Empty;
                    double? x = null;
                    double? y = null;
                    double? z = null;

                    for (int columnIndex = 0; columnIndex < ColumnCount; columnIndex++)
                    {
                        Excel.Range cell = null;
                        try
                        {
                            cell = sheet.Cells[excelRow, StartColumn + columnIndex];
                            object value = cell?.Value2;

                            if (baseline.HasData)
                            {
                                // Paint differences compared to the saved
                                // baseline so users can spot edits quickly.
                                bool changed = baseline.IsDifferent(columnIndex, rowIndex, value);
                                bool colorChanged = UpdateCellColor(cell, changed);
                                if (colorChanged)
                                {
                                    workbookModified = true;
                                    if (changed)
                                    {
                                        highlightedCells++;
                                    }
                                }
                            }

                            switch (columnIndex)
                            {
                                case 0:
                                    name = ReadText(value);
                                    break;
                                case 1:
                                    x = ReadNullableDouble(value);
                                    break;
                                case 2:
                                    y = ReadNullableDouble(value);
                                    break;
                                case 3:
                                    z = ReadNullableDouble(value);
                                    break;
                            }
                        }
                        finally
                        {
                            ExcelHelpers.ReleaseCom(cell);
                        }
                    }

                    if (string.IsNullOrWhiteSpace(name) && x == null && y == null && z == null)
                    {
                        continue;
                    }

                    rows.Add(new PointRow
                    {
                        UniqueName = name,
                        X = x,
                        Y = y,
                        Z = z
                    });
                }

                progress?.Invoke(
                    totalRows,
                    totalRows,
                    UiHelpers.FormatProgressStatus(totalRows, totalRows, "Reading Excel", "row", "rows"));

                if (workbookModified)
                {
                    workbook.Save();
                }
            }
            finally
            {
                try { workbook?.Close(false); } catch { }
                ExcelHelpers.ReleaseCom(usedRange);
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

        private static bool UpdateCellColor(Excel.Range cell, bool changed)
        {
            if (cell == null)
            {
                return false;
            }

            Excel.Interior interior = null;
            try
            {
                interior = cell.Interior;
                if (interior == null)
                {
                    return false;
                }

                int targetColor = changed ? Grey : White;
                interior.Pattern = Excel.XlPattern.xlPatternSolid;
                interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;

                object current = interior.Color;
                int currentInt = current is double d ? Convert.ToInt32(d) : current is int i ? i : -1;
                if (currentInt == targetColor)
                {
                    return false;
                }

                interior.Color = targetColor;
                return true;
            }
            finally
            {
                ExcelHelpers.ReleaseCom(interior);
            }
        }

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(List<PointRow> rows)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();

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

        private static void ProcessRows(
            cSapModel sapModel,
            List<PointRow> rows,
            PointBaseline baseline,
            HashSet<string> existingNames,
            List<string> actions,
            List<string> messages,
            Action<int, int, string> progress)
        {
            if (sapModel.SelectObj != null)
            {
                sapModel.SelectObj.ClearSelection();
            }

            int renameCount = 0;
            int moveCount = 0;
            int deleteCount = 0;
            int processed = 0;
            int total = rows?.Count ?? 0;

            HashSet<string> namesToKeep = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AdvanceProgress()
            {
                processed++;
                // Previously we never advanced the assignment-side progress bar when
                // rows were skipped or bailed out early, so the UI stayed frozen even
                // though the loop kept working. Publishing the progress after every
                // row (including skipped ones) keeps the bar in sync with the work.
                progress?.Invoke(
                    processed,
                    total,
                    UiHelpers.FormatProgressStatus(processed, total, "Updating ETABS points", "row", "rows"));
            }

            // Walk the Excel rows in order so operations respect the worksheet
            // ordering and the baseline's recorded indices.
            for (int index = 0; index < rows.Count; index++)
            {
                PointRow row = rows[index];
                string desiredName = row.UniqueName?.Trim();
                string baselineName = baseline.HasData ? baseline.GetUniqueName(index) : string.Empty;
                string workingName = DetermineWorkingName(desiredName, baselineName, existingNames);
                string rowLabel = $"Row {StartRow + index}";

                if (string.IsNullOrWhiteSpace(desiredName))
                {
                    MarkNameForKeep(namesToKeep, baselineName);
                    messages.Add($"{rowLabel}: UniqueName is empty. Skipped.");
                    AdvanceProgress();
                    continue;
                }

                if (!existingNames.Contains(workingName))
                {
                    MarkNameForKeep(namesToKeep, baselineName);
                    messages.Add($"{rowLabel}: Point '{workingName}' not found in ETABS. Skipped.");
                    AdvanceProgress();
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(baselineName) &&
                    !string.Equals(desiredName, baselineName, StringComparison.Ordinal) &&
                    string.Equals(workingName, baselineName, StringComparison.OrdinalIgnoreCase))
                {
                    // When names match ignoring case, push the rename through
                    // ETABS so unique name casing mirrors Excel exactly.
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
                        messages.Add($"{rowLabel}: ChangeName failed with code {ret}. Keeping old name '{baselineName}'.");
                    }
                }

                MarkNameForKeep(namesToKeep, workingName);

                if (!TryGetPoint(sapModel, workingName, out double currentX, out double currentY, out double currentZ))
                {
                    messages.Add($"{rowLabel}: Unable to read current coordinates for '{workingName}'.");
                    AdvanceProgress();
                    continue;
                }

                double targetX = row.X ?? currentX;
                double targetY = row.Y ?? currentY;
                double targetZ = row.Z ?? currentZ;

                double dx = targetX - currentX;
                double dy = targetY - currentY;
                double dz = targetZ - currentZ;

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

                int moveRet = sapModel.EditGeneral.Move(dx, dy, dz);
                sapModel.PointObj.SetSelected(workingName, false);

                if (moveRet == 0)
                {
                    moveCount++;
                    actions.Add($"{rowLabel}: Moved '{workingName}' to ({targetX:0.###}, {targetY:0.###}, {targetZ:0.###}).");
                }
                else
                {
                    messages.Add($"{rowLabel}: Move failed for '{workingName}' (code {moveRet}).");
                }

                AdvanceProgress();
            }

            if (sapModel.SelectObj != null)
            {
                sapModel.SelectObj.ClearSelection();
            }

            deleteCount += DeleteMissingPoints(sapModel, baseline, namesToKeep, actions, messages, existingNames);

            messages.Add($"Processed {rows.Count} row(s).");
            if (renameCount > 0)
            {
                messages.Add($"Renamed {renameCount} point(s).");
            }
            if (moveCount > 0)
            {
                messages.Add($"Moved {moveCount} point(s).");
            }
            if (deleteCount > 0)
            {
                messages.Add($"Deleted {deleteCount} point(s) removed from Excel.");
            }
        }

        private static string DetermineWorkingName(string desiredName, string baselineName, HashSet<string> existingNames)
        {
            // Try to use the requested Excel name when ETABS already knows it;
            // otherwise keep the baseline identifier for downstream deletes.
            if (!string.IsNullOrWhiteSpace(desiredName) && existingNames.Contains(desiredName))
            {
                return desiredName;
            }

            if (!string.IsNullOrWhiteSpace(baselineName) && existingNames.Contains(baselineName))
            {
                // When Excel renamed the point but ETABS still tracks the old
                // unique name, operate on that baseline identifier.
                return baselineName;
            }

            if (!string.IsNullOrWhiteSpace(desiredName))
            {
                return desiredName;
            }

            return baselineName ?? string.Empty;
        }

        private static void MarkNameForKeep(HashSet<string> namesToKeep, string name)
        {
            if (namesToKeep == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

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
            if (sapModel == null || baseline == null || !baseline.HasData)
            {
                return 0;
            }

            int deleteCount = 0;

            foreach (PointBaseline.Entry entry in baseline.Entries)
            {
                string name = entry.UniqueName;
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                if (namesToKeep != null && namesToKeep.Contains(name))
                {
                    continue;
                }

                if (!TryGetPoint(sapModel, name, out _, out _, out _))
                {
                    continue;
                }

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

            if (model == null || string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            int numberItems = 0;
            int[] objectTypes = null;
            string[] objectNames = null;
            int[] pointNumbers = null;

            try
            {
                int ret = model.PointObj.GetConnectivity(name, ref numberItems, ref objectTypes, ref objectNames, ref pointNumbers);
                if (ret != 0)
                {
                    return false;
                }

                hasConnectivity = numberItems > 0;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                // Arrays are managed; nothing to release.
            }
        }

        private static bool TryGetPoint(cSapModel model, string name, out double x, out double y, out double z)
        {
            x = 0;
            y = 0;
            z = 0;

            if (model == null || string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

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
                if (ret != 0)
                {
                    return null;
                }

                HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (names == null)
                {
                    return result;
                }

                foreach (string name in names)
                {
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        result.Add(name.Trim());
                    }
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
            if (document == null)
            {
                return 0;
            }

            List<GhcGetPointInfo> targets = new List<GhcGetPointInfo>();
            foreach (IGH_DocumentObject obj in document.Objects)
            {
                if (obj is GhcGetPointInfo component && !component.Locked && !component.Hidden)
                {
                    targets.Add(component);
                }
            }

            if (targets.Count == 0)
            {
                return 0;
            }

            document.ScheduleSolution(5, _ =>
            {
                foreach (GhcGetPointInfo component in targets)
                {
                    if (!component.Locked && !component.Hidden)
                    {
                        component.ExpireSolution(false);
                    }
                }
            });

            return targets.Count;
        }

        private static string ReadText(object value)
        {
            // Normalize any COM-provided value to a trimmed invariant string.
            if (value == null)
            {
                return string.Empty;
            }

            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return text.Trim();
        }

        private static double? ReadNullableDouble(object value)
        {
            // Interpret Excel cell contents as doubles when possible while
            // keeping empty cells as null.
            if (value == null)
            {
                return null;
            }

            if (value is double direct)
            {
                return direct;
            }

            string text = ReadText(value);
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }

            return null;
        }

        private class PointRow
        {
            public string UniqueName { get; set; }
            public double? X { get; set; }
            public double? Y { get; set; }
            public double? Z { get; set; }
        }

        private class PointBaseline
        {
            // OrderedLookup preserves the capture sequence from the baseline
            // tree while still letting us fetch rows by UniqueName instantly.
            private readonly OrderedLookup<string, Entry> _orderedEntries =
                new OrderedLookup<string, Entry>(StringComparer.Ordinal);
            private bool _hasData;

            private PointBaseline()
            {
            }

            public static PointBaseline FromTree(GH_Structure<IGH_Goo> tree)
            {
                if (tree == null)
                {
                    return new PointBaseline();
                }

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
                    // Record every row in capture order so surviving points keep
                    // their original positions when we diff against Excel.
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
                IReadOnlyList<Entry> entries = _orderedEntries.Entries;

                if (index < 0 || index >= entries.Count)
                {
                    return string.Empty;
                }

                Entry entry = entries[index];
                return entry.UniqueName ?? string.Empty;
            }

            public bool TryGetEntry(string uniqueName, out Entry entry)
            {
                string lookupKey = string.IsNullOrWhiteSpace(uniqueName) ? null : uniqueName;
                return _orderedEntries.TryGetValue(lookupKey, out entry);
            }

            public bool IsDifferent(int columnIndex, int rowIndex, object excelValue)
            {
                if (!_hasData)
                {
                    return false;
                }

                IReadOnlyList<Entry> entries = _orderedEntries.Entries;
                Entry entry = rowIndex >= 0 && rowIndex < entries.Count ? entries[rowIndex] : default;

                switch (columnIndex)
                {
                    case 0:
                        string oldName = entry.UniqueName ?? string.Empty;
                        string newName = ReadText(excelValue);
                        return !string.Equals(oldName, newName, StringComparison.Ordinal);
                    case 1:
                        return CompareNumber(entry.X, excelValue);
                    case 2:
                        return CompareNumber(entry.Y, excelValue);
                    case 3:
                        return CompareNumber(entry.Z, excelValue);
                    default:
                        return false;
                }
            }

            private static bool CompareNumber(double? baselineValue, object excelValue)
            {
                // Evaluate whether the Excel numeric cell diverges from the
                // baseline while respecting the movement tolerance.
                double? excel = ReadNullableDouble(excelValue);

                if (baselineValue.HasValue && excel.HasValue)
                {
                    return Math.Abs(baselineValue.Value - excel.Value) > Tolerance;
                }

                return baselineValue.HasValue != excel.HasValue;
            }

            private static IEnumerable<string> ReadStringBranch(GH_Structure<IGH_Goo> tree, int index)
            {
                // Iterate a Grasshopper branch, normalizing each goo entry to a
                // trimmed string.
                if (index < 0 || index >= tree.PathCount)
                {
                    yield break;
                }

                IList branch = tree.get_Branch(index);
                if (branch == null)
                {
                    yield break;
                }

                foreach (object item in branch)
                {
                    IGH_Goo goo = item as IGH_Goo;
                    if (goo == null)
                    {
                        yield return string.Empty;
                    }
                    else if (GH_Convert.ToString(goo, out string text, GH_Conversion.Both))
                    {
                        yield return string.IsNullOrWhiteSpace(text) ? string.Empty : text.Trim();
                    }
                    else
                    {
                        yield return string.Empty;
                    }
                }
            }

            private static IEnumerable<double?> ReadDoubleBranch(GH_Structure<IGH_Goo> tree, int index)
            {
                // Iterate a branch and convert entries into nullable doubles.
                if (index < 0 || index >= tree.PathCount)
                {
                    yield break;
                }

                IList branch = tree.get_Branch(index);
                if (branch == null)
                {
                    yield break;
                }

                foreach (object item in branch)
                {
                    IGH_Goo goo = item as IGH_Goo;
                    if (goo == null)
                    {
                        yield return null;
                    }
                    else if (GH_Convert.ToDouble(goo, out double value, GH_Conversion.Both))
                    {
                        yield return value;
                    }
                    else
                    {
                        yield return null;
                    }
                }
            }
            internal readonly record struct Entry(string UniqueName, double? X, double? Y, double? Z);
        }
    }
}

