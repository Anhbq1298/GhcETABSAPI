// -------------------------------------------------------------
// Component : Set Frame Distributed Loads (relative, Excel-driven)
// Author    : Anh Bui (extended)
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM), Microsoft.Office.Interop.Excel
// Panel     : "ETABS API" / "2.0 Frame Object Modelling"
// -------------------------------------------------------------
// Inputs (ordered):
//   0) run         (bool, item)    Rising-edge trigger.
//   1) sapModel    (ETABSv1.cSapModel, item)  ETABS model from Attach component.
//   2) excelPath   (string, item)  Full/relative path to workbook containing the sheet.
//   3) sheetName   (string, item)  Worksheet name. Defaults to "Assigned Loads On Frames".
//   4) replaceMode (bool, item)    True = replace, False = add.
//
// Outputs:
//   0) values      (generic, tree) Column-wise branches (11) read from Excel (header row excluded).
//   1) messages    (string, list)  Summary + diagnostics.
//
// Behavior Notes:
//   • Reads columns B..L (11 columns) from the specified sheet; row 1 treated as headers.
//   • Converts Excel rows into a column-oriented GH_Structure for downstream inspection.
//   • Assigns distributed loads using FrameObj.SetLoadDistributed with IsRelativeDist = true.
//   • Attempts to unlock the model automatically before assignment.
//   • Direction codes 1..3 default to the Local coordinate system unless overridden in Excel.
//   • Distances are clamped to [0,1]; swapped when start > end.
//   • Any row with missing/invalid core data is reported as "skipped".
//   • When run is false or not toggled, the component replays the last output messages/tree.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using ETABSv1;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Excel = Microsoft.Office.Interop.Excel;

namespace GhcETABSAPI
{
    public class GhcSetLoadDistOnFrames_Rel2 : GH_Component
    {
        private bool _lastRun;
        private GH_Structure<GH_ObjectWrapper> _lastValues = new GH_Structure<GH_ObjectWrapper>();
        private readonly List<string> _lastMessages = new List<string> { "No previous run. Toggle 'run' to assign." };

        public GhcSetLoadDistOnFrames_Rel2()
          : base(
                "Set Frame Distributed Loads (Rel, Excel)",
                "SetFrameUDLRelXl",
                "Assign distributed loads to ETABS frame objects by reading an Excel worksheet.",
                "ETABS API",
                "2.0 Frame Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("6AB30F5A-AFE1-4C53-B83D-19F2E6A64506");

        protected override Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("run", "run", "Rising-edge trigger; executes when this turns True.", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            p.AddTextParameter("excelPath", "excelPath", "Full or project-relative path to the workbook.", GH_ParamAccess.item, string.Empty);
            p.AddTextParameter("sheetName", "sheetName", "Worksheet name containing the data.", GH_ParamAccess.item, "Assigned Loads On Frames");
            p.AddBooleanParameter("replaceMode", "replace", "True = replace existing values, False = add.", GH_ParamAccess.item, true);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddGenericParameter("values", "values", "Column-wise data tree (11 branches) read from Excel (header row excluded).", GH_ParamAccess.tree);
            p.AddTextParameter("messages", "messages", "Summary and diagnostic messages.", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sapModel = null;
            string excelPath = null;
            string sheetName = "Assigned Loads On Frames";
            bool replaceMode = true;

            da.GetData(0, ref run);
            da.GetData(1, ref sapModel);
            da.GetData(2, ref excelPath);
            da.GetData(3, ref sheetName);
            da.GetData(4, ref replaceMode);

            bool rising = !_lastRun && run;
            if (!rising)
            {
                da.SetDataTree(0, _lastValues.Duplicate());
                da.SetDataList(1, _lastMessages.ToArray());
                _lastRun = run;
                return;
            }

            List<string> messages = new List<string>();
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
                    sheetName = "Assigned Loads On Frames";
                }

                ExcelLoadData excelData = ReadExcelSheet(fullPath, sheetName);
                valueTree = BuildValueTree(excelData);

                if (excelData.RowCount == 0)
                {
                    messages.Add($"Read 0 data rows from sheet '{sheetName}'. Nothing to assign.");
                }
                else
                {
                    messages.Add($"Read {excelData.RowCount} data rows from sheet '{sheetName}'.");

                    EnsureModelUnlocked(sapModel);
                    HashSet<string> existingNames = TryGetExistingFrameNames(sapModel);

                    int assignedCount = 0;
                    int failedCount = 0;
                    List<string> failedPairs = new List<string>();
                    List<string> skippedPairs = new List<string>();

                    for (int i = 0; i < excelData.RowCount; i++)
                    {
                        string frameName = TrimOrEmpty(excelData.FrameName[i]);
                        string loadPattern = TrimOrEmpty(excelData.LoadPattern[i]);
                        int? rawType = excelData.MyType[i];
                        int? rawDirection = excelData.Direction[i];
                        double? rawRelDist1 = excelData.RelDist1[i];
                        double? rawRelDist2 = excelData.RelDist2[i];
                        double? rawVal1 = excelData.Value1[i];
                        double? rawVal2 = excelData.Value2[i];
                        string coordinateOverride = TrimOrEmpty(excelData.CoordinateSystem[i]);

                        if (string.IsNullOrEmpty(frameName) || string.IsNullOrEmpty(loadPattern) ||
                            !rawType.HasValue || !rawDirection.HasValue ||
                            !rawRelDist1.HasValue || !rawRelDist2.HasValue ||
                            !rawVal1.HasValue || !rawVal2.HasValue)
                        {
                            skippedPairs.Add($"{i}:{frameName}");
                            continue;
                        }

                        if (existingNames != null && !existingNames.Contains(frameName))
                        {
                            failedCount++;
                            failedPairs.Add($"{i}:{frameName}");
                            continue;
                        }

                        int loadType = NormalizeLoadType(rawType.Value);
                        int direction = ClampDirCode(rawDirection.Value);
                        double relDist1 = Clamp01(rawRelDist1.Value);
                        double relDist2 = Clamp01(rawRelDist2.Value);
                        double val1 = rawVal1.Value;
                        double val2 = rawVal2.Value;

                        if (IsInvalidNumber(relDist1) || IsInvalidNumber(relDist2) ||
                            IsInvalidNumber(val1) || IsInvalidNumber(val2))
                        {
                            skippedPairs.Add($"{i}:{frameName}");
                            continue;
                        }

                        if (relDist1 > relDist2)
                        {
                            double tmp = relDist1;
                            relDist1 = relDist2;
                            relDist2 = tmp;
                        }

                        string coordinateSystem = !string.IsNullOrEmpty(coordinateOverride)
                            ? coordinateOverride
                            : ((direction >= 1 && direction <= 3) ? "Local" : "Global");

                        int ret = sapModel.FrameObj.SetLoadDistributed(
                            frameName,
                            loadPattern,
                            loadType,
                            direction,
                            relDist1,
                            relDist2,
                            val1,
                            val2,
                            coordinateSystem,
                            true,
                            replaceMode,
                            (int)eItemType.Objects);

                        if (ret == 0)
                        {
                            assignedCount++;
                        }
                        else
                        {
                            failedCount++;
                            failedPairs.Add($"{i}:{frameName}");
                        }
                    }

                    messages.Add($"{Plural(assignedCount, "member")} successfully assigned, {Plural(failedCount, "member")} unsuccessful.");

                    if (failedPairs.Count > 0)
                    {
                        messages.Add("Unsuccessful members (0-based index:name): " + string.Join(", ", failedPairs));
                    }

                    if (skippedPairs.Count > 0)
                    {
                        messages.Add("Skipped members (0-based index:name): " + string.Join(", ", skippedPairs));
                    }

                    try
                    {
                        sapModel.View.RefreshView(0, false);
                    }
                    catch
                    {
                        // ignored
                    }
                }
            }
            catch (Exception ex)
            {
                messages.Add("Error: " + ex.Message);
            }

            da.SetDataTree(0, valueTree);
            da.SetDataList(1, messages.ToArray());

            _lastValues = valueTree.Duplicate();
            _lastMessages.Clear();
            _lastMessages.AddRange(messages);
            _lastRun = run;
        }

        private static ExcelLoadData ReadExcelSheet(string fullPath, string sheetName)
        {
            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range usedRange = null;

            try
            {
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    UserControl = false
                };

                books = app.Workbooks;
                wb = books.Open(
                    Filename: fullPath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                const string expectedSheetName = "Assigned Loads On Frames";
                if (!string.Equals(sheetName, expectedSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Invalid workbook: expected sheet name '{expectedSheetName}'.");
                }

                ws = FindWorksheet(wb, sheetName);
                if (ws == null)
                {
                    throw new InvalidOperationException($"Worksheet '{sheetName}' not found in '{Path.GetFileName(fullPath)}'.");
                }

                ExcelLoadData data = new ExcelLoadData();

                const int startColumn = 2; // Column B
                const int columnCount = 11;

                // Capture headers (row 1)
                string[] expectedHeaders =
                {
                    "FrameName",
                    "LoadPattern",
                    "Type",
                    "CoordinateSystem",
                    "Direction",
                    "RelDist1",
                    "RelDist2",
                    "Dist1",
                    "Dist2",
                    "Value1",
                    "Value2"
                };

                for (int col = 0; col < columnCount; col++)
                {
                    Excel.Range headerCell = null;
                    try
                    {
                        headerCell = (Excel.Range)ws.Cells[1, startColumn + col];
                        string headerValue = TrimOrEmpty(headerCell?.Value2);
                        data.Headers.Add(headerValue);

                        if (!string.Equals(headerValue, expectedHeaders[col], StringComparison.OrdinalIgnoreCase))
                        {
                            char columnLetter = (char)('A' + startColumn + col - 1);
                            throw new InvalidOperationException(
                                $"Invalid workbook: expected header '{expectedHeaders[col]}' in column {columnLetter}, found '{headerValue}'.");
                        }
                    }
                    finally
                    {
                        ExcelHelpers.ReleaseCom(headerCell);
                    }
                }

                usedRange = ws.UsedRange;
                int lastRow = 1;
                if (usedRange != null)
                {
                    try
                    {
                        lastRow = Math.Max(lastRow, usedRange.Row + usedRange.Rows.Count - 1);
                    }
                    catch
                    {
                        lastRow = 1;
                    }
                }

                for (int row = 2; row <= lastRow; row++)
                {
                    object[] rowValues = new object[columnCount];
                    bool hasData = false;

                    for (int col = 0; col < columnCount; col++)
                    {
                        Excel.Range cell = null;
                        try
                        {
                            cell = (Excel.Range)ws.Cells[row, startColumn + col];
                            object value = cell?.Value2;
                            rowValues[col] = value;
                            if (!IsNullOrEmptyExcel(value))
                            {
                                hasData = true;
                            }
                        }
                        finally
                        {
                            ExcelHelpers.ReleaseCom(cell);
                        }
                    }

                    if (!hasData)
                    {
                        continue;
                    }

                    data.FrameName.Add(TrimOrEmpty(rowValues[0]));
                    data.LoadPattern.Add(TrimOrEmpty(rowValues[1]));
                    data.MyType.Add(ParseLoadType(rowValues[2]));
                    data.CoordinateSystem.Add(TrimOrEmpty(rowValues[3]));
                    data.Direction.Add(ParseNullableInt(rowValues[4]));
                    data.RelDist1.Add(ParseNullableDouble(rowValues[5]));
                    data.RelDist2.Add(ParseNullableDouble(rowValues[6]));
                    data.Dist1.Add(ParseNullableDouble(rowValues[7]));
                    data.Dist2.Add(ParseNullableDouble(rowValues[8]));
                    data.Value1.Add(ParseNullableDouble(rowValues[9]));
                    data.Value2.Add(ParseNullableDouble(rowValues[10]));
                }

                return data;
            }
            finally
            {
                ExcelHelpers.ReleaseCom(usedRange);

                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                }

                ExcelHelpers.ReleaseCom(ws);
                ExcelHelpers.ReleaseCom(wb);
                ExcelHelpers.ReleaseCom(books);

                if (app != null)
                {
                    try { app.Quit(); } catch { }
                }

                ExcelHelpers.ReleaseCom(app);
            }
        }

        private static Excel.Worksheet FindWorksheet(Excel.Workbook wb, string sheetName)
        {
            if (wb == null) return null;
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

            Excel.Worksheet result = null;

            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                Excel.Worksheet candidate = null;
                try
                {
                    candidate = (Excel.Worksheet)wb.Worksheets[i];
                    if (candidate != null && string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        result = candidate;
                        candidate = null;
                        break;
                    }
                }
                finally
                {
                    ExcelHelpers.ReleaseCom(candidate);
                }
            }

            return result;
        }

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(ExcelLoadData data)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();

            AppendBranch(tree, 0, data.FrameName);
            AppendBranch(tree, 1, data.LoadPattern);
            AppendBranch(tree, 2, data.MyType);
            AppendBranch(tree, 3, data.CoordinateSystem);
            AppendBranch(tree, 4, data.Direction);
            AppendBranch(tree, 5, data.RelDist1);
            AppendBranch(tree, 6, data.RelDist2);
            AppendBranch(tree, 7, data.Dist1);
            AppendBranch(tree, 8, data.Dist2);
            AppendBranch(tree, 9, data.Value1);
            AppendBranch(tree, 10, data.Value2);

            return tree;
        }

        private static void AppendBranch<T>(GH_Structure<GH_ObjectWrapper> tree, int index, IList<T> values)
        {
            GH_Path path = new GH_Path(index);
            tree.EnsurePath(path);

            if (values == null)
            {
                return;
            }

            for (int i = 0; i < values.Count; i++)
            {
                tree.Append(new GH_ObjectWrapper(values[i]), path);
            }
        }

        private static string TrimOrEmpty(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();
        }

        private static bool IsNullOrEmptyExcel(object value)
        {
            if (value == null)
            {
                return true;
            }

            if (value is string s)
            {
                return string.IsNullOrWhiteSpace(s);
            }

            return false;
        }

        private static int? ParseLoadType(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is double d)
            {
                return NormalizeLoadType((int)Math.Round(d, MidpointRounding.AwayFromZero));
            }

            string s = TrimOrEmpty(value);
            if (string.IsNullOrEmpty(s))
            {
                return null;
            }

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numeric))
            {
                return NormalizeLoadType(numeric);
            }

            if (string.Equals(s, "Uniform", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (string.Equals(s, "Trapezoidal", StringComparison.OrdinalIgnoreCase))
            {
                return 2;
            }

            return null;
        }

        private static int? ParseNullableInt(object value)
        {
            if (value == null)
            {
                return 0;
            }

            if (value is double d)
            {
                return (int)Math.Round(d, MidpointRounding.AwayFromZero);
            }

            string s = TrimOrEmpty(value);
            if (string.IsNullOrEmpty(s))
            {
                return null;
            }

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int result))
            {
                return result;
            }

            return null;
        }

        private static double? ParseNullableDouble(object value)
        {
            if (value == null)
            {
                return 0;
            }

            if (value is double d)
            {
                return d;
            }

            string s = TrimOrEmpty(value);
            if (string.IsNullOrEmpty(s))
            {
                return null;
            }

            if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double res))
            {
                return res;
            }

            return null;
        }

        private static int NormalizeLoadType(int loadType)
        {
            return loadType == 2 ? 2 : 1;
        }

        private static double Clamp01(double value)
        {
            if (value < 0.0) return 0.0;
            if (value > 1.0) return 1.0;
            return value;
        }

        private static int ClampDirCode(int dirCode)
        {
            if (dirCode < 1 || dirCode > 11)
            {
                return 10;
            }

            return dirCode;
        }

        private static bool IsInvalidNumber(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value);
        }

        private static string Plural(int count, string word)
        {
            return count == 1 ? $"{count} {word}" : $"{count} {word}s";
        }

        private static void EnsureModelUnlocked(cSapModel model)
        {
            if (model == null)
            {
                return;
            }

            try
            {
                bool isLocked = model.GetModelIsLocked();
                if (isLocked)
                {
                    model.SetModelIsLocked(false);
                }
            }
            catch
            {
                // ignored
            }
        }

        private static HashSet<string> TryGetExistingFrameNames(cSapModel model)
        {
            if (model == null)
            {
                return null;
            }

            try
            {
                int count = 0;
                string[] names = null;
                int ret = model.FrameObj.GetNameList(ref count, ref names);
                if (ret != 0)
                {
                    return null;
                }

                HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (names != null)
                {
                    for (int i = 0; i < names.Length; i++)
                    {
                        string nm = names[i];
                        if (!string.IsNullOrWhiteSpace(nm))
                        {
                            result.Add(nm.Trim());
                        }
                    }
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        private class ExcelLoadData
        {
            public List<string> Headers { get; } = new List<string>();
            public List<string> FrameName { get; } = new List<string>();
            public List<string> LoadPattern { get; } = new List<string>();
            public List<int?> MyType { get; } = new List<int?>();
            public List<string> CoordinateSystem { get; } = new List<string>();
            public List<int?> Direction { get; } = new List<int?>();
            public List<double?> RelDist1 { get; } = new List<double?>();
            public List<double?> RelDist2 { get; } = new List<double?>();
            public List<double?> Dist1 { get; } = new List<double?>();
            public List<double?> Dist2 { get; } = new List<double?>();
            public List<double?> Value1 { get; } = new List<double?>();
            public List<double?> Value2 { get; } = new List<double?>();

            public int RowCount => FrameName.Count;
        }
    }
}
