// -------------------------------------------------------------
// Component : Sync Point Coordinates From Excel
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 8.0 (x64)
// Depends   : Grasshopper, ETABSv1 (COM), Microsoft.Office.Interop.Excel
// Panel     : "MGT" / "2.0 Point Object Modelling"
// -------------------------------------------------------------
// Inputs (ordered):
//   0) run        (bool, item)    Rising-edge trigger; executes on False→True transition.
//   1) sapModel   (ETABSv1.cSapModel, item)  ETABS model from the Attach component.
//   2) excelPath  (string, item)  Path to the Excel workbook (relative paths resolved against the plug-in folder).
//   3) sheetName  (string, item)  Name of the worksheet containing point data. Defaults to "PointObjects".
//   4) startRow   (int, item)     First data row (headers expected on startRow-1). Defaults to 2.
//   5) scale      (double, item)  Coordinate multiplier (e.g., 1000 for mm→m). Defaults to 1.0.
//   6) tolerance  (double, item)  Coordinate change tolerance to ignore tiny differences. Defaults to 1e-6.
//
// Outputs:
//   0) headers    (text, tree)    Column headers describing the report tree.
//   1) values     (generic, tree) Update report aligned with the header order.
//   2) msg        (text, item)    Summary / diagnostics message.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ETABSv1;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Excel = Microsoft.Office.Interop.Excel;
using static MGT.ComponentShared;

namespace MGT
{
    public class GhcSyncPointCoordinatesFromExcel : GH_Component
    {
        private bool _lastRun;
        private GH_Structure<GH_String> _lastHeaders = PointCoordinateSyncWorkflow.CreateHeaderTree();
        private GH_Structure<GH_ObjectWrapper> _lastValues = new GH_Structure<GH_ObjectWrapper>();
        private string _lastMessage = "Idle.";

        public GhcSyncPointCoordinatesFromExcel()
          : base(
                "Sync Point Coordinates From Excel",
                "SyncPointsExcel",
                "Compare ETABS point coordinates against an Excel worksheet and push changes back to ETABS when detected.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                "MGT",
                "2.0 Point Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("977c2162-2fe7-46b6-90ff-b87cbb0b7af9");

        protected override System.Drawing.Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("run", "run", "Press to run once (rising edge).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            p.AddTextParameter("excelPath", "excelPath", "Excel workbook path (relative paths resolved against the plug-in folder).", GH_ParamAccess.item, string.Empty);
            p.AddTextParameter("sheetName", "sheetName", "Worksheet containing point data (defaults to \"PointObjects\").", GH_ParamAccess.item, "PointObjects");
            p.AddIntegerParameter("startRow", "startRow", "First data row (headers expected on startRow-1).", GH_ParamAccess.item, 2);
            p.AddNumberParameter("scale", "scale", "Coordinate multiplier (e.g., 1000 for mm→m).", GH_ParamAccess.item, 1.0);
            p.AddNumberParameter("tolerance", "tolerance", "Coordinate tolerance (same units as ETABS model).", GH_ParamAccess.item, 1e-6);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("headers", "headers", "Header labels describing each value column.", GH_ParamAccess.tree);
            p.AddGenericParameter("values", "values", "Point update report aligned with the headers.", GH_ParamAccess.tree);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sapModel = null;
            string excelPath = null;
            string sheetName = "PointObjects";
            int startRow = 2;
            double scale = 1.0;
            double tolerance = 1e-6;

            da.GetData(0, ref run);
            da.GetData(1, ref sapModel);
            da.GetData(2, ref excelPath);
            da.GetData(3, ref sheetName);
            da.GetData(4, ref startRow);
            da.GetData(5, ref scale);
            da.GetData(6, ref tolerance);

            bool rising = !_lastRun && run;
            if (!rising)
            {
                da.SetDataTree(0, _lastHeaders.Duplicate());
                da.SetDataTree(1, _lastValues.Duplicate());
                da.SetData(2, _lastMessage);
                _lastRun = run;
                return;
            }

            sheetName = string.IsNullOrWhiteSpace(sheetName) ? "PointObjects" : sheetName;
            if (startRow < 2)
            {
                startRow = 2;
            }

            if (IsInvalidNumber(scale) || scale <= 0.0)
            {
                scale = 1.0;
            }

            if (IsInvalidNumber(tolerance) || tolerance < 0.0)
            {
                tolerance = 1e-6;
            }

            // Wrap the raw Grasshopper inputs in a tiny request object so the workflow
            // stays easy to follow for newcomers.
            PointCoordinateSyncWorkflow.Request request = new PointCoordinateSyncWorkflow.Request(
                sapModel,
                excelPath,
                sheetName,
                startRow,
                scale,
                tolerance);

            GH_Structure<GH_String> headers = PointCoordinateSyncWorkflow.CreateHeaderTree();
            GH_Structure<GH_ObjectWrapper> values = new GH_Structure<GH_ObjectWrapper>();
            string message = string.Empty;

            try
            {
                PointCoordinateSyncWorkflow.Result result = PointCoordinateSyncWorkflow.Run(request);

                values = result.HasRows
                    ? PointCoordinateSyncWorkflow.BuildValueTree(result.Rows)
                    : new GH_Structure<GH_ObjectWrapper>();
                message = PointCoordinateSyncWorkflow.BuildSummary(result);

                foreach (string warning in result.Warnings)
                {
                    if (!string.IsNullOrWhiteSpace(warning))
                    {
                        AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warning);
                    }
                }
            }
            catch (Exception ex)
            {
                message = "Failed: " + ex.Message;
                values = new GH_Structure<GH_ObjectWrapper>();
            }

            da.SetDataTree(0, headers);
            da.SetDataTree(1, values);
            da.SetData(2, message);

            _lastHeaders = headers;
            _lastValues = values;
            _lastMessage = message;
            _lastRun = run;
        }
    }

    internal static class PointCoordinateSyncWorkflow
    {
        private static readonly string[] HeaderLabels =
        {
            "UniqueName",
            "ExcelX",
            "ExcelY",
            "ExcelZ",
            "ModelXBefore",
            "ModelYBefore",
            "ModelZBefore",
            "ModelXAfter",
            "ModelYAfter",
            "ModelZAfter",
            "Changed",
            "Status"
        };

        internal static GH_Structure<GH_String> CreateHeaderTree()
        {
            GH_Structure<GH_String> tree = new GH_Structure<GH_String>();
            GH_Path path = new GH_Path(0);

            foreach (string label in HeaderLabels)
            {
                tree.Append(new GH_String(label), path);
            }

            return tree;
        }

        internal static GH_Structure<GH_ObjectWrapper> BuildValueTree(IReadOnlyList<RowResult> rows)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();
            if (rows == null || rows.Count == 0)
            {
                return tree;
            }

            for (int column = 0; column < HeaderLabels.Length; column++)
            {
                GH_Path path = new GH_Path(column);
                tree.EnsurePath(path);

                for (int row = 0; row < rows.Count; row++)
                {
                    tree.Append(new GH_ObjectWrapper(ReadValue(rows[row], column)), path);
                }
            }

            return tree;
        }

        internal static string BuildSummary(Result summary)
        {
            if (summary == null)
            {
                return "No updates performed.";
            }

            if (!summary.HasRows)
            {
                return "No Excel rows were read.";
            }

            string message = string.Format(
                "Processed {0} row(s): {1} updated, {2} unchanged, {3} skipped, {4} failed.",
                summary.ExcelRowCount,
                summary.Updated,
                summary.Unchanged,
                summary.Skipped,
                summary.Failed);

            string warningText = string.Join(" | ", summary.Warnings.Where(w => !string.IsNullOrWhiteSpace(w)));
            if (!string.IsNullOrWhiteSpace(warningText))
            {
                message += " Warnings: " + warningText;
            }

            return message;
        }

        internal static Result Run(Request request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.SapModel == null)
            {
                throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");
            }

            string resolvedPath = ExcelHelpers.ProjectRelative(request.ExcelPath);
            if (string.IsNullOrWhiteSpace(resolvedPath))
            {
                throw new InvalidOperationException("excelPath is empty.");
            }

            if (!File.Exists(resolvedPath))
            {
                throw new FileNotFoundException("Excel workbook not found.", resolvedPath);
            }

            string sheetName = string.IsNullOrWhiteSpace(request.SheetName) ? "PointObjects" : request.SheetName;
            List<ExcelRow> excelRows = ExcelReader.ReadRows(resolvedPath, sheetName, request.StartRow);
            if (excelRows.Count == 0)
            {
                return new Result(0);
            }

            EnsureModelUnlocked(request.SapModel);
            return Updater.Apply(request.SapModel, excelRows, request.Scale, request.Tolerance);
        }

        private static object ReadValue(RowResult row, int columnIndex)
        {
            switch (columnIndex)
            {
                case 0:
                    return row.UniqueName;
                case 1:
                    return row.ExcelX ?? double.NaN;
                case 2:
                    return row.ExcelY ?? double.NaN;
                case 3:
                    return row.ExcelZ ?? double.NaN;
                case 4:
                    return Sanitize(row.ModelXBefore);
                case 5:
                    return Sanitize(row.ModelYBefore);
                case 6:
                    return Sanitize(row.ModelZBefore);
                case 7:
                    return Sanitize(row.ModelXAfter);
                case 8:
                    return Sanitize(row.ModelYAfter);
                case 9:
                    return Sanitize(row.ModelZAfter);
                case 10:
                    return row.Changed;
                case 11:
                    return row.Status ?? string.Empty;
                default:
                    return string.Empty;
            }
        }

        private static double Sanitize(double value)
        {
            return IsInvalidNumber(value) ? double.NaN : value;
        }

        internal sealed class Request
        {
            internal Request(cSapModel sapModel, string excelPath, string sheetName, int startRow, double scale, double tolerance)
            {
                SapModel = sapModel;
                ExcelPath = excelPath;
                SheetName = sheetName;
                StartRow = startRow;
                Scale = scale;
                Tolerance = tolerance;
            }

            internal cSapModel SapModel { get; }
            internal string ExcelPath { get; }
            internal string SheetName { get; }
            internal int StartRow { get; }
            internal double Scale { get; }
            internal double Tolerance { get; }
        }

        internal sealed class Result
        {
            internal Result(int excelRowCount)
            {
                ExcelRowCount = excelRowCount;
            }

            internal int ExcelRowCount { get; }
            internal bool HasRows => ExcelRowCount > 0;
            internal List<RowResult> Rows { get; } = new List<RowResult>();
            internal List<string> Warnings { get; } = new List<string>();
            internal int Updated { get; set; }
            internal int Unchanged { get; set; }
            internal int Skipped { get; set; }
            internal int Failed { get; set; }
        }

        internal sealed class RowResult
        {
            internal string UniqueName { get; set; }
            internal double? ExcelX { get; set; }
            internal double? ExcelY { get; set; }
            internal double? ExcelZ { get; set; }
            internal double ModelXBefore { get; set; }
            internal double ModelYBefore { get; set; }
            internal double ModelZBefore { get; set; }
            internal double ModelXAfter { get; set; }
            internal double ModelYAfter { get; set; }
            internal double ModelZAfter { get; set; }
            internal bool Changed { get; set; }
            internal string Status { get; set; }
            internal int RowNumber { get; set; }
        }

        private sealed class ExcelRow
        {
            internal string UniqueName { get; set; }
            internal double? X { get; set; }
            internal double? Y { get; set; }
            internal double? Z { get; set; }
            internal int RowNumber { get; set; }
        }

        private static class ExcelReader
        {
            internal static List<ExcelRow> ReadRows(string workbookPath, string sheetName, int startRow)
            {
                List<ExcelRow> rows = new List<ExcelRow>();

                Excel.Application app = null;
                Excel.Workbooks workbooks = null;
                Excel.Workbook workbook = null;
                Excel.Worksheet worksheet = null;
                Excel.Range usedRange = null;

                try
                {
                    app = new Excel.Application
                    {
                        Visible = false,
                        DisplayAlerts = false
                    };

                    workbooks = app.Workbooks;
                    workbook = workbooks.Open(
                        Filename: workbookPath,
                        UpdateLinks: 0,
                        ReadOnly: true,
                        IgnoreReadOnlyRecommended: true,
                        AddToMru: false);

                    worksheet = ResolveWorksheet(workbook, sheetName);
                    usedRange = worksheet?.UsedRange;
                    if (usedRange == null)
                    {
                        return rows;
                    }

                    if (usedRange.Value2 is not object[,] data)
                    {
                        return rows;
                    }

                    int firstRow = usedRange.Row;
                    int lastRow = firstRow + usedRange.Rows.Count - 1;
                    int firstColumn = usedRange.Column;
                    int lastColumn = firstColumn + usedRange.Columns.Count - 1;

                    int headerRowNumber = Math.Max(startRow - 1, firstRow);
                    Dictionary<string, int> headerMap = ReadHeaderMap(data, firstRow, firstColumn, headerRowNumber, lastColumn);

                    int uniqueNameColumn = ResolveColumn(headerMap, "UniqueName", firstColumn);
                    int xColumn = ResolveColumn(headerMap, "X", firstColumn + 1);
                    int yColumn = ResolveColumn(headerMap, "Y", firstColumn + 2);
                    int zColumn = ResolveColumn(headerMap, "Z", firstColumn + 3);

                    if (uniqueNameColumn < 0)
                    {
                        throw new InvalidOperationException("Worksheet must contain a 'UniqueName' column.");
                    }

                    if (xColumn < 0 || yColumn < 0 || zColumn < 0)
                    {
                        throw new InvalidOperationException("Worksheet must contain 'X', 'Y', and 'Z' columns.");
                    }

                    int dataStartRow = Math.Max(startRow, firstRow);
                    for (int rowNumber = dataStartRow; rowNumber <= lastRow; rowNumber++)
                    {
                        string uniqueName = ReadText(data, rowNumber, uniqueNameColumn, firstRow, firstColumn);
                        double? x = ReadNumber(data, rowNumber, xColumn, firstRow, firstColumn);
                        double? y = ReadNumber(data, rowNumber, yColumn, firstRow, firstColumn);
                        double? z = ReadNumber(data, rowNumber, zColumn, firstRow, firstColumn);

                        bool hasAnyValue = !string.IsNullOrWhiteSpace(uniqueName) || x.HasValue || y.HasValue || z.HasValue;
                        if (!hasAnyValue)
                        {
                            continue;
                        }

                        rows.Add(new ExcelRow
                        {
                            UniqueName = string.IsNullOrWhiteSpace(uniqueName) ? string.Empty : uniqueName.Trim(),
                            X = x,
                            Y = y,
                            Z = z,
                            RowNumber = rowNumber
                        });
                    }

                    return rows;
                }
                finally
                {
                    if (usedRange != null)
                    {
                        ExcelHelpers.ReleaseCom(usedRange);
                    }

                    if (worksheet != null)
                    {
                        ExcelHelpers.ReleaseCom(worksheet);
                    }

                    if (workbook != null)
                    {
                        try
                        {
                            workbook.Close(false);
                        }
                        catch
                        {
                        }

                        ExcelHelpers.ReleaseCom(workbook);
                    }

                    if (workbooks != null)
                    {
                        ExcelHelpers.ReleaseCom(workbooks);
                    }

                    if (app != null)
                    {
                        try
                        {
                            app.Quit();
                        }
                        catch
                        {
                        }

                        ExcelHelpers.ReleaseCom(app);
                    }
                }
            }

            private static Excel.Worksheet ResolveWorksheet(Excel.Workbook workbook, string sheetName)
            {
                if (workbook == null)
                {
                    return null;
                }

                if (!string.IsNullOrWhiteSpace(sheetName))
                {
                    foreach (Excel.Worksheet candidate in workbook.Worksheets)
                    {
                        if (string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return candidate;
                        }

                        ExcelHelpers.ReleaseCom(candidate);
                    }
                }

                return workbook.Worksheets.Count > 0 ? (Excel.Worksheet)workbook.Worksheets[1] : null;
            }

            private static Dictionary<string, int> ReadHeaderMap(object[,] data, int firstRow, int firstColumn, int headerRowNumber, int lastColumn)
            {
                Dictionary<string, int> map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                int headerRowIndex = headerRowNumber - firstRow + 1;
                if (headerRowIndex < 1 || headerRowIndex > data.GetLength(0))
                {
                    return map;
                }

                int columnCount = lastColumn - firstColumn + 1;
                for (int offset = 0; offset < columnCount; offset++)
                {
                    object raw = data[headerRowIndex, offset + 1];
                    string label = raw == null ? string.Empty : raw.ToString();
                    if (!string.IsNullOrWhiteSpace(label) && !map.ContainsKey(label))
                    {
                        map.Add(label.Trim(), firstColumn + offset);
                    }
                }

                return map;
            }

            private static int ResolveColumn(Dictionary<string, int> headerMap, string key, int fallbackColumn)
            {
                if (headerMap != null && headerMap.TryGetValue(key, out int column))
                {
                    return column;
                }

                return fallbackColumn;
            }

            private static string ReadText(object[,] data, int rowNumber, int columnNumber, int firstRow, int firstColumn)
            {
                object raw = ReadValue(data, rowNumber, columnNumber, firstRow, firstColumn);
                return raw?.ToString() ?? string.Empty;
            }

            private static double? ReadNumber(object[,] data, int rowNumber, int columnNumber, int firstRow, int firstColumn)
            {
                object raw = ReadValue(data, rowNumber, columnNumber, firstRow, firstColumn);
                if (raw == null)
                {
                    return null;
                }

                if (raw is double direct)
                {
                    return direct;
                }

                if (double.TryParse(raw.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed))
                {
                    return parsed;
                }

                return null;
            }

            private static object ReadValue(object[,] data, int rowNumber, int columnNumber, int firstRow, int firstColumn)
            {
                int rowIndex = rowNumber - firstRow + 1;
                int columnIndex = columnNumber - firstColumn + 1;

                if (rowIndex < 1 || columnIndex < 1 || rowIndex > data.GetLength(0) || columnIndex > data.GetLength(1))
                {
                    return null;
                }

                return data[rowIndex, columnIndex];
            }
        }

        private static class Updater
        {
            internal static Result Apply(cSapModel sapModel, IReadOnlyList<ExcelRow> rows, double scale, double tolerance)
            {
                Result report = new Result(rows?.Count ?? 0);
                if (sapModel == null || rows == null || rows.Count == 0)
                {
                    return report;
                }

                HashSet<string> seenUniqueNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (ExcelRow row in rows)
                {
                    if (row == null)
                    {
                        continue;
                    }

                    RowResult rowReport = new RowResult
                    {
                        UniqueName = row.UniqueName ?? string.Empty,
                        ExcelX = row.X,
                        ExcelY = row.Y,
                        ExcelZ = row.Z,
                        RowNumber = row.RowNumber,
                        ModelXBefore = double.NaN,
                        ModelYBefore = double.NaN,
                        ModelZBefore = double.NaN,
                        ModelXAfter = double.NaN,
                        ModelYAfter = double.NaN,
                        ModelZAfter = double.NaN,
                        Status = string.Empty,
                        Changed = false
                    };

                    if (!ValidateUniqueName(row, seenUniqueNames, report, rowReport))
                    {
                        continue;
                    }

                    if (!TryReadCurrentCoordinates(sapModel, row.UniqueName, row.RowNumber, report, rowReport, out double modelX, out double modelY, out double modelZ))
                    {
                        continue;
                    }

                    rowReport.ModelXBefore = modelX;
                    rowReport.ModelYBefore = modelY;
                    rowReport.ModelZBefore = modelZ;

                    if (!HasCoordinateOverride(row))
                    {
                        FinishRow(report, rowReport, RowOutcome.Unchanged, modelX, modelY, modelZ, "No Excel override; kept ETABS value.");
                        continue;
                    }

                    double targetX = row.X.HasValue ? row.X.Value * scale : modelX;
                    double targetY = row.Y.HasValue ? row.Y.Value * scale : modelY;
                    double targetZ = row.Z.HasValue ? row.Z.Value * scale : modelZ;

                    if (IsInvalidNumber(targetX) || IsInvalidNumber(targetY) || IsInvalidNumber(targetZ))
                    {
                        report.Warnings.Add($"Row {row.RowNumber}: invalid coordinate detected; skipped.");
                        FinishRow(report, rowReport, RowOutcome.Skipped, modelX, modelY, modelZ, "Invalid coordinate input.");
                        continue;
                    }

                    bool needsUpdate = RequiresUpdate(modelX, targetX, tolerance)
                        || RequiresUpdate(modelY, targetY, tolerance)
                        || RequiresUpdate(modelZ, targetZ, tolerance);

                    if (!needsUpdate)
                    {
                        FinishRow(report, rowReport, RowOutcome.Unchanged, modelX, modelY, modelZ, "Within tolerance; no update.");
                        continue;
                    }

                    if (!TryWriteCoordinates(sapModel, row.UniqueName, row.RowNumber, targetX, targetY, targetZ, report, rowReport))
                    {
                        continue;
                    }

                    double confirmedX = targetX;
                    double confirmedY = targetY;
                    double confirmedZ = targetZ;
                    TryConfirmCoordinates(sapModel, row.UniqueName, row.RowNumber, report, ref confirmedX, ref confirmedY, ref confirmedZ);

                    rowReport.ModelXAfter = confirmedX;
                    rowReport.ModelYAfter = confirmedY;
                    rowReport.ModelZAfter = confirmedZ;

                    FinishRow(report, rowReport, RowOutcome.Updated, confirmedX, confirmedY, confirmedZ, "Updated coordinates.");
                }

                return report;
            }

            private static bool ValidateUniqueName(ExcelRow row, HashSet<string> seenUniqueNames, Result report, RowResult rowReport)
            {
                if (string.IsNullOrWhiteSpace(row.UniqueName))
                {
                    FinishRow(report, rowReport, RowOutcome.Skipped, double.NaN, double.NaN, double.NaN, "UniqueName missing.");
                    return false;
                }

                if (!seenUniqueNames.Add(row.UniqueName))
                {
                    report.Warnings.Add($"Row {row.RowNumber}: duplicate UniqueName '{row.UniqueName}' skipped.");
                    FinishRow(report, rowReport, RowOutcome.Skipped, double.NaN, double.NaN, double.NaN, "Duplicate UniqueName.");
                    return false;
                }

                return true;
            }

            private static bool TryReadCurrentCoordinates(cSapModel sapModel, string uniqueName, int rowNumber, Result report, RowResult rowReport, out double x, out double y, out double z)
            {
                x = double.NaN;
                y = double.NaN;
                z = double.NaN;

                try
                {
                    int ret = sapModel.PointObj.GetCoordCartesian(uniqueName, ref x, ref y, ref z);
                    if (ret == 0)
                    {
                        return true;
                    }

                    report.Warnings.Add($"Row {rowNumber}: GetCoordCartesian for '{uniqueName}' returned {ret}.");
                }
                catch (Exception ex)
                {
                    report.Warnings.Add($"Row {rowNumber}: GetCoordCartesian exception for '{uniqueName}' - {ex.Message}");
                }

                FinishRow(report, rowReport, RowOutcome.Failed, x, y, z, "Failed to read ETABS coordinates.");
                return false;
            }

            private static bool TryWriteCoordinates(cSapModel sapModel, string uniqueName, int rowNumber, double x, double y, double z, Result report, RowResult rowReport)
            {
                try
                {
                    int ret = sapModel.PointObj.SetCoordCartesian(uniqueName, x, y, z);
                    if (ret == 0)
                    {
                        return true;
                    }

                    report.Warnings.Add($"Row {rowNumber}: SetCoordCartesian for '{uniqueName}' returned {ret}.");
                }
                catch (Exception ex)
                {
                    report.Warnings.Add($"Row {rowNumber}: SetCoordCartesian exception for '{uniqueName}' - {ex.Message}");
                }

                FinishRow(report, rowReport, RowOutcome.Failed, rowReport.ModelXBefore, rowReport.ModelYBefore, rowReport.ModelZBefore, "Failed to update ETABS.");
                return false;
            }

            private static void TryConfirmCoordinates(cSapModel sapModel, string uniqueName, int rowNumber, Result report, ref double x, ref double y, ref double z)
            {
                try
                {
                    int ret = sapModel.PointObj.GetCoordCartesian(uniqueName, ref x, ref y, ref z);
                    if (ret != 0)
                    {
                        report.Warnings.Add($"Row {rowNumber}: post-update GetCoordCartesian returned {ret}.");
                    }
                }
                catch (Exception ex)
                {
                    report.Warnings.Add($"Row {rowNumber}: post-update GetCoordCartesian exception for '{uniqueName}' - {ex.Message}");
                }
            }

            private static bool HasCoordinateOverride(ExcelRow row)
            {
                return row.X.HasValue || row.Y.HasValue || row.Z.HasValue;
            }

            private static bool RequiresUpdate(double current, double target, double tolerance)
            {
                if (IsInvalidNumber(current) || IsInvalidNumber(target))
                {
                    return true;
                }

                if (tolerance <= 0.0)
                {
                    return !current.Equals(target);
                }

                return Math.Abs(current - target) > tolerance;
            }

            private static void FinishRow(Result report, RowResult rowReport, RowOutcome outcome, double afterX, double afterY, double afterZ, string status)
            {
                rowReport.ModelXAfter = afterX;
                rowReport.ModelYAfter = afterY;
                rowReport.ModelZAfter = afterZ;
                rowReport.Status = status;

                switch (outcome)
                {
                    case RowOutcome.Updated:
                        rowReport.Changed = true;
                        report.Updated++;
                        break;
                    case RowOutcome.Unchanged:
                        report.Unchanged++;
                        break;
                    case RowOutcome.Skipped:
                        report.Skipped++;
                        break;
                    case RowOutcome.Failed:
                        report.Failed++;
                        break;
                }

                report.Rows.Add(rowReport);
            }

            private enum RowOutcome
            {
                Updated,
                Unchanged,
                Skipped,
                Failed
            }
        }
    }
}
