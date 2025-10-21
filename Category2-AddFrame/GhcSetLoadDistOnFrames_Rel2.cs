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
using static GhcETABSAPI.ComponentShared;

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

                UiHelpers.ShowDualProgressBar(
                    "Assign Frame Distributed Loads",
                    "Reading Excel...",
                    0,
                    string.Empty,
                    0);

                FrameExcelData excelData = FrameExcelData.Read(
                    fullPath,
                    sheetName,
                    (current, maximum, status) => UiHelpers.UpdateExcelProgressBar(current, maximum, status));

                if (excelData.RowCount == 0)
                {
                    valueTree = BuildValueTree(excelData);
                    messages.Add($"Read 0 data rows from sheet '{sheetName}'. Nothing to assign.");
                    UiHelpers.CloseProgressBar();
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
                    List<string> normalizedPairs = new List<string>();

                    List<PreparedLoadAssignment> preparedLoads = PrepareLoadAssignments(
                        sapModel,
                        excelData,
                        existingNames,
                        skippedPairs,
                        failedPairs,
                        normalizedPairs,
                        ref failedCount);

                    valueTree = BuildValueTree(excelData);

                    int totalPrepared = preparedLoads.Count;
                    UiHelpers.UpdateAssignmentProgressBar(
                        0,
                        totalPrepared,
                        BuildProgressStatus(0, totalPrepared));

                    try
                    {
                        for (int j = 0; j < preparedLoads.Count; j++)
                        {
                            PreparedLoadAssignment prepared = preparedLoads[j];

                            int ret = sapModel.FrameObj.SetLoadDistributed(
                                prepared.FrameName,
                                prepared.LoadPattern,
                                prepared.LoadType,
                                prepared.Direction,
                                prepared.RelDist1,
                                prepared.RelDist2,
                                prepared.Value1,
                                prepared.Value2,
                                prepared.CoordinateSystem,
                                true,
                                replaceMode,
                                (int)eItemType.Objects);

                            if (ret == 0)
                            {
                                assignedCount++;
                                UiHelpers.UpdateAssignmentProgressBar(
                                    assignedCount,
                                    totalPrepared,
                                    BuildProgressStatus(assignedCount, totalPrepared));
                            }
                            else
                            {
                                failedCount++;
                                failedPairs.Add($"{prepared.RowIndex}:{prepared.FrameName}");
                                UiHelpers.UpdateAssignmentProgressBar(
                                    assignedCount,
                                    totalPrepared,
                                    BuildProgressStatus(assignedCount, totalPrepared));
                            }
                        }

                        UiHelpers.UpdateAssignmentProgressBar(
                            assignedCount,
                            totalPrepared,
                            BuildProgressStatus(assignedCount, totalPrepared));
                    }
                    finally
                    {
                        UiHelpers.CloseProgressBar();
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

                    if (normalizedPairs.Count > 0)
                    {
                        messages.Add("Normalized distance inputs (rowIndex:frame): " + string.Join(", ", normalizedPairs));
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

            UiHelpers.CloseProgressBar();

            da.SetDataTree(0, valueTree);
            da.SetDataList(1, messages.ToArray());

            _lastValues = valueTree.Duplicate();
            _lastMessages.Clear();
            _lastMessages.AddRange(messages);
            _lastRun = run;
        }

        private static List<PreparedLoadAssignment> PrepareLoadAssignments(
            cSapModel sapModel,
            FrameExcelData excelData,
            HashSet<string> existingNames,
            List<string> skippedPairs,
            List<string> failedPairs,
            List<string> normalizedPairs,
            ref int failedCount)
        {
            List<PreparedLoadAssignment> prepared = new List<PreparedLoadAssignment>();
            if (excelData == null)
            {
                return prepared;
            }

            Dictionary<string, double?> lengthCache = new Dictionary<string, double?>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < excelData.RowCount; i++)
            {
                string frameName = TrimOrEmpty(excelData.FrameName[i]);
                string loadPattern = TrimOrEmpty(excelData.LoadPattern[i]);
                int? rawType = excelData.MyType[i];
                int? rawDirection = excelData.Direction[i];
                double? rawRelDist1 = excelData.RelDist1[i];
                double? rawRelDist2 = excelData.RelDist2[i];
                double? rawDist1 = excelData.Dist1[i];
                double? rawDist2 = excelData.Dist2[i];
                double? rawVal1 = excelData.Value1[i];
                double? rawVal2 = excelData.Value2[i];

                bool hasRelDistances = rawRelDist1.HasValue && rawRelDist2.HasValue &&
                    !IsInvalidNumber(rawRelDist1.Value) && !IsInvalidNumber(rawRelDist2.Value);
                bool hasAbsoluteDistances = rawDist1.HasValue && rawDist2.HasValue &&
                    !IsInvalidNumber(rawDist1.Value) && !IsInvalidNumber(rawDist2.Value);

                if (string.IsNullOrEmpty(frameName) || string.IsNullOrEmpty(loadPattern) ||
                    !rawType.HasValue || !rawDirection.HasValue ||
                    !rawVal1.HasValue || !rawVal2.HasValue ||
                    (!hasRelDistances && !hasAbsoluteDistances))
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
                double val1 = rawVal1.Value;
                double val2 = rawVal2.Value;

                if (IsInvalidNumber(val1) || IsInvalidNumber(val2))
                {
                    skippedPairs.Add($"{i}:{frameName}");
                    continue;
                }

                double? frameLength = GetCachedFrameLength(sapModel, frameName, lengthCache);

                if (!TryResolveDistances(
                        frameLength,
                        rawRelDist1,
                        rawRelDist2,
                        rawDist1,
                        rawDist2,
                        out double relDist1,
                        out double relDist2,
                        out double absDist1,
                        out double absDist2,
                        out bool adjusted))
                {
                    skippedPairs.Add($"{i}:{frameName}");
                    continue;
                }

                if (adjusted)
                {
                    normalizedPairs?.Add($"{i}:{frameName}");
                }

                excelData.RelDist1[i] = relDist1;
                excelData.RelDist2[i] = relDist2;
                excelData.Dist1[i] = absDist1;
                excelData.Dist2[i] = absDist2;

                string coordinateSystem = ResolveDirectionReference(direction);
                excelData.CoordinateSystem[i] = coordinateSystem;

                prepared.Add(new PreparedLoadAssignment(
                    i,
                    frameName,
                    loadPattern,
                    loadType,
                    direction,
                    relDist1,
                    relDist2,
                    val1,
                    val2,
                    coordinateSystem));
            }

            return prepared;
        }

        private static string BuildProgressStatus(int assignedCount, int totalPrepared)
        {
            if (totalPrepared <= 0)
            {
                return "";
            }

            double percent = totalPrepared == 0 ? 0.0 : (assignedCount / (double)totalPrepared) * 100.0;
            return $"Assigned {assignedCount} of {totalPrepared} members ({percent:0.##}%).";
        }

        private readonly struct PreparedLoadAssignment
        {
            internal PreparedLoadAssignment(
                int rowIndex,
                string frameName,
                string loadPattern,
                int loadType,
                int direction,
                double relDist1,
                double relDist2,
                double value1,
                double value2,
                string coordinateSystem)
            {
                RowIndex = rowIndex;
                FrameName = frameName;
                LoadPattern = loadPattern;
                LoadType = loadType;
                Direction = direction;
                RelDist1 = relDist1;
                RelDist2 = relDist2;
                Value1 = value1;
                Value2 = value2;
                CoordinateSystem = coordinateSystem;
            }

            internal int RowIndex { get; }
            internal string FrameName { get; }
            internal string LoadPattern { get; }
            internal int LoadType { get; }
            internal int Direction { get; }
            internal double RelDist1 { get; }
            internal double RelDist2 { get; }
            internal double Value1 { get; }
            internal double Value2 { get; }
            internal string CoordinateSystem { get; }
        }

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(FrameExcelData data)
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

            string s = ExcelHelpers.TrimOrEmpty(value);
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

            return ExcelHelpers.ParseNullableInt(value);
        }

        private static double? ParseNullableDouble(object value)
        {
            if (value == null)
            {
                return 0.0;
            }

            if (value is double d)
            {
                return d;
            }

            return ExcelHelpers.ParseNullableDouble(value);
        }

        private const double DistanceTolerance = 1e-6;
        private const double LengthTolerance = 1e-9;

        private static int NormalizeLoadType(int loadType)
        {
            return loadType == 2 ? 2 : 1;
        }

        private static bool TryResolveDistances(
            double? frameLength,
            double? relDist1In,
            double? relDist2In,
            double? dist1In,
            double? dist2In,
            out double relDist1,
            out double relDist2,
            out double dist1,
            out double dist2,
            out bool adjusted)
        {
            relDist1 = 0.0;
            relDist2 = 0.0;
            dist1 = 0.0;
            dist2 = 0.0;
            adjusted = false;

            bool hasRel = relDist1In.HasValue && relDist2In.HasValue &&
                !IsInvalidNumber(relDist1In.Value) && !IsInvalidNumber(relDist2In.Value);
            bool hasAbs = dist1In.HasValue && dist2In.HasValue &&
                !IsInvalidNumber(dist1In.Value) && !IsInvalidNumber(dist2In.Value);

            if (!hasRel && !hasAbs)
            {
                return false;
            }

            double? safeLength = (frameLength.HasValue && !IsInvalidNumber(frameLength.Value) && frameLength.Value > LengthTolerance)
                ? frameLength
                : (double?)null;

            if (hasRel)
            {
                double r1 = Clamp01(relDist1In.Value);
                double r2 = Clamp01(relDist2In.Value);

                if (!NearlyEqual(r1, relDist1In.Value) || !NearlyEqual(r2, relDist2In.Value))
                {
                    adjusted = true;
                }

                if (r1 > r2)
                {
                    double tmp = r1;
                    r1 = r2;
                    r2 = tmp;
                    adjusted = true;
                }

                relDist1 = r1;
                relDist2 = r2;

                if (safeLength.HasValue)
                {
                    double length = safeLength.Value;
                    double computedAbs1 = ClampAbsolute(r1 * length, length, out bool clamped1);
                    double computedAbs2 = ClampAbsolute(r2 * length, length, out bool clamped2);

                    if (clamped1 || clamped2)
                    {
                        adjusted = true;
                    }

                    if (hasAbs)
                    {
                        double adjAbs1 = ClampAbsolute(dist1In.Value, length, out bool clampedIn1);
                        double adjAbs2 = ClampAbsolute(dist2In.Value, length, out bool clampedIn2);

                        if (clampedIn1 || clampedIn2)
                        {
                            adjusted = true;
                        }

                        if (Math.Abs(adjAbs1 - computedAbs1) > DistanceTolerance * Math.Max(1.0, length))
                        {
                            adjusted = true;
                            dist1 = computedAbs1;
                        }
                        else
                        {
                            dist1 = adjAbs1;
                        }

                        if (Math.Abs(adjAbs2 - computedAbs2) > DistanceTolerance * Math.Max(1.0, length))
                        {
                            adjusted = true;
                            dist2 = computedAbs2;
                        }
                        else
                        {
                            dist2 = adjAbs2;
                        }
                    }
                    else
                    {
                        dist1 = computedAbs1;
                        dist2 = computedAbs2;
                    }
                }
                else
                {
                    dist1 = hasAbs ? dist1In.Value : 0.0;
                    dist2 = hasAbs ? dist2In.Value : 0.0;
                }

                return true;
            }

            if (!safeLength.HasValue)
            {
                return false;
            }

            double len = safeLength.Value;
            double abs1 = ClampAbsolute(dist1In.Value, len, out bool clampedAbs1);
            double abs2 = ClampAbsolute(dist2In.Value, len, out bool clampedAbs2);

            if (clampedAbs1 || clampedAbs2)
            {
                adjusted = true;
            }

            if (abs1 > abs2)
            {
                double tmp = abs1;
                abs1 = abs2;
                abs2 = tmp;
                adjusted = true;
            }

            double rawRel1 = len <= 0.0 ? 0.0 : abs1 / len;
            double rawRel2 = len <= 0.0 ? 0.0 : abs2 / len;
            double r1Out = Clamp01(rawRel1);
            double r2Out = Clamp01(rawRel2);

            if (!NearlyEqual(r1Out, rawRel1) || !NearlyEqual(r2Out, rawRel2))
            {
                adjusted = true;
            }

            relDist1 = r1Out;
            relDist2 = r2Out;
            dist1 = relDist1 * len;
            dist2 = relDist2 * len;
            return true;
        }

        private static double ClampAbsolute(double value, double length, out bool clamped)
        {
            double original = value;
            double max = Math.Max(0.0, length);

            if (value < 0.0)
            {
                value = 0.0;
            }
            if (value > max)
            {
                value = max;
            }

            clamped = Math.Abs(value - original) > DistanceTolerance * Math.Max(1.0, max);
            return value;
        }

        private static bool NearlyEqual(double a, double b)
        {
            double scale = Math.Max(1.0, Math.Abs(a) + Math.Abs(b));
            return Math.Abs(a - b) <= DistanceTolerance * scale;
        }

        private static double? GetCachedFrameLength(cSapModel model, string frameName, IDictionary<string, double?> cache)
        {
            if (cache == null || string.IsNullOrWhiteSpace(frameName))
            {
                return null;
            }

            if (cache.TryGetValue(frameName, out double? cached))
            {
                return cached;
            }

            double? length = TryGetFrameLength(model, frameName);
            cache[frameName] = length;
            return length;
        }

        private static double? TryGetFrameLength(cSapModel model, string frameName)
        {
            if (model == null || string.IsNullOrWhiteSpace(frameName))
            {
                return null;
            }

            try
            {
                string pointI = null;
                string pointJ = null;
                int ret = model.FrameObj.GetPoints(frameName, ref pointI, ref pointJ);
                if (ret != 0 || string.IsNullOrWhiteSpace(pointI) || string.IsNullOrWhiteSpace(pointJ))
                {
                    return null;
                }

                double xi = 0.0, yi = 0.0, zi = 0.0;
                double xj = 0.0, yj = 0.0, zj = 0.0;

                ret = model.PointObj.GetCoordCartesian(pointI, ref xi, ref yi, ref zi);
                if (ret != 0)
                {
                    return null;
                }

                ret = model.PointObj.GetCoordCartesian(pointJ, ref xj, ref yj, ref zj);
                if (ret != 0)
                {
                    return null;
                }

                double dx = xj - xi;
                double dy = yj - yi;
                double dz = zj - zi;
                double length = Math.Sqrt((dx * dx) + (dy * dy) + (dz * dz));

                if (IsInvalidNumber(length) || length <= LengthTolerance)
                {
                    return null;
                }

                return length;
            }
            catch
            {
                return null;
            }
        }

        private sealed class FrameExcelData
        {
            private FrameExcelData()
            {
            }

            internal static FrameExcelData Read(
                string fullPath,
                string sheetName,
                Action<int, int, string> progressCallback)
            {
                const string expectedSheetName = "Assigned Loads On Frames";
                if (!string.Equals(sheetName, expectedSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException($"Invalid workbook: expected sheet name '{expectedSheetName}'.");
                }

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

                ExcelSheetColumnarData table = ExcelHelpers.ReadColumnTable(
                    fullPath,
                    sheetName,
                    2,
                    expectedHeaders,
                    progressCallback);

                FrameExcelData data = new FrameExcelData();
                data.Headers.AddRange(table.Headers);

                int rowCount = table.RowCount;
                List<object> frameNames = table.GetColumn(0);
                List<object> loadPatterns = table.GetColumn(1);
                List<object> loadTypes = table.GetColumn(2);
                List<object> coordinateSystems = table.GetColumn(3);
                List<object> directions = table.GetColumn(4);
                List<object> relDist1 = table.GetColumn(5);
                List<object> relDist2 = table.GetColumn(6);
                List<object> dist1 = table.GetColumn(7);
                List<object> dist2 = table.GetColumn(8);
                List<object> value1 = table.GetColumn(9);
                List<object> value2 = table.GetColumn(10);

                for (int i = 0; i < rowCount; i++)
                {
                    data.FrameName.Add(ExcelHelpers.TrimOrEmpty(frameNames[i]));
                    data.LoadPattern.Add(ExcelHelpers.TrimOrEmpty(loadPatterns[i]));
                    data.MyType.Add(ParseLoadType(loadTypes[i]));
                    data.CoordinateSystem.Add(ExcelHelpers.TrimOrEmpty(coordinateSystems[i]));
                    data.Direction.Add(ParseNullableInt(directions[i]));
                    data.RelDist1.Add(ParseNullableDouble(relDist1[i]));
                    data.RelDist2.Add(ParseNullableDouble(relDist2[i]));
                    data.Dist1.Add(ParseNullableDouble(dist1[i]));
                    data.Dist2.Add(ParseNullableDouble(dist2[i]));
                    data.Value1.Add(ParseNullableDouble(value1[i]));
                    data.Value2.Add(ParseNullableDouble(value2[i]));
                }

                return data;
            }

            internal List<string> Headers { get; } = new List<string>();
            internal List<string> FrameName { get; } = new List<string>();
            internal List<string> LoadPattern { get; } = new List<string>();
            internal List<int?> MyType { get; } = new List<int?>();
            internal List<string> CoordinateSystem { get; } = new List<string>();
            internal List<int?> Direction { get; } = new List<int?>();
            internal List<double?> RelDist1 { get; } = new List<double?>();
            internal List<double?> RelDist2 { get; } = new List<double?>();
            internal List<double?> Dist1 { get; } = new List<double?>();
            internal List<double?> Dist2 { get; } = new List<double?>();
            internal List<double?> Value1 { get; } = new List<double?>();
            internal List<double?> Value2 { get; } = new List<double?>();

            internal int RowCount => FrameName.Count;
        }
    }
}
