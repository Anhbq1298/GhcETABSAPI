// -------------------------------------------------------------
// Component : Get Frame Distributed Loads (per-object query)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)  [Embed Interop Types = False]
// Panel     : "ETABS API" / "2.0 Frame Object Modelling"
// GUID      : a1cfe4a7-9d49-42eb-aac9-774cdd7d1e84
// -------------------------------------------------------------
// Inputs (ordered):
//   0) sapModel    (ETABSv1.cSapModel, item)  ETABS model from your Attach component.
//   1) frameNames  (string, list)  Frame object names to query. Blank/dup ignored (case-insensitive).
//   2) loadPattern (string, list)  OPTIONAL filters. If UNCONNECTED or empty → treated as null (no filter).
//
// Outputs:
//   0) header  (text, tree)   Single branch of column labels.
//   1) values  (generic, tree) Column-wise branches (0..10) aligned to header order.
//   2) msg     (text, item)   Status / diagnostics.
//
// Behavior Notes:
//   • frameNames are trimmed and de-duplicated (OrdinalIgnoreCase).
//   • loadPattern is OPTIONAL; null ⇒ return all patterns; when provided, filter is case-insensitive.
//   • Values tree is column-oriented to match header labels:
//       [0] FrameName, [1] LoadPattern, [2] Type, [3] CoordinateSystem, [4] Direction,
//       [5] RelDist1, [6] RelDist2, [7] Dist1, [8] Dist2, [9] Value1, [10] Value2.
//   • Uses per-object mode: FrameObj.GetLoadDistributed(..., eItemType.Objects).
//   • CoordinateSystem is derived from Direction (|dir| < 10 ⇒ "Local", otherwise "Global").
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using ETABSv1;

namespace GhcETABSAPI
{
    public class GhcGetLoadDistOnFrames : GH_Component
    {
        private const string IdleMessage = "Idle.";

        private static readonly string[] HeaderLabels =
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

        private bool lastRun;
        private GH_Structure<GH_String> lastHeaderTree = BuildHeaderTree();
        private GH_Structure<GH_ObjectWrapper> lastValueTree = new GH_Structure<GH_ObjectWrapper>();
        private string lastMessage = IdleMessage;

        public GhcGetLoadDistOnFrames()
          : base(
                "Get Frame Distributed Loads",
                "GetFrameDistLoads",
                "Query distributed loads assigned to ETABS frame objects (per object mode).",
                "ETABS API",
                "2.0 Frame Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("a1cfe4a7-9d49-42eb-aac9-774cdd7d1e84");

        protected override System.Drawing.Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("run", "run", "Press to query (rising edge trigger).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            int frameNameIndex = p.AddTextParameter(
                "frameNames",
                "frameNames",
                "Frame object names to query. Blank entries are ignored. If empty, returns zero results.",
                GH_ParamAccess.list);
            p[frameNameIndex].Optional = true;

            int loadPatternIndex = p.AddTextParameter(
                "loadPatterns",
                "loadPatterns",
                "Optional load pattern filters. Leave empty to return all load patterns.",
                GH_ParamAccess.list);
            p[loadPatternIndex].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("headers", "headers", "Header labels describing each value column.", GH_ParamAccess.tree);
            p.AddGenericParameter("values", "values", "Distributed load rows. Each branch matches the header order.", GH_ParamAccess.tree);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sapModel = null;
            List<string> frameNames = new List<string>();
            List<string> loadPatternFilters = new List<string>();

            da.GetData(0, ref run);
            da.GetData(1, ref sapModel);
            da.GetDataList(2, frameNames);
            da.GetDataList(3, loadPatternFilters);

            bool rising = !lastRun && run;

            if (!rising)
            {
                da.SetDataTree(0, lastHeaderTree.Duplicate());
                da.SetDataTree(1, lastValueTree.Duplicate());
                da.SetData(2, lastMessage);
                lastRun = run;
                return;
            }

            if (sapModel == null)
            {
                string warning = "sapModel is null. Wire it from the Attach component.";
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warning);
                UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), warning, run);
                return;
            }

            try
            {
                List<string> trimmed = new List<string>();
                HashSet<string> seen = new HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
                if (frameNames != null)
                {
                    for (int i = 0; i < frameNames.Count; i++)
                    {
                        string nm = frameNames[i];
                        if (string.IsNullOrWhiteSpace(nm))
                        {
                            continue;
                        }

                        string clean = nm.Trim();
                        if (seen.Add(clean))
                        {
                            trimmed.Add(clean);
                        }
                    }
                }

                var rawResult = GetFrameDistributed(sapModel, trimmed);

                List<string> trimmedFilters = new List<string>();
                if (loadPatternFilters != null)
                {
                    HashSet<string> filterSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    for (int i = 0; i < loadPatternFilters.Count; i++)
                    {
                        string lp = loadPatternFilters[i];
                        if (string.IsNullOrWhiteSpace(lp))
                        {
                            continue;
                        }

                        string cleanFilter = lp.Trim();
                        if (filterSet.Add(cleanFilter))
                        {
                            trimmedFilters.Add(cleanFilter);
                        }
                    }
                }

                bool hasFilters = trimmedFilters.Count > 0;
                var result = FilterByLoadPattern(rawResult, trimmedFilters);

                GH_Structure<GH_String> headerTree = BuildHeaderTree();
                GH_Structure<GH_ObjectWrapper> valueTree = BuildValueTree(result);

                string status;
                if (trimmed.Count == 0)
                {
                    status = "No valid frame names provided.";
                }
                else if (hasFilters)
                {
                    string patternSummary = FormatLoadPatternSummary(trimmedFilters);

                    if (result.total == 0)
                    {
                        if (rawResult.total > 0)
                        {
                            status = $"No distributed loads matched {patternSummary}.";
                            if (result.failCount > 0)
                            {
                                status += $" {result.failCount} frame calls failed.";
                            }
                        }
                        else if (result.failCount > 0)
                        {
                            status = $"No loads returned. {result.failCount} frame calls failed.";
                        }
                        else
                        {
                            status = "No distributed loads on the requested frames.";
                        }
                    }
                    else
                    {
                        status = $"Returned {result.total} distributed loads for {patternSummary}.";
                        if (result.failCount > 0)
                        {
                            status += $" {result.failCount} frame calls failed.";
                        }
                    }
                }
                else if (result.total == 0 && result.failCount > 0)
                {
                    status = $"No loads returned. {result.failCount} frame calls failed.";
                }
                else if (result.failCount > 0)
                {
                    status = $"Returned {result.total} distributed loads. {result.failCount} frame calls failed.";
                }
                else
                {
                    status = result.total == 0
                        ? "No distributed loads on the requested frames."
                        : $"Returned {result.total} distributed loads.";
                }

                UpdateAndPushOutputs(da, headerTree, valueTree, status, run);
            }
            catch (Exception ex)
            {
                string errorMessage = "Error: " + ex.Message;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
                UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), errorMessage, run);
            }
        }

        private static (int total, int failCount, List<string> frameName, List<string> loadPat, List<int> myType, List<string> cSys, List<int> dir,
            List<double> rd1, List<double> rd2, List<double> dist1, List<double> dist2, List<double> val1, List<double> val2)
            GetFrameDistributed(cSapModel sapModel, IList<string> uniqueNames)
        {
            var frameNameOut = new List<string>();
            var loadPatOut = new List<string>();
            var myTypeOut = new List<int>();
            var cSysOut = new List<string>();
            var dirOut = new List<int>();
            var rd1Out = new List<double>();
            var rd2Out = new List<double>();
            var dist1Out = new List<double>();
            var dist2Out = new List<double>();
            var val1Out = new List<double>();
            var val2Out = new List<double>();

            int total = 0;
            int failCount = 0;

            if (uniqueNames == null || uniqueNames.Count == 0)
            {
                return (0, 0, frameNameOut, loadPatOut, myTypeOut, cSysOut, dirOut, rd1Out, rd2Out, dist1Out, dist2Out, val1Out, val2Out);
            }

            for (int k = 0; k < uniqueNames.Count; k++)
            {
                string name = uniqueNames[k];
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                int n = 0;
                string[] frameName = null;
                string[] loadPat = null;
                string[] cSys = null;
                int[] myType = null;
                int[] dir = null;
                double[] rd1 = null;
                double[] rd2 = null;
                double[] dist1 = null;
                double[] dist2 = null;
                double[] val1 = null;
                double[] val2 = null;

                int ret = sapModel.FrameObj.GetLoadDistributed(
                    name.Trim(),
                    ref n,
                    ref frameName,
                    ref loadPat,
                    ref myType,
                    ref cSys,
                    ref dir,
                    ref rd1,
                    ref rd2,
                    ref dist1,
                    ref dist2,
                    ref val1,
                    ref val2,
                    eItemType.Objects);

                if (ret != 0)
                {
                    failCount++;
                }

                if (ret != 0 || n <= 0)
                {
                    continue;
                }

                total += n;

                for (int i = 0; i < n; i++)
                {
                    frameNameOut.Add(frameName[i]);
                    loadPatOut.Add(loadPat[i]);
                    myTypeOut.Add(myType[i]);
                    string directionReference = ResolveDirectionReference(dir[i]);
                    cSysOut.Add(directionReference);
                    dirOut.Add(dir[i]);
                    rd1Out.Add(rd1[i]);
                    rd2Out.Add(rd2[i]);
                    dist1Out.Add(dist1[i]);
                    dist2Out.Add(dist2[i]);
                    val1Out.Add(val1[i]);
                    val2Out.Add(val2[i]);
                }
            }

            return (total, failCount, frameNameOut, loadPatOut, myTypeOut, cSysOut, dirOut, rd1Out, rd2Out, dist1Out, dist2Out, val1Out, val2Out);
        }

        private static GH_Structure<GH_String> BuildHeaderTree()
        {
            GH_Structure<GH_String> tree = new GH_Structure<GH_String>();
            GH_Path path = new GH_Path(0);

            for (int i = 0; i < HeaderLabels.Length; i++)
            {
                tree.Append(new GH_String(HeaderLabels[i]), path);
            }

            return tree;
        }

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree((int total, int failCount, List<string> frameName, List<string> loadPat,
            List<int> myType, List<string> cSys, List<int> dir, List<double> rd1, List<double> rd2, List<double> dist1, List<double> dist2,
            List<double> val1, List<double> val2) result)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();

            int rowCount = result.frameName.Count;

            for (int col = 0; col < HeaderLabels.Length; col++)
            {
                GH_Path path = new GH_Path(col);
                tree.EnsurePath(path);

                for (int row = 0; row < rowCount; row++)
                {
                    tree.Append(new GH_ObjectWrapper(GetValueByColumn(result, col, row)), path);
                }
            }

            return tree;
        }

        private static object GetValueByColumn((int total, int failCount, List<string> frameName, List<string> loadPat, List<int> myType,
            List<string> cSys, List<int> dir, List<double> rd1, List<double> rd2, List<double> dist1, List<double> dist2, List<double> val1,
            List<double> val2) result, int columnIndex, int rowIndex)
        {
            switch (columnIndex)
            {
                case 0:
                    return result.frameName[rowIndex];
                case 1:
                    return result.loadPat[rowIndex];
                case 2:
                    return result.myType[rowIndex];
                case 3:
                    return result.cSys[rowIndex];
                case 4:
                    return result.dir[rowIndex];
                case 5:
                    return result.rd1[rowIndex];
                case 6:
                    return result.rd2[rowIndex];
                case 7:
                    return result.dist1[rowIndex];
                case 8:
                    return result.dist2[rowIndex];
                case 9:
                    return result.val1[rowIndex];
                case 10:
                    return result.val2[rowIndex];
                default:
                    throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
        }

        private static string ResolveDirectionReference(int direction)
        {
            return Math.Abs(direction) < 10 ? "Local" : "Global";
        }

        private static (int total, int failCount, List<string> frameName, List<string> loadPat, List<int> myType, List<string> cSys,
            List<int> dir, List<double> rd1, List<double> rd2, List<double> dist1, List<double> dist2, List<double> val1, List<double>
                val2) FilterByLoadPattern(
            (int total, int failCount, List<string> frameName, List<string> loadPat, List<int> myType, List<string> cSys, List<int>
                dir, List<double> rd1, List<double> rd2, List<double> dist1, List<double> dist2, List<double> val1, List<double> val2)
                result,
            IReadOnlyCollection<string> loadPatternFilters)
        {
            if (loadPatternFilters == null || loadPatternFilters.Count == 0)
            {
                return result;
            }

            var comparer = StringComparer.OrdinalIgnoreCase;
            HashSet<string> filterSet = new HashSet<string>(loadPatternFilters, comparer);

            var frameNameOut = new List<string>();
            var loadPatOut = new List<string>();
            var myTypeOut = new List<int>();
            var cSysOut = new List<string>();
            var dirOut = new List<int>();
            var rd1Out = new List<double>();
            var rd2Out = new List<double>();
            var dist1Out = new List<double>();
            var dist2Out = new List<double>();
            var val1Out = new List<double>();
            var val2Out = new List<double>();

            for (int i = 0; i < result.frameName.Count; i++)
            {
                if (!filterSet.Contains(result.loadPat[i]))
                {
                    continue;
                }

                frameNameOut.Add(result.frameName[i]);
                loadPatOut.Add(result.loadPat[i]);
                myTypeOut.Add(result.myType[i]);
                cSysOut.Add(result.cSys[i]);
                dirOut.Add(result.dir[i]);
                rd1Out.Add(result.rd1[i]);
                rd2Out.Add(result.rd2[i]);
                dist1Out.Add(result.dist1[i]);
                dist2Out.Add(result.dist2[i]);
                val1Out.Add(result.val1[i]);
                val2Out.Add(result.val2[i]);
            }

            return (frameNameOut.Count, result.failCount, frameNameOut, loadPatOut, myTypeOut, cSysOut, dirOut, rd1Out, rd2Out,
                dist1Out, dist2Out, val1Out, val2Out);
        }

        private static string FormatLoadPatternSummary(IReadOnlyList<string> filters)
        {
            if (filters == null || filters.Count == 0)
            {
                return string.Empty;
            }

            if (filters.Count == 1)
            {
                return $"load pattern \"{filters[0]}\"";
            }

            return $"load patterns ({string.Join(", ", filters)})";
        }

        private void UpdateAndPushOutputs(IGH_DataAccess da, GH_Structure<GH_String> headerTree, GH_Structure<GH_ObjectWrapper> valueTree,
            string message, bool currentRunState)
        {
            lastHeaderTree = headerTree.Duplicate();
            lastValueTree = valueTree.Duplicate();
            lastMessage = message;

            da.SetDataTree(0, headerTree);
            da.SetDataTree(1, valueTree);
            da.SetData(2, message);

            lastRun = currentRunState;
        }
    }
}