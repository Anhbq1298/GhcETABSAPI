// -------------------------------------------------------------
// Component : Get Area Uniform Loads (per-object query)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 8.0 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)  [Embed Interop Types = False]
// Panel     : "MGT" / "3.0 Shell (Area) Object Modelling"
// GUID      : 77f15ab0-1587-4e9b-8a47-335c50a62ddb
// -------------------------------------------------------------
// Inputs (ordered):
//   0) sapModel    (ETABSv1.cSapModel, item)  ETABS model from your Attach component.
//   1) areaNames   (string, list) Area object names to query. Blank/dup ignored (case-insensitive). Leave empty to query all.
//   2) loadPattern (string, list) OPTIONAL filters. If UNCONNECTED or empty → treated as null (no filter).
//
// Outputs:
//   0) header  (text, tree)   Single branch of column labels.
//   1) values  (generic, tree) Column-wise branches aligned to header order.
//   2) msg     (text, item)   Status / diagnostics.
//
// Behavior Notes:
//   • areaNames are trimmed and de-duplicated (OrdinalIgnoreCase).
//   • loadPattern is OPTIONAL; null ⇒ return all patterns; when provided, filter is case-insensitive.
//   • Values tree is column-oriented to match header labels:
//       [0] AreaName, [1] LoadPattern, [2] CoordinateSystem, [3] Direction, [4] Value.
//   • Uses per-object mode: AreaObj.GetLoadUniform(..., eItemType.Objects).
//   • CoordinateSystem falls back to Local/Global based on Direction when API returns blank.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using ETABSv1;
using static MGT.ComponentShared;

namespace MGT
{
    public class GhcGetLoadUniformOnAreas : GH_Component
    {
        private static readonly string[] HeaderLabels =
        {
            "AreaName",
            "LoadPattern",
            "CoordinateSystem",
            "Direction",
            "Value"
        };

        public GhcGetLoadUniformOnAreas()
          : base(
                "Get Area Uniform Loads",
                "GetAreaLoads",
                "Query uniform surface loads assigned to ETABS area objects (per object mode).\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                "MGT",
                "4.0 Area Object Modelling"  )
        {
        }

        public override Guid ComponentGuid => new Guid("77f15ab0-1587-4e9b-8a47-335c50a62ddb");

        protected override System.Drawing.Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            int areaNameIndex = p.AddTextParameter(
                "areaNames",
                "areaNames",
                "Area object names to query. Blank entries are ignored. Leave empty to query every area object.",
                GH_ParamAccess.list);
            p[areaNameIndex].Optional = true;

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
            p.AddGenericParameter("values", "values", "Uniform load rows. Each branch matches the header order.", GH_ParamAccess.tree);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            cSapModel sapModel = null;
            List<string> areaNames = new List<string>();
            List<string> loadPatternFilters = new List<string>();

            da.GetData(0, ref sapModel);
            da.GetDataList(1, areaNames);
            da.GetDataList(2, loadPatternFilters);

            if (sapModel == null)
            {
                string warning = "sapModel is null. Wire it from the Attach component.";
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warning);
                UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), warning);
                return;
            }

            try
            {
                List<string> trimmed = NormalizeDistinct(areaNames);
                bool autoFilledAllAreas = false;
                bool attemptedAutoFill = trimmed.Count == 0;

                if (attemptedAutoFill)
                {
                    HashSet<string> allAreas = TryGetExistingAreaNames(sapModel);
                    if (allAreas != null && allAreas.Count > 0)
                    {
                        trimmed.AddRange(allAreas);
                        trimmed.Sort(StringComparer.OrdinalIgnoreCase);
                        autoFilledAllAreas = true;
                    }
                }

                if (trimmed.Count == 0)
                {
                    string noAreasMessage = attemptedAutoFill
                        ? "No area objects exist in the model."
                        : "No valid area names provided.";

                    UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), noAreasMessage);
                    return;
                }

                var rawResult = GetAreaUniform(sapModel, trimmed);

                List<string> trimmedFilters = NormalizeDistinct(loadPatternFilters);
                bool hasFilters = trimmedFilters.Count > 0;
                var result = FilterByLoadPattern(rawResult, trimmedFilters);

                GH_Structure<GH_String> headerTree = BuildHeaderTree();
                GH_Structure<GH_ObjectWrapper> valueTree = BuildValueTree(result);

                string zeroTarget = autoFilledAllAreas ? "any area objects" : "the requested areas";
                string positiveTarget = autoFilledAllAreas ? "all area objects" : "the requested areas";

                string status;
                if (hasFilters)
                {
                    string patternSummary = FormatLoadPatternSummary(trimmedFilters);

                    if (result.total == 0)
                    {
                        if (rawResult.total > 0)
                        {
                            status = $"No uniform area loads matched {patternSummary} on {zeroTarget}.";
                            if (result.failCount > 0)
                            {
                                status += $" {result.failCount} area calls failed.";
                            }
                        }
                        else if (result.failCount > 0)
                        {
                            status = $"No loads returned. {result.failCount} area calls failed.";
                        }
                        else
                        {
                            status = $"No uniform area loads on {zeroTarget}.";
                        }
                    }
                    else
                    {
                        status = $"Returned {result.total} uniform area loads for {patternSummary} on {positiveTarget}.";
                        if (result.failCount > 0)
                        {
                            status += $" {result.failCount} area calls failed.";
                        }
                    }
                }
                else if (result.total == 0 && result.failCount > 0)
                {
                    status = $"No loads returned. {result.failCount} area calls failed.";
                }
                else if (result.failCount > 0)
                {
                    status = $"Returned {result.total} uniform area loads. {result.failCount} area calls failed.";
                }
                else
                {
                    status = result.total == 0
                        ? $"No uniform area loads on {zeroTarget}."
                        : $"Returned {result.total} uniform area loads on {positiveTarget}.";
                }

                UpdateAndPushOutputs(da, headerTree, valueTree, status);
            }
            catch (Exception ex)
            {
                string errorMessage = "Error: " + ex.Message;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
                UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), errorMessage);
            }
        }

        private static (int total, int failCount, List<string> areaName, List<string> loadPat, List<string> cSys, List<int> dir, List<double> value)
            GetAreaUniform(cSapModel sapModel, IList<string> uniqueNames)
        {
            var areaNameOut = new List<string>();
            var loadPatOut = new List<string>();
            var cSysOut = new List<string>();
            var dirOut = new List<int>();
            var valueOut = new List<double>();

            if (sapModel == null || uniqueNames == null || uniqueNames.Count == 0)
            {
                return (0, 0, areaNameOut, loadPatOut, cSysOut, dirOut, valueOut);
            }

            int total = 0;
            int failCount = 0;

            for (int k = 0; k < uniqueNames.Count; k++)
            {
                string name = uniqueNames[k];
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                int n = 0;
                string[] areaName = null;
                string[] loadPat = null;
                string[] cSys = null;
                int[] dir = null;
                double[] value = null;

                int ret = sapModel.AreaObj.GetLoadUniform(
                    name.Trim(),
                    ref n,
                    ref areaName,
                    ref loadPat,
                    ref cSys,
                    ref dir,
                    ref value,
                    eItemType.Objects);

                if (ret != 0)
                {
                    failCount++;
                }

                if (ret != 0 || n <= 0)
                {
                    continue;
                }

                if (areaName == null || loadPat == null || cSys == null || dir == null || value == null)
                {
                    continue;
                }

                total += n;

                for (int i = 0; i < n; i++)
                {
                    string resolvedAreaName = areaName[i];
                    int direction = dir[i];
                    string resolvedCoordinate = ResolveDirectionReferenceArea(direction);

                    areaNameOut.Add(resolvedAreaName);
                    loadPatOut.Add(loadPat[i]);
                    cSysOut.Add(resolvedCoordinate);
                    dirOut.Add(direction);
                    valueOut.Add(value[i]);
                }
            }

            return (total, failCount, areaNameOut, loadPatOut, cSysOut, dirOut, valueOut);
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

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(
            (int total, int failCount, List<string> areaName, List<string> loadPat, List<string> cSys, List<int> dir, List<double> value) result)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();

            int rowCount = result.areaName.Count;

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

        private static object GetValueByColumn(
            (int total, int failCount, List<string> areaName, List<string> loadPat, List<string> cSys, List<int> dir, List<double> value) result,
            int columnIndex,
            int rowIndex)
        {
            switch (columnIndex)
            {
                case 0:
                    return result.areaName[rowIndex];
                case 1:
                    return result.loadPat[rowIndex];
                case 2:
                    return result.cSys[rowIndex];
                case 3:
                    return result.dir[rowIndex];
                case 4:
                    return result.value[rowIndex];
                default:
                    throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
        }

        private static (int total, int failCount, List<string> areaName, List<string> loadPat, List<string> cSys, List<int> dir, List<double> value)
            FilterByLoadPattern(
                (int total, int failCount, List<string> areaName, List<string> loadPat, List<string> cSys, List<int> dir, List<double> value) result,
                IReadOnlyCollection<string> loadPatternFilters)
        {
            if (loadPatternFilters == null || loadPatternFilters.Count == 0)
            {
                return result;
            }

            var comparer = StringComparer.OrdinalIgnoreCase;
            HashSet<string> filterSet = new HashSet<string>(loadPatternFilters, comparer);

            var areaNameOut = new List<string>();
            var loadPatOut = new List<string>();
            var cSysOut = new List<string>();
            var dirOut = new List<int>();
            var valueOut = new List<double>();

            for (int i = 0; i < result.areaName.Count; i++)
            {
                if (!filterSet.Contains(result.loadPat[i]))
                {
                    continue;
                }

                areaNameOut.Add(result.areaName[i]);
                loadPatOut.Add(result.loadPat[i]);
                cSysOut.Add(result.cSys[i]);
                dirOut.Add(result.dir[i]);
                valueOut.Add(result.value[i]);
            }

            return (areaNameOut.Count, result.failCount, areaNameOut, loadPatOut, cSysOut, dirOut, valueOut);
        }

        private void UpdateAndPushOutputs(
            IGH_DataAccess da,
            GH_Structure<GH_String> headerTree,
            GH_Structure<GH_ObjectWrapper> valueTree,
            string message)
        {
            da.SetDataTree(0, headerTree);
            da.SetDataTree(1, valueTree);
            da.SetData(2, message);
        }
    }
}
