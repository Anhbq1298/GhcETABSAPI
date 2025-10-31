// -------------------------------------------------------------
// Component : Get Point Info
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 8.0 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)
// Panel     : "MGT" / "2.0 Point Object Modelling"
// -------------------------------------------------------------
// Inputs (ordered):
//   0) sapModel    (ETABSv1.cSapModel, item)  ETABS model from the Attach component.
//   1) pointNames  (string, list)  Optional specific point names. Leave blank to query all points in the model.
//
// Outputs:
//   0) headers     (text, tree)  Single branch describing the returned columns.
//   1) values      (generic, tree) Column-wise point data aligned with the header order.
//   2) msg         (text, item)  Status / diagnostics string.
//
// Behavior Notes:
//   • When no point names are supplied, the component queries all available point object names.
//   • Duplicated or blank point names are removed while preserving the first occurrence ordering.
//   • Each column branch lines up with the header labels: UniqueName, Label, Story, X, Y, Z.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using ETABSv1;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using static MGT.ComponentShared;

namespace MGT
{
    public class GhcGetPointInfo : GH_Component
    {
        private static readonly string[] HeaderLabels =
        {
            "UniqueName",
            "Label",
            "Story",
            "X",
            "Y",
            "Z"
        };

        public GhcGetPointInfo()
          : base(
                "Get Point Info",
                "GetPointInfo",
                "Retrieve key information about ETABS point objects (label, story, coordinates).",
                "MGT",
                "2.0 Point Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("94bc99e0-6f04-4b6c-8736-c36b43336ea0");

        protected override System.Drawing.Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            int pointNamesIndex = p.AddTextParameter(
                "pointNames",
                "pointNames",
                "Optional specific point names. Leave blank to pull all point objects.",
                GH_ParamAccess.list);
            p[pointNamesIndex].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("headers", "headers", "Header labels describing each value column.", GH_ParamAccess.tree);
            p.AddGenericParameter("values", "values", "Point info values organised per column.", GH_ParamAccess.tree);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            cSapModel sapModel = null;
            List<string> requestedNames = new List<string>();

            da.GetData(0, ref sapModel);
            da.GetDataList(1, requestedNames);

            if (sapModel == null)
            {
                string warning = "sapModel is null. Wire it from the Attach component.";
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warning);
                UpdateOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), warning);
                return;
            }

            int duplicateNameCount;
            List<string> normalizedRequested = NormalizePointNames(requestedNames, out duplicateNameCount);
            if (duplicateNameCount > 0)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Remark, $"{duplicateNameCount} duplicate point name(s) ignored.");
            }

            List<string> targetNames;
            bool usingRequestedList = normalizedRequested.Count > 0;

            try
            {
                targetNames = usingRequestedList
                    ? normalizedRequested
                    : GetAllPointNames(sapModel);
            }
            catch (Exception ex)
            {
                string errorMessage = "Error: " + ex.Message;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
                UpdateOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), errorMessage);
                return;
            }

            if (targetNames == null || targetNames.Count == 0)
            {
                string idleMessage = usingRequestedList
                    ? "No valid point names supplied."
                    : "Model returned no point names.";
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, idleMessage);
                UpdateOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), idleMessage);
                return;
            }

            List<string> runtimeWarnings = new List<string>();
            PointInfoResult result = FetchPointInfo(sapModel, targetNames, runtimeWarnings);

            GH_Structure<GH_String> headerTree = BuildHeaderTree();
            GH_Structure<GH_ObjectWrapper> valueTree = BuildValueTree(result);

            foreach (string warning in Deduplicate(runtimeWarnings))
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warning);
            }

            string sourceLabel = usingRequestedList ? "requested list" : "model name list";
            string message;

            if (result.Count == 0)
            {
                message = "No point info returned.";
            }
            else if (result.SuccessCount == result.Count)
            {
                message = $"Retrieved info for all {result.Count} point(s) from the {sourceLabel}.";
            }
            else if (result.SuccessCount == 0)
            {
                message = $"Failed to retrieve info for the {result.Count} point(s) from the {sourceLabel}.";
            }
            else
            {
                message = $"Retrieved info for {result.SuccessCount} of {result.Count} point(s) from the {sourceLabel}.";
            }

            UpdateOutputs(da, headerTree, valueTree, message);
        }

        private static List<string> NormalizePointNames(IList<string> source, out int duplicates)
        {
            duplicates = 0;

            if (source == null)
            {
                return new List<string>();
            }

            int validCount = 0;
            for (int i = 0; i < source.Count; i++)
            {
                string entry = source[i];
                if (!string.IsNullOrWhiteSpace(entry) && entry.Trim().Length > 0)
                {
                    validCount++;
                }
            }

            List<string> normalized = NormalizeDistinct(source);
            duplicates = Math.Max(0, validCount - normalized.Count);
            return normalized;
        }

        private static List<string> GetAllPointNames(cSapModel sapModel)
        {
            if (sapModel == null)
            {
                return new List<string>();
            }

            int count = 0;
            string[] names = null;

            int ret = sapModel.PointObj.GetNameList(ref count, ref names);
            if (ret != 0)
            {
                throw new InvalidOperationException($"PointObj.GetNameList failed with error code {ret}.");
            }

            return NormalizeDistinct(names);
        }

        private static PointInfoResult FetchPointInfo(cSapModel sapModel, IList<string> pointNames, List<string> warnings)
        {
            PointInfoResult result = new PointInfoResult();

            if (sapModel == null || pointNames == null)
            {
                return result;
            }

            for (int i = 0; i < pointNames.Count; i++)
            {
                string name = pointNames[i];
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                string trimmedName = name.Trim();
                string label = string.Empty;
                string story = string.Empty;
                double x = double.NaN;
                double y = double.NaN;
                double z = double.NaN;

                bool success = true;

                try
                {
                    int retCoord = sapModel.PointObj.GetCoordCartesian(trimmedName, ref x, ref y, ref z);
                    if (retCoord != 0)
                    {
                        warnings?.Add($"Point \"{trimmedName}\": GetCoordCartesian returned {retCoord}.");
                        x = y = z = double.NaN;
                        success = false;
                    }
                }
                catch (Exception ex)
                {
                    warnings?.Add($"Point \"{trimmedName}\": GetCoordCartesian exception - {ex.Message}");
                    x = y = z = double.NaN;
                    success = false;
                }

                try
                {
                    int retLabel = sapModel.PointObj.GetLabelFromName(trimmedName, ref label, ref story);
                    if (retLabel != 0)
                    {
                        warnings?.Add($"Point \"{trimmedName}\": GetLabelFromName returned {retLabel}.");
                        label = string.Empty;
                        story = string.Empty;
                        success = false;
                    }
                }
                catch (Exception ex)
                {
                    warnings?.Add($"Point \"{trimmedName}\": GetLabelFromName exception - {ex.Message}");
                    label = string.Empty;
                    story = string.Empty;
                    success = false;
                }

                result.UniqueName.Add(trimmedName);
                result.Label.Add(label ?? string.Empty);
                result.Story.Add(story ?? string.Empty);
                result.X.Add(SanitizeDouble(x));
                result.Y.Add(SanitizeDouble(y));
                result.Z.Add(SanitizeDouble(z));
                result.Success.Add(success);
            }

            return result;
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

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(PointInfoResult result)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();
            int rowCount = result?.Count ?? 0;

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

        private static object GetValueByColumn(PointInfoResult result, int columnIndex, int rowIndex)
        {
            switch (columnIndex)
            {
                case 0:
                    return result.UniqueName[rowIndex];
                case 1:
                    return result.Label[rowIndex];
                case 2:
                    return result.Story[rowIndex];
                case 3:
                    return result.X[rowIndex];
                case 4:
                    return result.Y[rowIndex];
                case 5:
                    return result.Z[rowIndex];
                default:
                    throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
        }

        private static double SanitizeDouble(double value)
        {
            if (IsInvalidNumber(value))
            {
                return double.NaN;
            }

            return value;
        }

        private static IEnumerable<string> Deduplicate(IEnumerable<string> warnings)
        {
            if (warnings == null)
            {
                yield break;
            }

            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string warning in warnings)
            {
                if (string.IsNullOrWhiteSpace(warning))
                {
                    continue;
                }

                if (seen.Add(warning))
                {
                    yield return warning;
                }
            }
        }

        private void UpdateOutputs(IGH_DataAccess da, GH_Structure<GH_String> headerTree, GH_Structure<GH_ObjectWrapper> valueTree, string message)
        {
            da.SetDataTree(0, headerTree);
            da.SetDataTree(1, valueTree);
            da.SetData(2, message);
        }

        private sealed class PointInfoResult
        {
            internal List<string> UniqueName { get; } = new List<string>();
            internal List<string> Label { get; } = new List<string>();
            internal List<string> Story { get; } = new List<string>();
            internal List<double> X { get; } = new List<double>();
            internal List<double> Y { get; } = new List<double>();
            internal List<double> Z { get; } = new List<double>();
            internal List<bool> Success { get; } = new List<bool>();

            internal int Count => UniqueName.Count;

            internal int SuccessCount
            {
                get
                {
                    int total = 0;
                    for (int i = 0; i < Success.Count; i++)
                    {
                        if (Success[i])
                        {
                            total++;
                        }
                    }

                    return total;
                }
            }
        }
    }
}
