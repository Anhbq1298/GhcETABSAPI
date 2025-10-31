// -------------------------------------------------------------
// Component : Get All Frame Info
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)  [Embed Interop Types = False]
// Panel     : "MGT" / "2.0 Frame Object Modelling"
// GUID      : b8c5cf76-2b02-44ad-885d-1be3cc8d2b5c
// -------------------------------------------------------------
// Inputs (ordered):
//   0) sapModel    (ETABSv1.cSapModel, item)  ETABS model from your Attach component.
//   1) coordinateSystem (string, item) OPTIONAL. Coordinate system passed to FrameObj.GetAllFrames. Defaults to "Global".
//
// Outputs:
//   0) header  (text, tree)   Single branch of column labels.
//   1) values  (generic, tree) Column-wise branches aligned to header order.
//   2) msg     (text, item)   Status / diagnostics.
//
// Behavior Notes:
//   • Retrieves all frame objects at once via FrameObj.GetAllFrames.
//   • Values tree is column-oriented to match header labels:
//       [0] UniqueName, [1] Section, [2] PointName1, [3] PointName2,
//       [4] Point1X, [5] Point1Y, [6] Point1Z, [7] Point2X, [8] Point2Y, [9] Point2Z.
// -------------------------------------------------------------

using ETABSv1;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using System;
using System.Collections.Generic;

namespace MGT
{
    public class GhcGetAllFrameInfo : GH_Component
    {
        private static readonly string[] HeaderLabels =
        {
            "UniqueName",
            "Section",
            "PointName1",
            "PointName2",
            "Point1X",
            "Point1Y",
            "Point1Z",
            "Point2X",
            "Point2Y",
            "Point2Z"
        };

        public GhcGetAllFrameInfo()
          : base(
                "Get All Frame Info",
                "GetAllFrames",
                "Retrieve key information for every ETABS frame object via FrameObj.GetAllFrames.",
                "MGT",
                "2.0 Frame Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("b8c5cf76-2b02-44ad-885d-1be3cc8d2b5c");

        protected override System.Drawing.Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            int csysIndex = p.AddTextParameter(
                "coordinateSystem",
                "csys",
                "Optional coordinate system passed to FrameObj.GetAllFrames. Leave blank for Global.",
                GH_ParamAccess.item);
            p[csysIndex].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("headers", "headers", "Header labels describing each value column.", GH_ParamAccess.tree);
            p.AddGenericParameter("values", "values", "Frame info rows. Each branch matches the header order.", GH_ParamAccess.tree);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            cSapModel sapModel = null;
            string coordinateSystem = "Global";

            da.GetData(0, ref sapModel);
            da.GetData(1, ref coordinateSystem);

            if (sapModel == null)
            {
                string warning = "sapModel is null. Wire it from the Attach component.";
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, warning);
                UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), warning);
                return;
            }

            string resolvedCoordinateSystem = string.IsNullOrWhiteSpace(coordinateSystem)
                ? "Global"
                : coordinateSystem.Trim();

            try
            {
                FrameInfoResult result = GetAllFrameInfo(sapModel, resolvedCoordinateSystem);

                GH_Structure<GH_String> headerTree = BuildHeaderTree();
                GH_Structure<GH_ObjectWrapper> valueTree = BuildValueTree(result);

                string message;
                if (result.ErrorCode != 0)
                {
                    message = $"FrameObj.GetAllFrames failed with error code {result.ErrorCode}.";
                }
                else if (result.Count == 0)
                {
                    message = "No frame objects returned.";
                }
                else
                {
                    message = $"Returned {result.Count} frame objects using the \"{resolvedCoordinateSystem}\" coordinate system.";
                }

                UpdateAndPushOutputs(da, headerTree, valueTree, message);
            }
            catch (Exception ex)
            {
                string errorMessage = "Error: " + ex.Message;
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
                UpdateAndPushOutputs(da, BuildHeaderTree(), new GH_Structure<GH_ObjectWrapper>(), errorMessage);
            }
        }

        private static FrameInfoResult GetAllFrameInfo(cSapModel sapModel, string coordinateSystem)
        {
            FrameInfoResult result = new FrameInfoResult
            {
                CoordinateSystem = coordinateSystem ?? "Global"
            };

            if (sapModel == null)
            {
                return result;
            }

            int numberNames = 0;
            string[] myName = null;
            string[] propName = null;
            string[] storyName = null;
            string[] pointName1 = null;
            string[] pointName2 = null;
            double[] point1X = null;
            double[] point1Y = null;
            double[] point1Z = null;
            double[] point2X = null;
            double[] point2Y = null;
            double[] point2Z = null;
            double[] angle = null;
            double[] offset1X = null;
            double[] offset2X = null;
            double[] offset1Y = null;
            double[] offset2Y = null;
            double[] offset1Z = null;
            double[] offset2Z = null;
            int[] cardinalPoint = null;

            int ret = sapModel.FrameObj.GetAllFrames(
                ref numberNames,
                ref myName,
                ref propName,
                ref storyName,
                ref pointName1,
                ref pointName2,
                ref point1X,
                ref point1Y,
                ref point1Z,
                ref point2X,
                ref point2Y,
                ref point2Z,
                ref angle,
                ref offset1X,
                ref offset2X,
                ref offset1Y,
                ref offset2Y,
                ref offset1Z,
                ref offset2Z,
                ref cardinalPoint,
                result.CoordinateSystem);

            result.ErrorCode = ret;

            if (ret != 0 || numberNames <= 0)
            {
                return result;
            }

            for (int i = 0; i < numberNames; i++)
            {
                result.UniqueName.Add(SafeGet(myName, i));
                result.Section.Add(SafeGet(propName, i));
                result.PointName1.Add(SafeGet(pointName1, i));
                result.PointName2.Add(SafeGet(pointName2, i));
                result.Point1X.Add(SafeGet(point1X, i));
                result.Point1Y.Add(SafeGet(point1Y, i));
                result.Point1Z.Add(SafeGet(point1Z, i));
                result.Point2X.Add(SafeGet(point2X, i));
                result.Point2Y.Add(SafeGet(point2Y, i));
                result.Point2Z.Add(SafeGet(point2Z, i));
            }

            result.Count = result.UniqueName.Count;
            return result;
        }

        private static string SafeGet(string[] source, int index)
        {
            if (source == null || index < 0 || index >= source.Length)
            {
                return string.Empty;
            }

            return source[index] ?? string.Empty;
        }

        private static double SafeGet(double[] source, int index)
        {
            if (source == null || index < 0 || index >= source.Length)
            {
                return double.NaN;
            }

            double value = source[index];
            return double.IsNaN(value) || double.IsInfinity(value) ? double.NaN : value;
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

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(FrameInfoResult result)
        {
            GH_Structure<GH_ObjectWrapper> tree = new GH_Structure<GH_ObjectWrapper>();
            int rowCount = result?.UniqueName.Count ?? 0;

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

        private static object GetValueByColumn(FrameInfoResult result, int columnIndex, int rowIndex)
        {
            switch (columnIndex)
            {
                case 0:
                    return result.UniqueName[rowIndex];
                case 1:
                    return result.Section[rowIndex];
                case 2:
                    return result.PointName1[rowIndex];
                case 3:
                    return result.PointName2[rowIndex];
                case 4:
                    return result.Point1X[rowIndex];
                case 5:
                    return result.Point1Y[rowIndex];
                case 6:
                    return result.Point1Z[rowIndex];
                case 7:
                    return result.Point2X[rowIndex];
                case 8:
                    return result.Point2Y[rowIndex];
                case 9:
                    return result.Point2Z[rowIndex];
                default:
                    throw new ArgumentOutOfRangeException(nameof(columnIndex));
            }
        }

        private void UpdateAndPushOutputs(IGH_DataAccess da, GH_Structure<GH_String> headerTree, GH_Structure<GH_ObjectWrapper> valueTree, string message)
        {
            da.SetDataTree(0, headerTree);
            da.SetDataTree(1, valueTree);
            da.SetData(2, message);
        }

        private sealed class FrameInfoResult
        {
            internal int Count { get; set; }
            internal int ErrorCode { get; set; }
            internal string CoordinateSystem { get; set; } = "Global";
            internal List<string> UniqueName { get; } = new List<string>();
            internal List<string> Section { get; } = new List<string>();
            internal List<string> PointName1 { get; } = new List<string>();
            internal List<string> PointName2 { get; } = new List<string>();
            internal List<double> Point1X { get; } = new List<double>();
            internal List<double> Point1Y { get; } = new List<double>();
            internal List<double> Point1Z { get; } = new List<double>();
            internal List<double> Point2X { get; } = new List<double>();
            internal List<double> Point2Y { get; } = new List<double>();
            internal List<double> Point2Z { get; } = new List<double>();
        }
    }
}
