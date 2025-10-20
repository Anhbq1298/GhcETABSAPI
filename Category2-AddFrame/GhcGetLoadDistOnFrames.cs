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
            p.AddTextParameter(
                "frameNames",
                "frameNames",
                "Frame object names to query. Blank entries are ignored. If empty, returns zero results.",
                GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("header", "header", "Header labels describing each value column.", GH_ParamAccess.tree);
            p.AddGenericParameter("values", "values", "Distributed load rows. Each branch matches the header order.", GH_ParamAccess.tree);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sapModel = null;
            List<string> frameNames = new List<string>();

            da.GetData(0, ref run);
            da.GetData(1, ref sapModel);
            da.GetDataList(2, frameNames);

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

                var result = GetFrameDistributed(sapModel, trimmed);

                GH_Structure<GH_String> headerTree = BuildHeaderTree();
                GH_Structure<GH_ObjectWrapper> valueTree = BuildValueTree(result);

                string status;
                if (trimmed.Count == 0)
                {
                    status = "No valid frame names provided.";
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
                    cSysOut.Add(cSys[i]);
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
            for (int i = 0; i < rowCount; i++)
            {
                GH_Path path = new GH_Path(i);

                tree.Append(new GH_ObjectWrapper(result.frameName[i]), path);
                tree.Append(new GH_ObjectWrapper(result.loadPat[i]), path);
                tree.Append(new GH_ObjectWrapper(result.myType[i]), path);
                tree.Append(new GH_ObjectWrapper(result.cSys[i]), path);
                tree.Append(new GH_ObjectWrapper(result.dir[i]), path);
                tree.Append(new GH_ObjectWrapper(result.rd1[i]), path);
                tree.Append(new GH_ObjectWrapper(result.rd2[i]), path);
                tree.Append(new GH_ObjectWrapper(result.dist1[i]), path);
                tree.Append(new GH_ObjectWrapper(result.dist2[i]), path);
                tree.Append(new GH_ObjectWrapper(result.val1[i]), path);
                tree.Append(new GH_ObjectWrapper(result.val2[i]), path);
            }

            return tree;
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

