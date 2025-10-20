using System;
using System.Collections.Generic;
using Grasshopper.Kernel;
using ETABSv1;

namespace GhcETABSAPI
{
    public class GhcGetLoadDistOnFrames : GH_Component
    {
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
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            p.AddTextParameter(
                "frameNames",
                "frameNames",
                "Frame object names to query. Blank entries are ignored. If empty, returns zero results.",
                GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddIntegerParameter("count", "count", "Number of distributed load assignments returned.", GH_ParamAccess.item);
            p.AddTextParameter("frameName", "frameName", "Frame object name for each distributed load.", GH_ParamAccess.list);
            p.AddTextParameter("loadPattern", "loadPattern", "Load pattern name.", GH_ParamAccess.list);
            p.AddIntegerParameter("type", "type", "Distributed load type (ETABS enumeration value).", GH_ParamAccess.list);
            p.AddTextParameter("cSys", "cSys", "Coordinate system used for the assignment.", GH_ParamAccess.list);
            p.AddIntegerParameter("direction", "direction", "Load direction (ETABS enumeration value).", GH_ParamAccess.list);
            p.AddNumberParameter("relDist1", "relDist1", "Relative distance 1.", GH_ParamAccess.list);
            p.AddNumberParameter("relDist2", "relDist2", "Relative distance 2.", GH_ParamAccess.list);
            p.AddNumberParameter("dist1", "dist1", "Absolute distance 1.", GH_ParamAccess.list);
            p.AddNumberParameter("dist2", "dist2", "Absolute distance 2.", GH_ParamAccess.list);
            p.AddNumberParameter("value1", "value1", "Load value 1.", GH_ParamAccess.list);
            p.AddNumberParameter("value2", "value2", "Load value 2.", GH_ParamAccess.list);
            p.AddTextParameter("msg", "msg", "Status / diagnostic message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            cSapModel sapModel = null;
            List<string> frameNames = new List<string>();

            if (!da.GetData(0, ref sapModel) || sapModel == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No ETABS SapModel provided.");
                PushOutputs(da, 0, new List<string>(), new List<string>(), new List<int>(), new List<string>(), new List<int>(),
                    new List<double>(), new List<double>(), new List<double>(), new List<double>(), new List<double>(), new List<double>(),
                    "sapModel is null.");
                return;
            }

            da.GetDataList(1, frameNames);

            try
            {
                List<string> trimmed = new List<string>();
                if (frameNames != null)
                {
                    for (int i = 0; i < frameNames.Count; i++)
                    {
                        string nm = frameNames[i];
                        if (!string.IsNullOrWhiteSpace(nm))
                        {
                            trimmed.Add(nm.Trim());
                        }
                    }
                }

                var result = GetFrameDistributed(sapModel, trimmed);

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
                    status = $"Returned {result.total} distributed loads.";
                }

                PushOutputs(
                    da,
                    result.total,
                    result.frameName,
                    result.loadPat,
                    result.myType,
                    result.cSys,
                    result.dir,
                    result.rd1,
                    result.rd2,
                    result.dist1,
                    result.dist2,
                    result.val1,
                    result.val2,
                    status);
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
                PushOutputs(da, 0, new List<string>(), new List<string>(), new List<int>(), new List<string>(), new List<int>(),
                    new List<double>(), new List<double>(), new List<double>(), new List<double>(), new List<double>(), new List<double>(),
                    "Error: " + ex.Message);
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

        private static void PushOutputs(
            IGH_DataAccess da,
            int total,
            IList<string> frameName,
            IList<string> loadPat,
            IList<int> myType,
            IList<string> cSys,
            IList<int> dir,
            IList<double> rd1,
            IList<double> rd2,
            IList<double> dist1,
            IList<double> dist2,
            IList<double> val1,
            IList<double> val2,
            string message)
        {
            da.SetData(0, total);
            da.SetDataList(1, frameName);
            da.SetDataList(2, loadPat);
            da.SetDataList(3, myType);
            da.SetDataList(4, cSys);
            da.SetDataList(5, dir);
            da.SetDataList(6, rd1);
            da.SetDataList(7, rd2);
            da.SetDataList(8, dist1);
            da.SetDataList(9, dist2);
            da.SetDataList(10, val1);
            da.SetDataList(11, val2);
            da.SetData(12, message);
        }
    }
}

