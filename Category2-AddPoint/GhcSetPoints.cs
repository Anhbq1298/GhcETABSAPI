// -------------------------------------------------------------
// Component : ETABS Set Points from Grasshopper
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 8.0 (x64)
// Depends   : RhinoCommon, Grasshopper, ETABSv1 (COM)
// Panel     : "MGT" / "2.0 Point Object Modelling"
// -------------------------------------------------------------
// Inputs (ordered):
//   0) add         (bool, item)    Rising-edge trigger; executes on False→True transition.
//   1) sapModel    (ETABSv1.cSapModel, item)  ETABS model from the Attach component.
//   2) points      (Rhino.Geometry.Point3d, list)  Grasshopper points to transfer to ETABS.
//   3) uniqueNames (string, list)  Optional point names; matched by index, with blanks letting ETABS auto-name.
//
// Outputs:
//   0) msg         (string, item)  Summary/status message.
//   1) etabsNames  (string, list)  ETABS point object names that were created/updated.
//
// Behavior Notes:
//   • Rising-edge gate ensures ETABS is only touched on explicit button press.
//   • Invalid Grasshopper points are skipped and reported.
//   • Optional uniqueNames can rename the created points when unique and valid.
//   • Model is unlocked before modifications via ComponentShared.EnsureModelUnlocked.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using ETABSv1;
using Grasshopper.Kernel;
using Rhino.Geometry;
using static MGT.ComponentShared;

namespace MGT
{
    public class GhcSetPoints : GH_Component
    {
        private bool _lastAdd;
        private string _lastMsg = "Idle.";
        private readonly List<string> _lastNames = new List<string>();

        public GhcSetPoints()
          : base(
                "ETABS Set Points",
                "SetPoints",
                "Create ETABS point objects from Grasshopper Point3d geometry with optional custom names.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                "MGT",
                "2.0 Point Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("8fbc63c9-5b8a-4f95-bc62-9d198f27f908");

        protected override Bitmap Icon
        {
            get { return null; }
        }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Press to run once (rising edge).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            p.AddPointParameter("points", "points", "Grasshopper points to create in ETABS.", GH_ParamAccess.list);
            p.AddTextParameter("uniqueNames", "uniqueNames", "Optional ETABS point names matched by index; leave blank to let ETABS decide.", GH_ParamAccess.list);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Status message.", GH_ParamAccess.item);
            p.AddTextParameter("etabsNames", "etabsNames", "ETABS point names created/updated.", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool add = false;
            cSapModel sapModel = null;
            List<Point3d> points = new List<Point3d>();
            List<string> uniqueNames = new List<string>();

            da.GetData(0, ref add);
            da.GetData(1, ref sapModel);
            da.GetDataList(2, points);
            da.GetDataList(3, uniqueNames);

            bool rising = !_lastAdd && add;
            if (!rising)
            {
                da.SetData(0, _lastMsg);
                da.SetDataList(1, _lastNames);
                _lastAdd = add;
                return;
            }

            string message;
            List<string> createdNames = new List<string>();

            try
            {
                if (sapModel == null)
                {
                    throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");
                }

                if (points == null || points.Count == 0)
                {
                    throw new InvalidOperationException("No input points provided.");
                }

                EnsureModelUnlocked(sapModel);

                int count = points.Count;
                string broadcastWarning = null;
                List<string> resolvedNames = ResolveNames(uniqueNames, count, out broadcastWarning);

                int added = 0;
                int renamed = 0;
                int skipped = 0;
                int failed = 0;
                int duplicateRequested = 0;
                int renameFailed = 0;

                HashSet<string> requestedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i < count; i++)
                {
                    Point3d ghPoint = points[i];
                    if (!ghPoint.IsValid)
                    {
                        skipped++;
                        createdNames.Add(string.Empty);
                        continue;
                    }

                    string etabsName = string.Empty;
                    int ret = sapModel.PointObj.AddCartesian(ghPoint.X, ghPoint.Y, ghPoint.Z, ref etabsName);
                    if (ret != 0 || string.IsNullOrWhiteSpace(etabsName))
                    {
                        failed++;
                        createdNames.Add(string.Empty);
                        continue;
                    }

                    added++;

                    string requested = resolvedNames[i];
                    if (!string.IsNullOrEmpty(requested))
                    {
                        if (requestedSet.Add(requested))
                        {
                            int renameRet = 0;
                            try
                            {
                                renameRet = sapModel.PointObj.ChangeName(etabsName, requested);
                            }
                            catch
                            {
                                renameRet = -1;
                            }

                            if (renameRet == 0)
                            {
                                etabsName = requested;
                                renamed++;
                            }
                            else
                            {
                                renameFailed++;
                            }
                        }
                        else
                        {
                            duplicateRequested++;
                        }
                    }

                    createdNames.Add(etabsName);
                }

                if (!string.IsNullOrEmpty(broadcastWarning))
                {
                    message = broadcastWarning;
                }
                else
                {
                    message = string.Empty;
                }

                string summary = $"Done: {added} point(s) created, {renamed} renamed, {skipped} skipped, {failed} failed.";

                if (duplicateRequested > 0)
                {
                    summary += $" {duplicateRequested} duplicate name request(s) ignored.";
                }

                if (renameFailed > 0)
                {
                    summary += $" {renameFailed} rename attempt(s) failed.";
                }

                message = string.IsNullOrEmpty(message) ? summary : message + " " + summary;
            }
            catch (Exception ex)
            {
                message = "Failed: " + ex.Message;
                createdNames.Clear();
            }

            this.Message = message;

            da.SetData(0, message);
            da.SetDataList(1, createdNames);

            _lastMsg = message;
            _lastNames.Clear();
            _lastNames.AddRange(createdNames);
            _lastAdd = add;
        }

        private static List<string> ResolveNames(IList<string> source, int count, out string warning)
        {
            warning = null;
            List<string> result = new List<string>(count);

            if (count <= 0)
            {
                return result;
            }

            if (source == null || source.Count == 0)
            {
                for (int i = 0; i < count; i++)
                {
                    result.Add(string.Empty);
                }
                return result;
            }

            int namesProvided = source.Count;

            if (namesProvided > count)
            {
                warning = "uniqueNames list longer than points list. Extra names ignored.";
            }

            int limit = Math.Min(namesProvided, count);
            for (int i = 0; i < limit; i++)
            {
                result.Add(NormalizeRequestedName(source[i]));
            }

            if (limit < count)
            {
                if (warning == null && namesProvided < count)
                {
                    warning = "uniqueNames list shorter than points list. Remaining points will use default ETABS names.";
                }

                for (int i = limit; i < count; i++)
                {
                    result.Add(string.Empty);
                }
            }

            return result;
        }

        private static string NormalizeRequestedName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string trimmed = value.Trim();
            if (trimmed.Equals("none", StringComparison.OrdinalIgnoreCase))
            {
                return string.Empty;
            }

            return trimmed;
        }
    }
}
