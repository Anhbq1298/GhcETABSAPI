// -------------------------------------------------------------
// Component : ETABS Add Points (Add-only)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET 8.0 (x64)
// Depends   : RhinoCommon, Grasshopper, ETABSv1 (COM)
// Panel     : "MGT" / "2.0 Point Object Modelling"
// -------------------------------------------------------------
// Description:
//   Add ETABS point objects from Grasshopper Point3d. No move, no rename.
//   • If uniqueNames[i] is non-empty, request ETABS to use it (ETABS may alter).
//   • If uniqueNames[i] is empty, let ETABS auto-name.
//   • Coordinates multiplied by 'scale' (e.g., 1000 for mm→m).
//
// Inputs (ordered):
//   0) add         (bool, item)    Rising-edge trigger (exec on False→True).
//   1) sapModel    (ETABSv1.cSapModel, item)  From Attach component.
//   2) points      (Rhino.Geometry.Point3d, list)  Points to add.
//   3) uniqueNames (string, list)  Optional names by index; blanks let ETABS decide.
//   4) scale       (double, item)  Coord multiplier.
//
// Outputs:
//   0) msg         (string, item)  Summary/status.
//   1) etabsNames  (string, list)  Final ETABS names created this run.
//
// Behavior:
//   • Rising-edge + message replay.
//   • Skips invalid GH points.
//   • Ignores in-batch duplicate requested names (lets ETABS decide).
//   • Unlocks model before modifications.
//   • After success: refresh ETABS view + auto-expire GhcGetPointInfo (INLINE).
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
    public class GhcAddPoints : GH_Component
    {
        private bool _lastAdd;
        private string _lastMsg = "Idle.";
        private readonly List<string> _lastNames = new List<string>();

        public GhcAddPoints()
          : base(
                "ETABS Add Points",
                "AddPoints",
                "Add ETABS point objects from Grasshopper Point3d with optional custom names.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                "MGT",
                "2.0 Point Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("8fbc63c9-5b8a-4f95-bc62-9d198f27f908");
        protected override Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Press to run once (rising edge).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            p.AddPointParameter("points", "points", "Grasshopper points to add in ETABS.", GH_ParamAccess.list);
            int uniqueNamesIndex = p.AddTextParameter("uniqueNames", "uniqueNames", "Optional ETABS point names by index; blank lets ETABS auto-name.", GH_ParamAccess.list);
            p[uniqueNamesIndex].Optional = true;
            p.AddNumberParameter("scale", "scale", "Coordinate multiplier (e.g., 1000 for mm→m).", GH_ParamAccess.item, 1.0);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Status message.", GH_ParamAccess.item);
            p.AddTextParameter("etabsNames", "etabsNames", "ETABS point names created.", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            // Inputs
            bool add = false;
            cSapModel sapModel = null;
            var points = new List<Point3d>();
            var uniqueNames = new List<string>();
            double scale = 1.0;

            da.GetData(0, ref add);
            da.GetData(1, ref sapModel);
            da.GetDataList(2, points);
            da.GetDataList(3, uniqueNames);
            da.GetData(4, ref scale);

            if (IsInvalidNumber(scale) || scale <= 0.0) scale = 1.0;

            // Rising-edge gate
            bool rising = (!_lastAdd && add);
            if (!rising)
            {
                da.SetData(0, _lastMsg);
                da.SetDataList(1, _lastNames);
                _lastAdd = add;
                return;
            }

            string message;
            var createdNames = new List<string>();

            try
            {
                if (sapModel == null)
                    throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");
                if (points == null || points.Count == 0)
                    throw new InvalidOperationException("No input points provided.");

                EnsureModelUnlocked(sapModel);

                int count = points.Count;
                string broadcastWarning = null;
                var resolvedNames = ResolveNames(uniqueNames, count, out broadcastWarning);

                int added = 0;
                int customNamed = 0;
                int skipped = 0;
                int failed = 0;
                int duplicateRequested = 0;
                int namingAdjustedByEtabs = 0;

                var requestedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i < count; i++)
                {
                    var ghPoint = points[i];
                    if (!ghPoint.IsValid)
                    {
                        skipped++;
                        createdNames.Add(string.Empty);
                        continue;
                    }

                    // Prepare requested name; ignore duplicates within this batch
                    string requested = resolvedNames[i];
                    if (!string.IsNullOrEmpty(requested) && !requestedSet.Add(requested))
                    {
                        duplicateRequested++;
                        requested = string.Empty; // fall back to ETABS auto-name
                    }

                    string etabsName = string.IsNullOrEmpty(requested) ? string.Empty : requested;
                    int ret = sapModel.PointObj.AddCartesian(
                        ghPoint.X * scale,
                        ghPoint.Y * scale,
                        ghPoint.Z * scale,
                        ref etabsName);

                    if (ret != 0 || string.IsNullOrWhiteSpace(etabsName))
                    {
                        failed++;
                        createdNames.Add(string.Empty);
                        continue;
                    }

                    // Success
                    added++;
                    if (!string.IsNullOrEmpty(requested))
                    {
                        if (string.Equals(etabsName, requested, StringComparison.OrdinalIgnoreCase))
                            customNamed++;
                        else
                            namingAdjustedByEtabs++;
                    }

                    createdNames.Add(etabsName);
                }

                // Summary + warnings
                var warnings = new List<string>();
                if (!string.IsNullOrEmpty(broadcastWarning))
                    warnings.Add(broadcastWarning);
                if (duplicateRequested > 0)
                    warnings.Add($"{duplicateRequested} duplicate name request(s) ignored.");
                if (namingAdjustedByEtabs > 0)
                    warnings.Add($"{namingAdjustedByEtabs} requested name(s) changed by ETABS.");

                string summary = $"Done: {added} added, {customNamed} custom-named, {skipped} skipped, {failed} failed.";
                if (warnings.Count > 0) summary += " Warnings: " + string.Join(" | ", warnings);
                message = summary;

                // INLINE: refresh ETABS view + auto-update GhcGetPointInfo
                try { sapModel.View.RefreshView(0, false); } catch { /* non-fatal */ }
                try
                {
                    var doc = OnPingDocument();
                    if (doc != null)
                    {
                        doc.ScheduleSolution(1, d =>
                        {
                            foreach (var obj in d.Objects)
                            {
                                if (obj is GH_Component comp)
                                {
                                    bool match =
                                        string.Equals(comp.GetType().Name, "GhcGetPointInfo", StringComparison.OrdinalIgnoreCase) ||
                                        (comp.NickName?.IndexOf("GetPointInfo", StringComparison.OrdinalIgnoreCase) >= 0) ||
                                        (comp.Name?.IndexOf("Get Point Info", StringComparison.OrdinalIgnoreCase) >= 0);
                                    if (match)
                                    {
                                        try { comp.ExpireSolution(true); } catch { /* ignore */ }
                                    }
                                }
                            }
                        });
                    }
                }
                catch { /* non-fatal */ }
            }
            catch (Exception ex)
            {
                message = "Failed: " + ex.Message;
                createdNames.Clear();
            }

            // Outputs + sticky memory
            da.SetData(0, message);
            da.SetDataList(1, createdNames);

            _lastMsg = message;
            _lastNames.Clear();
            _lastNames.AddRange(createdNames);
            _lastAdd = add;
        }

        // ---- Utilities

        private static List<string> ResolveNames(IList<string> source, int count, out string warning)
        {
            warning = null;
            var result = new List<string>(count);
            if (count <= 0) return result;

            if (source == null || source.Count == 0)
            {
                for (int i = 0; i < count; i++) result.Add(string.Empty);
                return result;
            }

            int provided = source.Count;
            if (provided > count) warning = "uniqueNames list longer than points list. Extra names ignored.";

            int limit = Math.Min(provided, count);
            for (int i = 0; i < limit; i++) result.Add(NormalizeRequestedName(source[i]));

            if (limit < count)
            {
                if (warning == null && provided < count)
                    warning = "uniqueNames list shorter than points list. Remaining points will use default ETABS names.";
                for (int i = limit; i < count; i++) result.Add(string.Empty);
            }
            return result;
        }

        private static string NormalizeRequestedName(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string t = value.Trim();
            if (t.Equals("none", StringComparison.OrdinalIgnoreCase)) return string.Empty;
            return t;
        }
    }
}
