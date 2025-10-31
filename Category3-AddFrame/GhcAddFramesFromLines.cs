// -------------------------------------------------------------
// Component : Add Frames to ETABS
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : RhinoCommon, Grasshopper, ETABSv1 (COM)
// Panel     : Category = "ETABS", Subcategory = "Frame Object Modelling"
// Build     : x64; ETABSv1 reference -> Embed Interop Types = False
//
// INPUTS (ordered exactly as shown on the component):
//   0) add        (bool, item)    Rising-edge trigger. Executes only when it goes False→True.
//   1) sapModel   (ETABSv1.cSapModel, item)  ETABS model from your "Attach" component.
//   2) crvs       (Rhino.Geometry.Line, list) Lines to convert into ETABS frame objects.
//   3) propNames  (string, list)  Section property names. 1:1 with lines, or single value broadcast.
//   4) userNames  (string, list)  User-defined frame names. 1:1 with lines, or single value broadcast.
//   5) scale      (double, item)  Coordinate multiplier (e.g., 1000 for mm→m). If <= 0 → uses 1.0.
//
// OUTPUTS:
//   0) msg        (string, item)  Status message (replayed when not running).
//
// BEHAVIOR NOTES:
//   • Rising-edge gate: Only runs when 'add' transitions False→True. Otherwise replays last message.
//   • Per-instance memory: Uses private fields lastAdd/lastMsg that persist for this component
//     instance during the GH session. They reset if you delete the component or restart Rhino.
//   • Broadcasting: If a names list has 1 item, it is used for all lines; if empty → "".
//   • Zero-length/invalid lines: Skipped using EPS tolerance; they do not call ETABS.
//   • ETABS view refresh: Best-effort call at the end; failure is non-fatal.
//   • Explicit types: No 'var'. Uses Rhino.Geometry types via 'using Rhino.Geometry;' and ETABSv1.cSapModel.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Grasshopper.Kernel;
using Rhino.Geometry;
using ETABSv1;

namespace MGT
{
    public class GhcAddFramesToETABS : GH_Component
    {
        // ---- Config / constants ----
        private const double EPS = 1e-9;   // Zero-length tolerance for lines

        // ---- Per-instance memory (fields persist while the component exists) ----
        private bool lastAdd = false;      // remembers last trigger state
        private string lastMsg = "Idle.";  // remembers last status message

        public GhcAddFramesToETABS()
          : base(
                "Add Frames to ETABS",              // Display name
                "AddFramesToETABS",                 // Nickname
                "Create ETABS frame objects from Rhino.Geometry.Line with per-line property and optional user name.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                "MGT",                        // Category (tab)
                "3.0 Frame Object Modelling"        // Subcategory (panel)
            )
        { }

        // Unique ID for this component (generate once; keep stable).
        public override Guid ComponentGuid
        {
            get { return new Guid("64b6a6c3-6f1a-4d7b-bc9f-9c9a1b7d2f31"); }
        }

        // Replace with an embedded 24×24 PNG for a crisp toolbar icon (optional).
        protected override Bitmap Icon
        {
            get
            {
                try
                {
                    Bitmap raw = Properties.Resources.addFrameIcon;
                    return new Bitmap(raw, new Size(24, 24));
                }
                catch { return null; }
            }
        }

        // ---------------------------
        // Inputs / Outputs
        // ---------------------------
        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Press to run once (rising edge).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel (wire from your Attach component).", GH_ParamAccess.item);
            p.AddLineParameter("crvs", "crvs", "Input Rhino.Geometry.Line list.", GH_ParamAccess.list);
            p.AddTextParameter("propNames", "propNames", "Section property names (1:1 with lines, or single value broadcast to all).", GH_ParamAccess.list);
            p.AddTextParameter("userNames", "userNames", "User-defined frame names (1:1 with lines, or single value broadcast to all).", GH_ParamAccess.list);
            p.AddNumberParameter("scale", "scale", "Coordinate multiplier (e.g., 1000 for mm→m).", GH_ParamAccess.item, 1.0);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Status message.", GH_ParamAccess.item);
        }

        // ---------------------------
        // Solve
        // ---------------------------
        protected override void SolveInstance(IGH_DataAccess da)
        {
            // Inputs (explicit declarations)
            bool add = false;
            ETABSv1.cSapModel sapModel = null;
            List<Line> lines = new List<Line>();
            List<string> propNames = new List<string>();
            List<string> userNames = new List<string>();
            double scale = 1.0;

            da.GetData(0, ref add);
            da.GetData(1, ref sapModel);
            da.GetDataList(2, lines);
            da.GetDataList(3, propNames);
            da.GetDataList(4, userNames);
            da.GetData(5, ref scale);

            // Rising-edge gate
            bool rising = (!lastAdd) && add;

            if (!rising)
            {
                da.SetData(0, lastMsg);
                lastAdd = add;
                return;
            }

            // ---- Running branch ----
            string msg;

            // Basic guards
            if (sapModel == null)
            {
                msg = "sapModel is null. Wire it from your Attach component.";
                Finish(da, add, msg);
                return;
            }
            if (lines == null || lines.Count == 0)
            {
                msg = "No input lines.";
                Finish(da, add, msg);
                return;
            }
            if (scale <= 0.0) scale = 1.0;

            int added = 0;
            int skipped = 0;
            int failed = 0;
            int lastRet = 0;

            int n = lines.Count;
            for (int i = 0; i < n; i++)
            {
                Line ln = lines[i];

                // Skip invalid / zero-length
                if (!ln.IsValid || ln.Length <= EPS)
                {
                    skipped++;
                    continue;
                }

                // Broadcast property
                string prop = (propNames != null && propNames.Count > 0)
                    ? propNames[Math.Min(i, propNames.Count - 1)] ?? string.Empty
                    : string.Empty;

                // Broadcast user name
                string uname = (userNames != null && userNames.Count > 0)
                    ? userNames[Math.Min(i, userNames.Count - 1)] ?? string.Empty
                    : string.Empty;

                // Scale coordinates
                Point3d a = ln.From;
                Point3d b = ln.To;

                string newName = string.Empty;
                int ret;

                try
                {
                    ret = sapModel.FrameObj.AddByCoord(
                        a.X * scale, a.Y * scale, a.Z * scale,
                        b.X * scale, b.Y * scale, b.Z * scale,
                        ref newName, prop, uname);
                }
                catch
                {
                    ret = -1;
                }

                if (ret == 0) added++;
                else { failed++; lastRet = ret; }
            }

            // Refresh ETABS view
            try { sapModel.View.RefreshView(0, false); } catch { }

            // Compose message
            msg = BuildFinalMessage(added, skipped, failed, lastRet);
            Finish(da, add, msg);

            // ---- MessageBox feedback ----
            this.Message = msg;
            MessageBoxIcon icon = (added > 0) ? MessageBoxIcon.Information : MessageBoxIcon.Error;
            MessageBox.Show(msg, "ETABS Frame Transfer", MessageBoxButtons.OK, icon);
        }

        // Build final status
        private static string BuildFinalMessage(int added, int skipped, int failed, int lastRet)
        {
            if (added > 0)
                return $"Done: {added} frame(s) transferred.";

            List<string> parts = new List<string>();
            if (skipped > 0) parts.Add($"{skipped} invalid/zero-length line(s)");
            if (failed > 0) parts.Add($"{failed} API failure(s) (last ret={lastRet})");
            return (parts.Count > 0) ? "Failed: " + string.Join(", ", parts) : "Failed.";
        }

        // Output + update persistent fields
        private void Finish(IGH_DataAccess da, bool add, string msg)
        {
            da.SetData(0, msg);
            lastMsg = msg;
            lastAdd = add;
            TriggerGetAllFrameInfoRefresh();
        }

        private void TriggerGetAllFrameInfoRefresh()
        {
            try
            {
                GH_Document document = OnPingDocument();
                if (document == null)
                {
                    return;
                }

                List<GhcGetAllFrameInfo> targets = new List<GhcGetAllFrameInfo>();
                foreach (IGH_DocumentObject obj in document.Objects)
                {
                    if (obj is GhcGetAllFrameInfo target &&
                        ReferenceEquals(target.OnPingDocument(), document) &&
                        !target.Locked &&
                        !target.Hidden)
                    {
                        targets.Add(target);
                    }
                }

                if (targets.Count == 0)
                {
                    return;
                }

                document.ScheduleSolution(5, _ =>
                {
                    foreach (GhcGetAllFrameInfo target in targets)
                    {
                        if (!target.Locked && !target.Hidden)
                        {
                            target.ExpireSolution(false);
                        }
                    }
                });
            }
            catch
            {
                // Swallow exceptions to avoid interrupting the main solve.
            }
        }
    }
}
