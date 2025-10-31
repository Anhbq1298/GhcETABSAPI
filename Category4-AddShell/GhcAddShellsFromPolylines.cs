// -------------------------------------------------------------
// Component : ETABS Add Shells from Polylines
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 8.0 (x64)
// Depends   : RhinoCommon, Grasshopper, ETABSv1 (COM)
// Panel     : "ETABS" / "Area Object Modelling"
// Author    : Anh Bui
// -------------------------------------------------------------
//
// Inputs:
//   add        (bool, item)      Rising-edge trigger; True only on button click
//   sapModel   (object, item)    ETABS cSapModel COM object
//   Bnds       (list[curve])     Closed boundary polylines or curves
//   propNames  (list[text])      ETABS shell/area section property names (broadcasted if shorter)
//   userNames  (list[text])      Optional per-area user names (broadcasted if shorter)
//   scale      (float, item)     Rhino → ETABS scale factor (default = 1.0)
//
// Outputs:
//   msg        (text, item)      Status message ("Done: X slab(s) transferred.")
//
// -------------------------------------------------------------
// Behavior Notes:
//   + Rising-edge trigger via "add" input.
//   + Converts closed Rhino curves to vertex arrays.
//   + Each curve → one ETABS shell (AreaObj.AddByCoord).
//   + Scales XYZ by user "scale" (Rhino→ETABS).
//   + Property and username lists auto-broadcasted.
//   + Ends with View.RefreshView(0, false).
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Grasshopper.Kernel;
using Rhino.Geometry;
using ETABSv1;

namespace MGT
{
    public class GhcAddShellsFromPolylines : GH_Component
    {
        private bool lastAdd = false;
        private string lastMsg = "Idle.";

        public GhcAddShellsFromPolylines()
          : base("ETABS Add Shells", "AddShells",
                 "Create ETABS shell areas from closed Rhino polylines.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                 "MGT",                            // Category (tab)
                 "4.0 Area Object Modelling"            // Subcategory (panel)
                 )
        { }

        public override Guid ComponentGuid => new Guid("7cf9db55-7c04-4c9c-bd12-6e5c2e073c8b");

        protected override Bitmap Icon
        {
            get
            {
                Bitmap raw = Properties.Resources.addShellIcon;
                Bitmap resized = new Bitmap(raw, new Size(24, 24));
                return resized;
            }
        }

        // ---------------------------------------------------------
        // Inputs / Outputs
        // ---------------------------------------------------------
        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Trigger execution (True only on click).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel object.", GH_ParamAccess.item);
            p.AddCurveParameter("Bnds", "Bnds", "Closed boundary curves (list or tree).", GH_ParamAccess.list);
            p.AddTextParameter("propNames", "propNames", "Area section property name(s).", GH_ParamAccess.list);
            p.AddTextParameter("userNames", "userNames", "Optional per-area user name(s).", GH_ParamAccess.list);
            p.AddNumberParameter("scale", "scale", "Rhino→ETABS scale factor.", GH_ParamAccess.item, 1.0);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Status message.", GH_ParamAccess.item);
        }

        // ---------------------------------------------------------
        // Main SolveInstance
        // ---------------------------------------------------------
        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool add = false;
            cSapModel sapModel = null;
            List<Curve> Bnds = new List<Curve>();
            List<string> propNames = new List<string>();
            List<string> userNames = new List<string>();
            double scale = 1.0;

            if (!DA.GetData(0, ref add)) return;
            if (!DA.GetData(1, ref sapModel)) return;
            if (!DA.GetDataList(2, Bnds)) return;
            DA.GetDataList(3, propNames);
            DA.GetDataList(4, userNames);
            DA.GetData(5, ref scale);

            bool risingEdge = (!lastAdd && add);

            if (!risingEdge)
            {
                this.Message = lastMsg;
                DA.SetData(0, lastMsg);
                lastAdd = add;
                return;
            }

            string msg;

            if (sapModel == null)
            {
                msg = "sapModel is null.";
                this.Message = msg;
                DA.SetData(0, msg);
                lastMsg = msg;
                lastAdd = add;
                return;
            }

            List<List<Point3d>> polys = CurvesToPolys(Bnds);
            if (polys.Count == 0)
            {
                msg = "No valid closed polylines.";
                this.Message = msg;
                MessageBox.Show(msg, "ETABS Shell Transfer", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DA.SetData(0, msg);
                lastMsg = msg;
                lastAdd = add;
                return;
            }

            int n = CreateEtabsAreas(sapModel, polys, propNames, userNames, scale);
            msg = (n > 0) ? $"Done: {n} slab(s) transferred." : "Failed.";
            this.Message = msg;
            MessageBoxIcon icon = (n > 0) ? MessageBoxIcon.Information : MessageBoxIcon.Error;
            MessageBox.Show(msg, "ETABS Shell Transfer", MessageBoxButtons.OK, icon);

            DA.SetData(0, msg);

            lastMsg = msg;
            lastAdd = add;
        }

        // ---------------------------------------------------------
        // Helper: Convert Rhino curves to vertex arrays
        // ---------------------------------------------------------
        private List<List<Point3d>> CurvesToPolys(List<Curve> Bnds)
        {
            List<List<Point3d>> polys = new List<List<Point3d>>();
            double tol = 1e-9;

            foreach (Curve c in Bnds)
            {
                if (c == null || !c.IsClosed) continue;

                Polyline pl;
                if (c is PolylineCurve plc)
                {
                    pl = plc.ToPolyline();
                }
                else if (!c.TryGetPolyline(out pl))
                {
                    continue;
                }

                List<Point3d> pts = pl.ToList();
                if (pts.Count < 3) continue;
                if (pts.First().DistanceTo(pts.Last()) <= tol)
                    pts.RemoveAt(pts.Count - 1);

                polys.Add(pts);
            }

            return polys;
        }

        // ---------------------------------------------------------
        // Helper: Normalize list size
        // ---------------------------------------------------------
        private List<string> BroadcastList(List<string> input, int n, string defaultVal)
        {
            List<string> list = new List<string>();
            if (input == null || input.Count == 0)
            {
                for (int i = 0; i < n; i++) list.Add(defaultVal);
                return list;
            }

            for (int i = 0; i < n; i++)
            {
                string val = (i < input.Count) ? input[i] : input[input.Count - 1];
                if (string.IsNullOrEmpty(val)) val = defaultVal;
                list.Add(val);
            }

            return list;
        }

        // ---------------------------------------------------------
        // Helper: ETABS API call
        // ---------------------------------------------------------
        private int CreateEtabsAreas(cSapModel sapModel, List<List<Point3d>> polys,
                                     List<string> propNames, List<string> userNames,
                                     double scale)
        {
            if (polys.Count == 0) return 0;

            List<string> propList = BroadcastList(propNames, polys.Count, "Default");
            List<string> userList = BroadcastList(userNames, polys.Count, "");

            int created = 0;

            for (int i = 0; i < polys.Count; i++)
            {
                List<Point3d> pts = polys[i];
                int nPts = pts.Count;
                double[] x = new double[nPts];
                double[] y = new double[nPts];
                double[] z = new double[nPts];

                for (int j = 0; j < nPts; j++)
                {
                    x[j] = pts[j].X * scale;
                    y[j] = pts[j].Y * scale;
                    z[j] = pts[j].Z * scale;
                }

                string nameOut = "";
                try
                {
                    sapModel.AreaObj.AddByCoord(nPts, ref x, ref y, ref z,
                                                ref nameOut, propList[i],
                                                userList[i], "Global");
                    created++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("AddByCoord failed: " + ex.Message);
                }
            }

            try
            {
                sapModel.View.RefreshView(0, false);
            }
            catch { }

            return created;
        }
    }
}
