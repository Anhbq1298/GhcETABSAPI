// -------------------------------------------------------------
// Component : ETABS Attach (minimal)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper (.NET Framework 4.8 x64)
// Depends   : Grasshopper, ETABSv1 (COM), System.Windows.Forms
// Panel     : "1.0" / "IO"
//
// Behavior Notes:
//   + Requires ETABS already running (no auto-start).
//   + When run = False ? outputs are (null, null, "Idle").
//   + On success ? shows a MessageBox and returns cOAPI & cSapModel.
//   + On failure ? shows a MessageBox and returns null outputs with error text.
//
// Inputs:
//   + run (bool, item)              Pulse True to attach; False leaves outputs idle.
//
// Outputs:
//   + ETABSObject (generic, item)   ETABS cOAPI instance or null on idle/error.
//   + SapModel    (generic, item)   ETABS cSapModel or null on idle/error.
//   + message     (text, item)      Status string describing the result.
// -------------------------------------------------------------

using ETABSv1;
using Grasshopper.Kernel;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace MGT
{
    public class GhcAttachETABSInstance : GH_Component
    {
        public GhcAttachETABSInstance()
          : base("ETABS Attach", "ETABSAttach",
                 "Attach to a running ETABS instance and return cOAPI & cSapModel.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                 "MGT", // Category (tab)
                 "01. IO" // Subcategory (panel)
                )
        { }

        // New GUID
        public override Guid ComponentGuid => new Guid("7f2b1a86-3c6a-44f1-9f19-4d7b1bd613a5");

        // Icon: embedded resource named "etabsIcon" resized to 24x24
        protected override System.Drawing.Bitmap Icon
        {
            get
            {
                try
                {
                    Bitmap raw = Properties.Resources.etabsIcon;
                    return new Bitmap(raw, new Size(24, 24));
                }
                catch { return null; }
            }
        }

        // Input params
        protected override void RegisterInputParams(GH_InputParamManager p)
          => p.AddBooleanParameter("run", "run", "Pulse to attach", GH_ParamAccess.item, false);

        // Output params
        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddGenericParameter("ETABSObject", "ETABSObject", "ETABS cOAPI instance", GH_ParamAccess.item);
            p.AddGenericParameter("SapModel", "SapModel", "ETABS cSapModel object", GH_ParamAccess.item);
            p.AddTextParameter("message", "msg", "Status message", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            da.GetData(0, ref run);

            ETABSv1.cOAPI ETABSObject = null;
            ETABSv1.cSapModel SapModel = null;
            string msg = "Idle";

            if (!run)
            {
                da.SetData(0, ETABSObject);
                da.SetData(1, SapModel);
                da.SetData(2, msg);
                return;
            }

            try
            {
                // Ensure ETABSv1.dll can be found before casting
                ETABSv1.cHelper helper = new ETABSv1.Helper();

                // Attach to running ETABS
                ETABSObject = helper.GetObject("CSI.ETABS.API.ETABSObject");
                if (ETABSObject == null)
                    throw new Exception("ETABS cOAPI not found. Ensure ETABS is running.");

                SapModel = ETABSObject.SapModel
                    ?? throw new Exception("ETABS cSapModel not available from cOAPI.");

                string fileName = "";
                try { fileName = SapModel.GetModelFilename(); } catch { }

                msg = string.IsNullOrEmpty(fileName) ? "Attached to ETABS." : "Attached. Model: " + fileName;
                MessageBox.Show(msg, "ETABS Attach", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                msg = "Attach failed: " + ex.Message;
                MessageBox.Show(msg, "ETABS Attach Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ETABSObject = null;
                SapModel = null;
            }

            da.SetData(0, ETABSObject);
            da.SetData(1, SapModel);
            da.SetData(2, msg);
        }
    }
}
