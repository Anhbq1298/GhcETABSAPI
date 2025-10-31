// -------------------------------------------------------------
// Component : ETABS Add Frames to Group (SetGroup_1, minimal)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)  [Embed Interop Types = False]
// Panel     : "MGT" / "2.0 Frame Object Modelling"
// -------------------------------------------------------------
//
// Inputs (ordered):
//   0) add        (bool, item)      Rising-edge trigger
//   1) sapModel   (ETABSv1.cSapModel, item)  ETABS model
//   2) groupName  (string, item)    Target group name (auto-defined via SetGroup_1)
//   3) frameNames (string, list)    ETABS frame object names to assign
//                                   NOTE: This list MUST already be unique upstream.
//                                         The component does NOT de-duplicate; blank/whitespace names are ignored.
//
// Outputs:
//   0) msg        (string, item)    "X/Y frames assigned to 'GroupName'." (Y = input list count)
//
// Behavior Notes:
//   + Defines/updates the group unconditionally via SetGroup_1(groupName).
//   + Assigns each frame via SetGroupAssign(frameName, groupName, false, eItemType.Objects).
//   + Success counted by ETABS return code (0 = success). No model scans or group listing.
//   + Rising-edge execution only (False→True on 'add'); replays last message otherwise.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;
using ETABSv1;

namespace MGT
{
    public class GhcAddFramesToGroup : GH_Component
    {
        private bool _lastAdd = false;
        private string _lastMsg = "Idle";

        public GhcAddFramesToGroup()
          : base(
                "Add Frames to Group",
                "FrToGroup",
                "Add ETABS frame objects to a group (auto-define with SetGroup_1, minimal).\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                "MGT",
                "2.0 Frame Object Modelling")
        { }

        // New GUID for this version
        public override Guid ComponentGuid => new Guid("C2C6A7B5-6A49-4E28-93C3-8D39E3C6B4E7");

        protected override Bitmap Icon
        {
            get
            {
                try
                {
                    Bitmap raw = Properties.Resources.addFramesToGroupIcon;
                    return new Bitmap(raw, new Size(24, 24));
                }
                catch { return null; }
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Rising-edge trigger; executes when this turns True.", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from your Attach component.", GH_ParamAccess.item);
            p.AddTextParameter("groupName", "groupName", "Target ETABS group name. Created/updated via SetGroup_1.", GH_ParamAccess.item, "GH_Frames");            
            p.AddTextParameter(
                "frameNames",
                "frameNames",
                "ETABS frame object names to assign. IMPORTANT: This list must already be unique upstream; the component does not de-duplicate. Blank/whitespace items are ignored.",
                GH_ParamAccess.list
            );
            p.AddBooleanParameter("removeMode", "replace", "True = remove existing FRAME members of the target group before adding.", GH_ParamAccess.item, false);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Summary: \"X/Y frames assigned to 'groupName'.\"", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            bool add = false;
            ETABSv1.cSapModel sapModel = null;
            string groupName = null;
            List<string> frameNames = new List<string>();
            bool removeMode = false;

            if (!DA.GetData(0, ref add)) add = false;
            if (!DA.GetData(1, ref sapModel)) sapModel = null;
            if (!DA.GetData(2, ref groupName)) groupName = null;
            DA.GetDataList(3, frameNames);
            if (!DA.GetData(4, ref removeMode)) removeMode = false;
            // Rising-edge gate (False→True)
            if (!(_lastAdd == false && add == true))
            {
                DA.SetData(0, _lastMsg);
                _lastAdd = add;
                return;
            }

            string targetGroup = (groupName ?? string.Empty).Trim();
            int total = frameNames.Count;   // use the list as provided (assumed unique)
            string frameStatus = removeMode ? "removed" : "assigned";
            int success = 0;

            try
            {
                if (sapModel == null) throw new Exception("sapModel is null.");
                if (string.IsNullOrWhiteSpace(targetGroup)) throw new Exception("groupName is empty.");

                
                // Define/update group (creates if missing)
                int gr = sapModel.GroupDef.SetGroup_1(targetGroup);

                if (gr != 0) throw new Exception("SetGroup_1 failed.");

                // Assign frames (skip blanks; no de-duplication or existence pre-checks)
                for (int i = 0; i < frameNames.Count; i++)
                {
                    string nm = frameNames[i];
                    if (string.IsNullOrWhiteSpace(nm)) continue;
                    int ret = sapModel.FrameObj.SetGroupAssign(nm.Trim(), targetGroup, removeMode, eItemType.Objects);
                    if (ret == 0) success++;
                }

                _lastMsg = $"{success}/{total} frames '{frameStatus}' to '{targetGroup}'.";
            }
            catch (Exception ex)
            {
                _lastMsg = "Error: " + ex.Message;
            }

            DA.SetData(0, _lastMsg);
            _lastAdd = add;
        }
    }
}
