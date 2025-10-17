// -------------------------------------------------------------
// Component : ETABS Assign Frame Section (FrameObj.SetSection)
// Author    : Anh Bui (original pattern), extended by OpenAI assistant
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)
// Panel     : "ETABS API" / "2.0 Frame Object Modelling"
// -------------------------------------------------------------
//
// Inputs (ordered):
//   0) add         (bool, item)     Rising-edge trigger (False→True executes).
//   1) sapModel    (ETABSv1.cSapModel, item)  ETABS model from Attach component.
//   2) frameNames  (string, list)   ETABS frame object names to update.
//   3) sectionNames(string, list)   Section property names. Provide one item to broadcast
//                                   to all frames, or supply 1:1 list matching frameNames.
//   4) itemType    (int, item)      Optional ETABS eItemType value (default = Objects = 0).
//
// Outputs:
//   0) msg         (string, item)   Summary / status message.
//
// Behavior Notes:
//   + Rising-edge execution only (stores last trigger/message per component instance).
//   + Requires sectionNames to contain either 1 item (broadcast) or match frameNames count.
//   + Blank / whitespace frame names are ignored (do not count toward total attempts).
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;
using ETABSv1;
using System.Windows.Forms;

namespace GhcETABSAPI
{
    public class GhcAssignFrameSection : GH_Component
    {
        private bool _lastAdd = false;
        private string _lastMsg = "Idle";

        public GhcAssignFrameSection()
          : base(
                "Assign Frame Section",
                "FrSetSection",
                "Assign ETABS frame objects to section properties via FrameObj.SetSection.",
                "ETABS API",
                "2.0 Frame Object Modelling")
        { }

        public override Guid ComponentGuid => new Guid("0f6d2420-a1a1-4d65-8ed5-b0a49d94bc4a");

        protected override Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Rising-edge trigger; executes when this turns True.", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from your Attach component.", GH_ParamAccess.item);
            p.AddTextParameter(
                "frameNames",
                "frameNames",
                "ETABS frame object names to assign. Blank/whitespace names are ignored.",
                GH_ParamAccess.list);
            p.AddTextParameter(
                "sectionNames",
                "sectionNames",
                "Section property names. Provide 1 item to broadcast or 1:1 list matching frameNames.",
                GH_ParamAccess.list);
            p.AddIntegerParameter(
                "itemType",
                "itemType",
                "Optional ETABS eItemType value (0 = Objects, 1 = Group, etc.).",
                GH_ParamAccess.item,
                (int)eItemType.Objects);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Summary / status message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool add = false;
            cSapModel sapModel = null;
            List<string> frameNames = new List<string>();
            List<string> sectionNames = new List<string>();
            int rawItemType = (int)eItemType.Objects;

            if (!da.GetData(0, ref add)) add = false;
            if (!da.GetData(1, ref sapModel)) sapModel = null;
            da.GetDataList(2, frameNames);
            da.GetDataList(3, sectionNames);
            da.GetData(4, ref rawItemType);

            if (!(_lastAdd == false && add == true))
            {
                da.SetData(0, _lastMsg);
                _lastAdd = add;
                return;
            }

            string notification = string.Empty;

            try
            {
                if (sapModel == null) throw new Exception("sapModel is null.");
                if (frameNames == null || frameNames.Count == 0) throw new Exception("frameNames list is empty.");
                if (sectionNames == null || sectionNames.Count == 0) throw new Exception("sectionNames list is empty.");

                bool broadcast = sectionNames.Count == 1;
                if (!broadcast && sectionNames.Count != frameNames.Count)
                {
                    throw new Exception("sectionNames must have 1 item or match frameNames count.");
                }

                eItemType itemType = Enum.IsDefined(typeof(eItemType), rawItemType)
                    ? (eItemType)rawItemType
                    : eItemType.Objects;

                int totalAttempts = 0;
                int success = 0;
                int failures = 0;

                for (int i = 0; i < frameNames.Count; i++)
                {
                    string frame = frameNames[i];
                    if (string.IsNullOrWhiteSpace(frame))
                    {
                        continue;
                    }

                    string section = broadcast ? sectionNames[0] : sectionNames[i];
                    if (string.IsNullOrWhiteSpace(section))
                    {
                        continue;
                    }

                    totalAttempts++;
                    int ret = sapModel.FrameObj.SetSection(frame.Trim(), section.Trim(), itemType);
                    if (ret == 0)
                    {
                        success++;
                    }
                    else
                    {
                        failures++;
                    }
                }

                if (totalAttempts == 0)
                {
                    _lastMsg = "No valid frame/section pairs to assign.";
                }
                else
                {
                    string suffix = broadcast ? $" (section '{sectionNames[0].Trim()}')" : string.Empty;
                    _lastMsg = $"{success}/{totalAttempts} frame sections assigned{suffix}.";
                    if (failures > 0)
                    {
                        _lastMsg += $" {failures} failure(s).";
                    }
                }

                notification = _lastMsg;
            }
            catch (Exception ex)
            {
                _lastMsg = "Error: " + ex.Message;
                notification = _lastMsg;
            }

            da.SetData(0, _lastMsg);
            _lastAdd = add;

            if (!string.IsNullOrEmpty(notification))
            {
                try
                {
                    MessageBox.Show(notification, "ETABS Assign Frame Section", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch
                {
                    // Swallow any exception from UI notifications to avoid breaking SolveInstance.
                }
            }
        }
    }
}
