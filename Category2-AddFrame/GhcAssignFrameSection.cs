// -------------------------------------------------------------
// Component : ETABS Assign Frame Section (FrameObj.SetSection)
// Author    : Anh Bui (original pattern), extended by OpenAI assistant
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)
// Panel     : "MGT" / "2.0 Frame Object Modelling"
// -------------------------------------------------------------
//
// Inputs (ordered):
//   0) add         (bool, item)     Rising-edge trigger (False→True executes).
//   1) sapModel    (ETABSv1.cSapModel, item)  ETABS model from Attach component.
//   2) frameNames  (string, list)   ETABS frame object names to update.
//   3) sectionNames(string, list)   Section property names. Provide one item to broadcast
//                                   to all frames, or supply a list; if it runs short, the
//                                   last valid section is assumed for remaining frames.
//
// Outputs:
//   0) msg         (string, item)   Summary / status message.
//
// Behavior Notes:
//   + Rising-edge execution only (stores last trigger/message per component instance).
//   + sectionNames may contain 1 item (broadcast) or run shorter than frameNames; in the
//     latter case the last valid section provided is reused for remaining assignments.
//   + Blank / whitespace frame names are ignored (do not count toward total attempts).
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;
using ETABSv1;
using System.Windows.Forms;

namespace MGT
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
                "MGT",
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
                "Section property names. Provide 1 item to broadcast or a list; if it is shorter than frameNames the last valid section is reused.",
                GH_ParamAccess.list);
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

            if (!da.GetData(0, ref add)) add = false;
            if (!da.GetData(1, ref sapModel)) sapModel = null;
            da.GetDataList(2, frameNames);
            da.GetDataList(3, sectionNames);

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

                int totalAttempts = 0;
                int success = 0;
                int failures = 0;
                bool reusedSectionForRemaining = false;
                bool broadcast = sectionNames.Count == 1;
                string lastValidSection = null;

                for (int i = 0; i < frameNames.Count; i++)
                {
                    string frame = frameNames[i];
                    if (string.IsNullOrWhiteSpace(frame))
                    {
                        continue;
                    }

                    string section;
                    if (broadcast)
                    {
                        section = sectionNames[0];
                    }
                    else if (i < sectionNames.Count)
                    {
                        section = sectionNames[i];
                    }
                    else
                    {
                        section = lastValidSection ?? (sectionNames.Count > 0 ? sectionNames[sectionNames.Count - 1] : null);
                        if (!string.IsNullOrWhiteSpace(section))
                        {
                            reusedSectionForRemaining = true;
                        }
                    }

                    if (string.IsNullOrWhiteSpace(section))
                    {
                        continue;
                    }

                    section = section.Trim();
                    lastValidSection = section;

                    totalAttempts++;
                    int ret = sapModel.FrameObj.SetSection(frame.Trim(), section, eItemType.Objects);
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
                    string suffix = broadcast && !string.IsNullOrWhiteSpace(sectionNames[0])
                        ? $" (section '{sectionNames[0].Trim()}')"
                        : string.Empty;
                    _lastMsg = $"{success}/{totalAttempts} frame sections assigned{suffix}.";
                    if (failures > 0)
                    {
                        _lastMsg += $" {failures} failure(s).";
                    }
                    if (reusedSectionForRemaining)
                    {
                        _lastMsg += " Section list shorter than assignments; last valid section reused for remaining frames.";
                    }
                }
                // Refresh ETABS view
                try { sapModel.View.RefreshView(0, false); } catch { }

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