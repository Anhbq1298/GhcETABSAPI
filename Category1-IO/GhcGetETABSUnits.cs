// -------------------------------------------------------------
// Component: GhcGetETABSUnits
// Purpose  : Return current ETABS units (force/length/temperature)
// Target   : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends  : Grasshopper, ETABSv1 (COM)  [Embed Interop Types = False]
// Panel    : "ETABS" / "IO"
// Author   : Anh Bui
// -------------------------------------------------------------

using System;
using System.Drawing;
using Grasshopper.Kernel;
using ETABSv1;

namespace GhcETABSAPI
{
    public class GhcGetETABSUnits : GH_Component
    {
        public GhcGetETABSUnits()
          : base("Get ETABS Units", 
                "ETABSUnits",
                 "Get current ETABS working units (force/length/temperature)",
                 "ETABS API", // Category (tab)
                 "01. IO" // Subcategory (panel)
                )
        { }
        private string testGit = "1";
        public override Guid ComponentGuid => new Guid("a6f0d0b2-6f3f-4e6c-9db0-9f9f2b0a6e21");

        protected override Bitmap Icon
        {
            get
            {
                try
                {
                    Bitmap raw = Properties.Resources.getEtabsUnitsIcon;
                    return new Bitmap(raw, new Size(24, 24));
                }
                catch { return null; }
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel.", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            // Output order matches your screenshot
            p.AddTextParameter("forceUnit", "forceUnit", "Force unit abbreviation, e.g. kN.", GH_ParamAccess.item);
            p.AddTextParameter("lengthUnit", "lengthUnit", "Length unit abbreviation, e.g. m.", GH_ParamAccess.item);
            p.AddTextParameter("temperatureUnit", "temperatureUnit", "Temperature unit abbreviation, e.g. C.", GH_ParamAccess.item);
            p.AddTextParameter("unitName", "unitName", "Combined unit string, e.g. kN-m-C.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string forceUnit = "";
            string lengthUnit = "";
            string temperatureUnit = "";
            string unitName = "";

            ETABSv1.cSapModel sapModel = null;           // (cSapModel, item)

            if (!DA.GetData(0, ref sapModel) || sapModel == null)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, "No ETABS SapModel provided.");
                DA.SetData(0, forceUnit);
                DA.SetData(1, lengthUnit);
                DA.SetData(2, temperatureUnit);
                DA.SetData(3, unitName);
                return;
            }


            try
            {
                eForce f = 0;
                eLength L = 0;
                eTemperature T = 0;

                int ret = sapModel.GetPresentUnits_2(ref f, ref L, ref T);
                if (ret != 0)
                    AddRuntimeMessage(GH_RuntimeMessageLevel.Warning, $"GetPresentUnits_2 returned code {ret}.");

                forceUnit = MapForce((int)f);
                lengthUnit = MapLength((int)L);
                temperatureUnit = MapTemp((int)T);
                unitName = $"{forceUnit}-{lengthUnit}-{temperatureUnit}";
            }
            catch (Exception ex)
            {
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, $"Failed to query units: {ex.Message}");
            }

            DA.SetData(0, forceUnit);
            DA.SetData(1, lengthUnit);
            DA.SetData(2, temperatureUnit);
            DA.SetData(3, unitName);
        }

        private static string MapForce(int code)
        {
            switch (code)
            {
                case 0: return "N/A";
                case 1: return "lb";
                case 2: return "kip";
                case 3: return "N";
                case 4: return "kN";
                case 5: return "kgf";
                case 6: return "tonf";
                default: return $"?({code})";
            }
        }

        private static string MapLength(int code)
        {
            switch (code)
            {
                case 0: return "N/A";
                case 1: return "in";
                case 2: return "ft";
                case 3: return "µm";
                case 4: return "mm";
                case 5: return "cm";
                case 6: return "m";
                default: return $"?({code})";
            }
        }

        private static string MapTemp(int code)
        {
            switch (code)
            {
                case 0: return "N/A";
                case 1: return "F";
                case 2: return "C";
                default: return $"?({code})";
            }
        }
    }
}
