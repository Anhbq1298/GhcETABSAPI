// -------------------------------------------------------------
// Component : Get Rhino Units (minimal)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper (.NET Framework 4.8 x64)
// Depends   : RhinoCommon, Grasshopper
// Panel     : "Rhino" / "Document"
//
// Behavior Notes:
//   + No inputs; reads RhinoDoc.ActiveDoc each solve.
//   + If no active document → outputs ("None", "N/A").
//
// Inputs:
//   + None.
//
// Outputs:
//   + unitName   (text, item)  Friendly name, e.g., "Millimeters".
//   + lengthUnit (text, item)  Abbreviation, e.g., "mm".
// -------------------------------------------------------------

using System;
using System.Drawing;
using Grasshopper.Kernel;
using Rhino;

namespace MGT
{
    public class GhcGetRhinoUnits : GH_Component
    {
        public GhcGetRhinoUnits()
          : base("Get Rhino Units", // Display name
                "RhUnits", // Nickname
                 "Return the current Rhino model unit system (name and abbreviation).\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                 "MGT", // Category (tab)
                 "01. IO" // Subcategory (panel)
                )
        { }

        // New GUID for this component/version
        public override Guid ComponentGuid => new Guid("9f13a3ea-2a37-4f7b-a4f9-1f9020a68e92");

        protected override Bitmap Icon
        {
            get
            {
                try
                {
                    Bitmap raw = Properties.Resources.getRhinoUnitIcon;
                    return new Bitmap(raw, new Size(24, 24));
                }
                catch { return null; }
            }
        }
        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            // No inputs
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("unitName", "Name", "Friendly name of the Rhino model unit system.", GH_ParamAccess.item);
            p.AddTextParameter("lengthUnit", "U", "Abbreviation of the Rhino model unit system.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            RhinoDoc doc = RhinoDoc.ActiveDoc;
            if (doc == null)
            {
                da.SetData(0, "None");
                da.SetData(1, "N/A");
                return;
            }

            Rhino.UnitSystem us = doc.ModelUnitSystem;
            string unitName = us.ToString();      // e.g. "Millimeters"
            string abbrev = Abbrev(us);         // e.g. "mm"

            da.SetData(0, unitName);
            da.SetData(1, abbrev);
        }

        private static string Abbrev(Rhino.UnitSystem u)
        {
            switch (u)
            {
                case Rhino.UnitSystem.Angstroms: return "Å";
                case Rhino.UnitSystem.Nanometers: return "nm";
                case Rhino.UnitSystem.Microns: return "µm";
                case Rhino.UnitSystem.Millimeters: return "mm";
                case Rhino.UnitSystem.Centimeters: return "cm";
                case Rhino.UnitSystem.Decimeters: return "dm";
                case Rhino.UnitSystem.Meters: return "m";
                case Rhino.UnitSystem.Dekameters: return "dam";
                case Rhino.UnitSystem.Hectometers: return "hm";
                case Rhino.UnitSystem.Kilometers: return "km";
                case Rhino.UnitSystem.Megameters: return "Mm";
                case Rhino.UnitSystem.Gigameters: return "Gm";
                case Rhino.UnitSystem.Inches: return "in";
                case Rhino.UnitSystem.Feet: return "ft";
                case Rhino.UnitSystem.Yards: return "yd";
                case Rhino.UnitSystem.Miles: return "mi";
                case Rhino.UnitSystem.NauticalMiles: return "nmi";
                case Rhino.UnitSystem.None: return "None";
                case Rhino.UnitSystem.Unset: return "Unset";
                default: return u.ToString();
            }
        }
    }
}
