// -------------------------------------------------------------
// Component : Unit Scale Factor
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 8.0 (x64)
// Depends   : Grasshopper
// Panel     : "MGT" / "01. IO"
// Author    : Anh Bui
// -------------------------------------------------------------
//
// Behavior Notes:
//   + Converts between two length units (e.g. mm → m, in → ft, etc.)
//   + Returns multiplier to convert from srcUnit to dstUnit
//   + Case-insensitive; ignores leading/trailing spaces
//
// Inputs:
//   srcUnit (text, item)    Source unit (e.g. "mm", "m", "in", "ft", "µm")
//   dstUnit (text, item)    Destination unit (e.g. "m", "cm", "ft", etc.)
//
// Outputs:
//   scaleFactor (number, item)   Multiply value in srcUnit by this to get dstUnit
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using Grasshopper.Kernel;

namespace MGT
{
    public class GhcUnitScaleFactor : GH_Component
    {
        public GhcUnitScaleFactor()
          : base("Unit Scale Factor", "UnitScale",
                 "Compute scale factor to convert from one length unit to another.\nDeveloped by Mark Bui Quang Anh - Mark.Bui@meinhardtgroup.com",
                 "MGT", "01. IO")
        { }

        public override Guid ComponentGuid => new Guid("a3cfb5a4-0b45-4b28-bf5a-874c9467b01c");

        protected override Bitmap Icon
        {
            get
            {
                // 24x24 icon for Grasshopper panel
                Bitmap raw = Properties.Resources.unitConversionIcon;
                return new Bitmap(raw, new Size(24, 24));
            }
        }

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddTextParameter("srcUnit", "srcUnit", "Source unit (e.g. mm, m, cm, in, ft, µm).", GH_ParamAccess.item);
            p.AddTextParameter("dstUnit", "dstUnit", "Destination unit (e.g. m, cm, ft, etc.).", GH_ParamAccess.item);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddNumberParameter("scaleFactor", "scaleFactor", "Multiply value in srcUnit by this to convert to dstUnit.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess DA)
        {
            string src = string.Empty;
            string dst = string.Empty;
            DA.GetData(0, ref src);
            DA.GetData(1, ref dst);

            string su = Normalize(src);
            string du = Normalize(dst);

            // Conversion factors → meters
            var toMeters = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase)
            {
                { "m", 1.0 },
                { "cm", 0.01 },
                { "mm", 0.001 },
                { "µm", 1e-6 },
                { "um", 1e-6 },
                { "inch", 0.0254 },
                { "in", 0.0254 },
                { "ft", 0.3048 },
                { "feet", 0.3048 },
                { "yd", 0.9144 },
                { "mi", 1609.344 },
                { "km", 1000.0 }
            };

            double scale = double.NaN;
            try
            {
                if (toMeters.ContainsKey(su) && toMeters.ContainsKey(du))
                    scale = toMeters[su] / toMeters[du];
            }
            catch
            {
                scale = double.NaN;
            }

            DA.SetData(0, scale);
        }

        private string Normalize(string u)
        {
            if (string.IsNullOrWhiteSpace(u)) return string.Empty;
            return u.Trim().ToLower();
        }
    }
}
