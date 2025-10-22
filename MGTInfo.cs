using System;
using System.Drawing;
using Grasshopper;
using Grasshopper.Kernel;

namespace MGT
{
    public class MGTInfo : GH_AssemblyInfo
    {
        public override string Name => "MGT";

        //Return a 24x24 pixel bitmap to represent this GHA library.
        public override Bitmap Icon => null;

        //Return a short string describing the purpose of this GHA library.
        public override string Description => "Meinhardt Grasshopper Tool";

        public override Guid Id => new Guid("ac55d3ed-5221-4c35-b9d5-762b9208e5d6");

        //Return a string identifying you or your company.
        public override string AuthorName => "Anh Bui";

        //Return a string representing your preferred contact details.
        public override string AuthorContact => "";

        //Return a string representing the version.  This returns the same version as the assembly.
        public override string AssemblyVersion => GetType().Assembly.GetName().Version.ToString();
    }
}