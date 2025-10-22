// -------------------------------------------------------------
// Component : Write GH Tree To Excel
// Author    : Anh Bui (extended by ChatGPT)
// Encoding  : UTF-8
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : RhinoCommon, Grasshopper, Microsoft.Office.Interop.Excel
// Panel     : Category = "ETABS API", Subcategory = "0.0 Utility"
// Build     : x64; Excel interop reference -> Embed Interop Types = False
//
// INPUTS (ordered exactly as shown on the component):
//   0) add            (bool, item)  Rising-edge trigger. Runs when it goes False→True.
//   1) tree           (generic, tree) Data tree to export. Each GH_Path becomes an Excel column.
//   2) path           (string, item) Workbook path (relative → plugin directory). Default "TreeExport.xlsx".
//   3) ws             (string, item) Worksheet name. Created if missing. Default "Sheet1".
//   4) address        (string, item) Starting Excel address (e.g., "A1"). Defaults to "A1".
//   5) excelOptions   (bool, list)  Optional list: [0]=visible (default true), [1]=saveAfterWrite (default true),
//                                 [2]=readOnly (default false).
//   6) headers        (string, list) Optional column headers; blanks auto-fill ("Header", "Header_1", ...).
//
// OUTPUTS:
//   0) msg            (string, item) Status message (replayed while idle).
//
// BEHAVIOR NOTES:
//   • Rising-edge gate identical to other ETABS utility components (per-instance memory).
//   • Uses ExcelHelpers to attach to running Excel or create a new instance.
//   • Each column header = supplied header (blank → auto) OR branch path string when headers input omitted.
//   • Saves workbook only when saveAfterWrite = true AND readOnly = false.
//   • COM cleanup handled inside ExcelHelpers.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Text.RegularExpressions;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using Excel = Microsoft.Office.Interop.Excel;

namespace MGT
{
    public class GhcWriteGHTreeToExcel : GH_Component
    {
        private bool lastAdd = false;
        private string lastMsg = "Idle.";

        private const string DefaultHeaderBase = "Header";

        public GhcWriteGHTreeToExcel()
          : base(
                "Write Tree To Excel",
                "WriteTreeExcel",
                "Write a Grasshopper data tree to an Excel worksheet. Each path becomes a column and the branch values populate rows.",
                "ETABS API",
                "0.0 Utility")
        { }

        public override Guid ComponentGuid => new Guid("8f222ec4-2d36-4b3c-9a36-3a642c224f2f");

        protected override Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("add", "add", "Press to run once (rising edge).", GH_ParamAccess.item, false);
            p.AddGenericParameter("tree", "tree", "Data tree to export. Each branch becomes an Excel column.", GH_ParamAccess.tree);

            // path: null/empty => dùng template; nếu cung cấp thì PHẢI là absolute .xlsx
            int pathIndex = p.AddTextParameter(
                "path", "path",
                "Workbook path (.xlsx). Leave blank/null to auto-use the template; if provided it MUST be an absolute path (e.g., C:\\...\\file.xlsx). Relative paths are not allowed.",
                GH_ParamAccess.item, string.Empty);
            p[pathIndex].Optional = true;

            p.AddTextParameter("ws", "ws", "Worksheet name. Created if missing.", GH_ParamAccess.item, "Sheet1");
            p.AddTextParameter("address", "address", "Starting Excel cell address (e.g., A1).", GH_ParamAccess.item, "A1");

            int optionsIndex = p.AddBooleanParameter(
                "excelOptions", "excelOpt",
                "Optional toggles: [0]=visible (default true), [1]=saveAfterWrite (default true), [2]=readOnly (default false).",
                GH_ParamAccess.list);
            p[optionsIndex].Optional = true;

            int headerIndex = p.AddTextParameter(
                "headers", "headers",
                "Optional header labels. If empty, branch paths are used.",
                GH_ParamAccess.list);
            p[headerIndex].Optional = true;
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("msg", "msg", "Status message.", GH_ParamAccess.item);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool add = false;
            da.GetData(0, ref add);

            bool rising = (!lastAdd) && add;
            if (!rising)
            {
                da.SetData(0, lastMsg);
                lastAdd = add;
                return;
            }

            GH_Structure<IGH_Goo> tree = null;
            string workbookPath = null;
            string worksheetName = "Sheet1";
            string address = "A1";
            bool visible = true;
            bool saveAfterWrite = true;
            bool readOnly = false;
            var excelOptions = new List<bool>();
            var headerOverrides = new List<string>();
            string message;

            da.GetDataTree(1, out tree);
            da.GetData(2, ref workbookPath);
            da.GetData(3, ref worksheetName);
            da.GetData(4, ref address);
            da.GetDataList(5, excelOptions);
            da.GetDataList(6, headerOverrides);

            if (excelOptions != null)
            {
                if (excelOptions.Count > 0)
                    visible = excelOptions[0];
                if (excelOptions.Count > 1)
                    saveAfterWrite = excelOptions[1];
                if (excelOptions.Count > 2)
                    readOnly = excelOptions[2];
            }

            

            if (!TryParseAddress(address, out int startRow, out int startColumn))
            {
                Finish(da, add, "Address is invalid. Use format like A1.");
                return;
            }

            Dictionary<string, List<object>> dictionary = ConvertTreeToDictionary(
                tree,
                headerOverrides,
                out List<string> columnKeys,
                out List<string> headers);

            if (dictionary.Count == 0)
            {
                Finish(da, add, "Tree is empty.");
                return;
            }

            // Attach/open workbook
            Excel.Application app;
            Excel.Workbook wb;
            ExcelHelpers.AttachOrOpenWorkbook(out app, out wb, workbookPath, visible: visible);

            // Write
            message = ExcelHelpers.WriteDictionaryToWorksheet(
                dictionary,
                headers,
                columnKeys,
                wb,                 // reuse existing workbook
                worksheetName,      // create if missing
                startRow,
                startColumn,
                address,            // for message text only
                saveAfterWrite,
                readOnly);

            Finish(da, add, message);
        }

        private static Dictionary<string, List<object>> ConvertTreeToDictionary(
            GH_Structure<IGH_Goo> tree,
            IList<string> headerOverrides,
            out List<string> columnKeys,
            out List<string> headers)
        {
            headers = new List<string>();
            columnKeys = new List<string>();
            var dict = new Dictionary<string, List<object>>();
            if (tree == null) return dict;

            // Track used labels to guarantee uniqueness
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int index = 0;
            bool headersProvided = headerOverrides != null && headerOverrides.Count > 0;

            foreach (GH_Path path in tree.Paths)
            {
                var branch = tree.get_Branch(path);
                var list = new List<object>();
                if (branch != null)
                {
                    foreach (IGH_Goo goo in branch)
                        list.Add(GooToExcelValue(goo));
                }

                string key = path.ToString();
                dict[key] = list;
                columnKeys.Add(key);

                // If user provided a headers list, blanks fall back to "Header", "Header_1", ...
                // If no headers list at all, use the branch path string (original behavior).
                string desired = (headersProvided && index < headerOverrides.Count) ? headerOverrides[index] : null;
                string fallback = headersProvided ? DefaultHeaderBase : key;

                string headerLabel = ResolveHeader(desired, fallback, used);
                headers.Add(headerLabel);

                index++;
            }

            return dict;
        }

        private static string ResolveHeader(string desired, string fallback, HashSet<string> used)
        {
            string baseHeader = string.IsNullOrWhiteSpace(desired) ? fallback : desired.Trim();
            if (string.IsNullOrEmpty(baseHeader))
                baseHeader = DefaultHeaderBase;

            string candidate = baseHeader;
            int suffix = 1;
            while (used.Contains(candidate))
                candidate = string.Format(CultureInfo.InvariantCulture, "{0}_{1}", baseHeader, suffix++);

            used.Add(candidate);
            return candidate;
        }

        private static object GooToExcelValue(IGH_Goo goo)
        {
            if (goo == null) return string.Empty;

            try
            {
                if (goo is GH_Number num) return num.Value;
                if (goo is GH_Integer ghInt) return ghInt.Value;
                if (goo is GH_Boolean ghBool) return ghBool.Value;
                if (goo is GH_String ghString) return ghString.Value ?? string.Empty;
                if (goo is GH_Time ghTime) return ghTime.Value;
                if (goo is GH_ComplexNumber ghComplex) return ghComplex.ToString();

                if (goo.CastTo(out string casted)) return casted;

                object script = goo.ScriptVariable();
                if (script == null) return goo.ToString();

                if (script is string) return script;
                if (script is bool || script is int || script is double || script is float || script is decimal)
                    return script;

                return script.ToString();
            }
            catch
            {
                return goo.ToString();
            }
        }

        private static bool TryParseAddress(string address, out int row, out int column)
        {
            row = 1;
            column = 1;

            if (string.IsNullOrWhiteSpace(address))
                return true;

            Match match = Regex.Match(address.Trim(), "^([A-Za-z]+)(\\d+)$");
            if (!match.Success) return false;

            column = ColumnLettersToNumber(match.Groups[1].Value);
            if (!int.TryParse(match.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out row))
                row = 1;

            return column > 0 && row > 0;
        }

        private static int ColumnLettersToNumber(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters)) return 1;

            int result = 0;
            string upper = letters.ToUpperInvariant();
            for (int i = 0; i < upper.Length; i++)
            {
                char c = upper[i];
                if (c < 'A' || c > 'Z') return 1;
                result = result * 26 + (c - 'A' + 1);
            }

            return Math.Max(1, result);
        }

        private void Finish(IGH_DataAccess da, bool add, string message)
        {
            lastMsg = message ?? string.Empty;
            da.SetData(0, lastMsg);
            lastAdd = add;
        }
    }
}
