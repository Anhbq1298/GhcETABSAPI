using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.CSharp.RuntimeBinder;
using Microsoft.VisualBasic;



namespace GhcETABSAPI
{
    internal class ExcelHelpers
    {
        // Prevent overlapping COM calls in multi-threaded contexts
        private static readonly object _excelLock = new object();


        /// <summary>
        /// Convert a relative path to full path using AppDomain.BaseDirectory.
        /// Returns full path if input is already absolute; null/empty stays null.
        /// </summary>
        internal static string ProjectRelative(string filePathOrRelative)
        {
            if (string.IsNullOrWhiteSpace(filePathOrRelative)) return null;
            try
            {
                if (Path.IsPathRooted(filePathOrRelative))
                    return Path.GetFullPath(filePathOrRelative);

                string baseDir = AppDomain.CurrentDomain.BaseDirectory; // e.g. bin\Debug
                return Path.GetFullPath(Path.Combine(baseDir, filePathOrRelative));
            }
            catch
            {
                return filePathOrRelative;
            }
        }

        /// <summary>
        /// Release a COM RCW immediately (safe no-op for managed objects).
        /// </summary>
        internal static void ReleaseCom(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch { /* ignore */ }
        }

        /// <summary>
        /// Attach to a running Excel instance and bind to the target workbook if it is already open.
        /// Otherwise, open the workbook from the provided path, creating a new .xlsx if the file does not exist.
        /// ALWAYS activates the workbook before returning. If 'visible' is true, Excel will be shown even when attaching.
        /// Returns true if a new Excel Application was created by this method.
        /// </summary>
        internal static bool AttachOrOpenWorkbook(
            out Excel.Application app,
            out Excel.Workbook wb,
            string filePathOrRelative,
            bool visible = true,
            bool readOnly = false)
        {
            lock (_excelLock)
            {
                app = null;
                wb = null;
                bool createdApplication = false;

                string path = ProjectRelative(filePathOrRelative);
                if (string.IsNullOrWhiteSpace(path)) return false;

                string fullPath = Path.GetFullPath(path);

                // 1) Try bind to an *already-open* workbook by FILE (no Excel option changes)
                try
                {
                    if (File.Exists(fullPath))
                    {
                        var obj = Interaction.GetObject(fullPath);    // attaches if that workbook is open
                        wb = obj as Excel.Workbook;                   // OR starts Excel & opens file if not already open
                        if (wb != null) app = wb.Application;
                    }
                }
                catch
                {
                    wb = null; app = null; // ignore and fall through
                }

                // 2) If not bound, create our own Excel and open / create workbook
                if (wb == null || app == null)
                {
                    app = new Excel.Application();
                    createdApplication = true;

                    try
                    {
                        app.Visible = visible;
                        if (visible) app.UserControl = true;
                    }
                    catch { }

                    if (File.Exists(fullPath))
                    {
                        wb = app.Workbooks.Open(fullPath, ReadOnly: readOnly);
                    }
                    else
                    {
                        string dir = Path.GetDirectoryName(fullPath);
                        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                        {
                            try { Directory.CreateDirectory(dir); } catch { }
                        }

                        wb = app.Workbooks.Add();
                        try { wb.SaveAs(fullPath, Excel.XlFileFormat.xlOpenXMLWorkbook); } catch { }
                    }
                }

                // 3) Finish
                try { if (visible) app.Visible = true; } catch { }
                try { wb.Activate(); } catch { }

                return createdApplication;
            }
        }

        /// <summary>
        /// Get a worksheet by name; create if missing.
        /// </summary>
        internal static Excel.Worksheet GetOrCreateWorksheet(Excel.Workbook wb, string sheetName)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

            Excel.Worksheet ws = null;
            try
            {
                for (int i = 1; i <= wb.Worksheets.Count; i++)
                {
                    var s = (Excel.Worksheet)wb.Worksheets[i];
                    if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ws = s;
                        break;
                    }
                }
            }
            catch { /* ignore and try create */ }

            if (ws == null)
            {
                ws = (Excel.Worksheet)wb.Worksheets.Add(After: wb.Worksheets[wb.Worksheets.Count]);
                try { ws.Name = sheetName; } catch { /* could clash, leave default */ }
            }

            return ws;
        }

        // Using existing workbook/worksheet (no attach/open here)
        internal static string WriteDictionaryToWorksheet(
            IDictionary<string, List<object>> data,
            IList<string> headers,
            IList<string> columnOrder,
            Excel.Workbook wb,          // reuse existing workbook (no attach/open here)
            string worksheetName,       // sheet name instead of Excel.Worksheet
            int startRow,
            int startColumn,
            string startAddress,
            bool saveAfterWrite,
            bool readOnly)
        {
            if (data == null || data.Count == 0) return "Dictionary is empty.";
            if (wb == null) return "Workbook is null.";
            if (string.IsNullOrWhiteSpace(worksheetName)) worksheetName = "Sheet1";

            Excel.Worksheet ws = null;
            Excel.Range range = null, topLeft = null, bottomRight = null;

            try
            {
                // Get or create the target worksheet
                ws = GetOrCreateWorksheet(wb, worksheetName);
                if (ws == null) return "Failed to access worksheet.";

                // Build column order (explicit order preferred, otherwise dictionary order)
                List<string> columnKeys = null;
                if (columnOrder != null && columnOrder.Count > 0)
                {
                    columnKeys = new List<string>(columnOrder);
                }
                else if (data != null)
                {
                    columnKeys = new List<string>(data.Keys);
                }
                else
                {
                    columnKeys = new List<string>();
                }

                if (data != null)
                {
                    foreach (var key in data.Keys)
                    {
                        if (!columnKeys.Contains(key))
                        {
                            columnKeys.Add(key);
                        }
                    }
                }

                int colCount = columnKeys.Count;
                if (colCount == 0) return "Dictionary is empty.";

                int maxRows = 0;
                foreach (var k in columnKeys)
                {
                    if (data != null && data.TryGetValue(k, out var lst) && lst != null && lst.Count > maxRows)
                    {
                        maxRows = lst.Count;
                    }
                }

                int totalRows = Math.Max(1, maxRows + 1); // +1 for header row
                var values = new object[totalRows, colCount];

                for (int c = 0; c < colCount; c++)
                {
                    string key = columnKeys[c] ?? string.Empty;
                    string headerLabel = key;
                    if (headers != null && c < headers.Count && !string.IsNullOrWhiteSpace(headers[c]))
                    {
                        headerLabel = headers[c];
                    }

                    values[0, c] = headerLabel ?? string.Empty;

                    if (data == null || !data.TryGetValue(key, out var branch) || branch == null) continue;
                    for (int r = 0; r < branch.Count; r++)
                        values[r + 1, c] = branch[r];
                }

                // Write block
                startRow = Math.Max(1, startRow);
                startColumn = Math.Max(1, startColumn);

                topLeft = (Excel.Range)ws.Cells[startRow, startColumn];
                bottomRight = (Excel.Range)ws.Cells[startRow + totalRows - 1, startColumn + colCount - 1];
                range = ws.Range[topLeft, bottomRight];
                range.Value2 = values;

                if (saveAfterWrite && !readOnly)
                {
                    try { wb.Save(); } catch { /* ignore */ }
                }

                string wsName = ws?.Name ?? worksheetName;
                string startLabel = string.IsNullOrWhiteSpace(startAddress)
                    ? ColumnNumberToLetters(startColumn) + Math.Max(1, startRow).ToString(CultureInfo.InvariantCulture)
                    : startAddress.ToUpperInvariant();

                return string.Format(CultureInfo.InvariantCulture,
                    "Wrote {0} columns × {1} rows to '{2}' starting at {3}.",
                    colCount, totalRows, wsName, startLabel);
            }
            catch (Exception ex)
            {
                return "Failed: " + ex.Message;
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(topLeft);
                ReleaseCom(bottomRight);
                ReleaseCom(ws); // we created the WS RCW here; caller still owns the workbook
            }
        }

        private static string ColumnNumberToLetters(int column)
        {
            if (column < 1) column = 1;

            StringBuilder builder = new StringBuilder();
            int dividend = column;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                builder.Insert(0, (char)('A' + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return builder.Length > 0 ? builder.ToString() : "A";
        }

    }
}
