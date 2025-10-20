using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

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

                // 1) Try to attach to running Excel
                try { app = Marshal.GetActiveObject("Excel.Application") as Excel.Application; }
                catch { /* none running */ }

                // 2) Create application if none
                if (app == null)
                {
                    app = new Excel.Application();
                    createdApplication = true;
                }

                // ALWAYS honor visibility (even on attach)
                try
                {
                    if (visible)
                    {
                        app.Visible = true;
                        app.UserControl = true;
                        app.WindowState = Excel.XlWindowState.xlMaximized;
                    }
                }
                catch { /* ignore */ }

                // 3) If workbook already open in this app, bind to it
                try
                {
                    string fullPath = Path.GetFullPath(path);
                    string fileName = Path.GetFileName(fullPath);

                    for (int i = 1; i <= app.Workbooks.Count; i++)
                    {
                        var w = app.Workbooks[i];
                        bool same = false;
                        try
                        {
                            same = string.Equals(
                                Path.GetFullPath(w.FullName),
                                fullPath,
                                StringComparison.OrdinalIgnoreCase);
                        }
                        catch { }

                        if (same || string.Equals(w.Name, fileName, StringComparison.OrdinalIgnoreCase))
                        {
                            wb = w;
                            break;
                        }
                    }
                }
                catch { /* ignore */ }

                // 4) Open or create the workbook
                if (wb == null)
                {
                    if (File.Exists(path))
                    {
                        wb = app.Workbooks.Open(path, ReadOnly: readOnly);
                    }
                    else
                    {
                        string dir = Path.GetDirectoryName(path);
                        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                            Directory.CreateDirectory(dir);

                        wb = app.Workbooks.Add();
                        wb.SaveAs(path, Excel.XlFileFormat.xlOpenXMLWorkbook);
                    }
                }

                // 5) Always activate before returning (no input needed)
                try { wb.Activate(); } catch { /* ignore */ }

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

        internal static string WriteDictionaryToWorksheet(
            IDictionary<string, List<object>> data,
            IList<string> headers,
            string workbookPath,
            string worksheetName,
            int startRow,
            int startColumn,
            string startAddress,
            bool visible,
            bool saveAfterWrite,
            bool readOnly)
        {
            if (data == null || data.Count == 0)
                return "Dictionary is empty.";

            Excel.Application app = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;
            Excel.Range topLeft = null;
            Excel.Range bottomRight = null;
            bool createdApp = false;

            try
            {
                createdApp = AttachOrOpenWorkbook(out app, out wb, workbookPath, visible, readOnly);
                if (wb == null)
                    return "Failed to open workbook.";

                ws = GetOrCreateWorksheet(wb, worksheetName);
                if (ws == null)
                    return "Failed to access worksheet.";

                List<string> columnKeys = headers != null && headers.Count > 0
                    ? new List<string>(headers)
                    : new List<string>(data.Keys);

                foreach (var key in data.Keys)
                {
                    if (!columnKeys.Contains(key))
                        columnKeys.Add(key);
                }

                int columnCount = columnKeys.Count;
                if (columnCount == 0)
                    return "Dictionary is empty.";

                int maxBranchCount = 0;
                foreach (string key in columnKeys)
                {
                    if (data.TryGetValue(key, out var list) && list != null && list.Count > maxBranchCount)
                        maxBranchCount = list.Count;
                }

                int totalRows = Math.Max(1, maxBranchCount + 1);
                object[,] values = new object[totalRows, columnCount];

                for (int col = 0; col < columnCount; col++)
                {
                    string header = columnKeys[col] ?? string.Empty;
                    values[0, col] = header;

                    if (!data.TryGetValue(header, out var branch) || branch == null)
                        continue;

                    int count = branch.Count;
                    for (int row = 0; row < count; row++)
                    {
                        values[row + 1, col] = branch[row];
                    }
                }

                topLeft = (Excel.Range)ws.Cells[startRow, startColumn];
                bottomRight = (Excel.Range)ws.Cells[startRow + totalRows - 1, startColumn + columnCount - 1];
                range = ws.Range[topLeft, bottomRight];
                range.Value2 = values;

                if (saveAfterWrite && !readOnly)
                {
                    try { wb.Save(); }
                    catch { /* ignore */ }
                }

                string wsName = worksheetName;
                try { wsName = ws?.Name ?? worksheetName; } catch { }

                string startLabel = string.IsNullOrWhiteSpace(startAddress)
                    ? ColumnNumberToLetters(startColumn) + Math.Max(1, startRow).ToString(CultureInfo.InvariantCulture)
                    : startAddress.ToUpperInvariant();

                return string.Format(CultureInfo.InvariantCulture,
                    "Wrote {0} columns × {1} rows to '{2}' starting at {3}.",
                    columnCount,
                    totalRows,
                    wsName,
                    startLabel);
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
                ReleaseCom(ws);

                if (wb != null && readOnly && createdApp)
                {
                    try { wb.Close(false); }
                    catch { }
                }

                ReleaseCom(wb);

                if (app != null && createdApp && !visible)
                {
                    try { app.Quit(); }
                    catch { }
                }

                ReleaseCom(app);
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
