using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
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
        /// Force-restart Excel: if a running instance exists, close it completely (discard unsaved),
        /// then start a fresh Excel instance and open the workbook at 'filePathOrRelative'.
        /// Returns true if a new Excel Application was created (always true on success with this strategy).
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

                // Resolve full path and validate existence (avoid Excel popups)
                string path = ProjectRelative(filePathOrRelative);
                if (string.IsNullOrWhiteSpace(path)) return false;

                string fullPath = Path.GetFullPath(path);
                if (!File.Exists(fullPath)) return false;

                // 0) If an Excel instance is running, close it completely to start clean.
                var running = TryGetRunningExcelApplication();
                if (running != null)
                {
                    SafeQuitExcel(running); // discards unsaved changes, releases RCWs, GC x2
                    running = null;
                }

                // 1) Create a fresh Excel instance and open the workbook (no prompts)
                bool createdApplication = false;
                try
                {
                    app = new Excel.Application();
                    createdApplication = true;

                    bool prevAlerts = false;
                    try { prevAlerts = app.DisplayAlerts; app.DisplayAlerts = false; } catch { }

                    try
                    {
                        wb = app.Workbooks.Open(
                            Filename: fullPath,
                            UpdateLinks: 0,
                            ReadOnly: readOnly,
                            IgnoreReadOnlyRecommended: true,
                            AddToMru: false
                        );
                    }
                    catch
                    {
                        // Fallback: try read-only once
                        if (!readOnly)
                        {
                            wb = app.Workbooks.Open(
                                Filename: fullPath,
                                UpdateLinks: 0,
                                ReadOnly: true,
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false
                            );
                        }
                        else
                        {
                            throw;
                        }
                    }
                    finally
                    {
                        try { app.DisplayAlerts = prevAlerts; } catch { }
                    }

                    // 2) UI polish
                    try
                    {
                        if (visible) { app.Visible = true; app.UserControl = true; }
                        wb.Activate();
                    }
                    catch { /* no-op */ }

                    // With the force-restart strategy, a new Excel.exe is always spawned on success
                    return createdApplication; // true
                }
                catch
                {
                    // Cleanup on failure
                    try { if (createdApplication) app?.Quit(); } catch { }
                    try { ReleaseCom(wb); } catch { }
                    try { ReleaseCom(app); } catch { }
                    wb = null; app = null;
                    return false;
                }
            }
        }

        /// <summary>
        /// Close all workbooks (no prompts), quit the app, release RCWs, then double-GC.
        /// </summary>
        private static void SafeQuitExcel(Excel.Application excel)
        {
            if (excel == null) return;

            bool prevAlerts = false;
            try { prevAlerts = excel.DisplayAlerts; excel.DisplayAlerts = false; } catch { }
            try { excel.ScreenUpdating = false; } catch { }
            try { excel.UserControl = false; } catch { }

            Excel.Workbooks books = null;
            try
            {
                books = excel.Workbooks;
                // Close from last to first; do NOT use foreach (creates hidden enumerator RCWs)
                for (int i = books.Count; i >= 1; i--)
                {
                    Excel.Workbook wb = null;
                    try { wb = books[i]; wb.Close(SaveChanges: false); }
                    catch { /* ignore */ }
                    finally { ReleaseCom(wb); }
                }
            }
            catch { }
            finally { ReleaseCom(books); }

            try { excel.Quit(); } catch { }
            ReleaseCom(excel);

            // Ensure COM proxies are finalized so EXCEL.EXE actually exits
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        /// <summary>
        /// Try to get a running Excel instance WITHOUT Marshal.GetActiveObject.
        /// 1) VB Interaction.GetObject
        /// 2) HWND/Accessibility (FindWindowEx + AccessibleObjectFromWindow)
        /// Returns null if not found / not accessible.
        /// </summary>
        private static Excel.Application TryGetRunningExcelApplication()
        {
            Excel.Application result = null;
            Exception captured = null;

            void Probe()
            {
                // --- Route 1: VB Interaction ---
                try
                {
                    object obj = Interaction.GetObject(null, "Excel.Application");
                    result = obj as Excel.Application;
                    if (result != null) return;

                    // If late-bound object, try Application property
                    if (obj != null)
                    {
                        var appObj = obj.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, obj, null);
                        result = appObj as Excel.Application;
                        if (result != null) return;
                    }
                }
                catch (Exception ex)
                {
                    captured = ex; // fall through to HWND route
                    result = null;
                }

                // --- Route 2: HWND / Accessibility ---
                try
                {
                    IntPtr prev = IntPtr.Zero;
                    while (true)
                    {
                        IntPtr xlMain = FindWindowEx(IntPtr.Zero, prev, "XLMAIN", null);
                        if (xlMain == IntPtr.Zero) break;
                        prev = xlMain;

                        if (TryGetAppFromHwnd(xlMain, out result) && result != null) return;

                        IntPtr xlDesk = FindWindowEx(xlMain, IntPtr.Zero, "XLDESK", null);
                        if (xlDesk != IntPtr.Zero)
                        {
                            IntPtr excel7 = FindWindowEx(xlDesk, IntPtr.Zero, "EXCEL7", null);
                            if (excel7 != IntPtr.Zero)
                            {
                                if (TryGetAppFromHwnd(excel7, out result) && result != null) return;

                                if (TryGetComFromHwnd(excel7, out object sheetObj) && sheetObj != null)
                                {
                                    try
                                    {
                                        var appObj = sheetObj.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, sheetObj, null);
                                        result = appObj as Excel.Application;
                                        if (result != null) return;
                                    }
                                    catch { /* ignore */ }
                                    finally
                                    {
                                        try { if (Marshal.IsComObject(sheetObj)) Marshal.FinalReleaseComObject(sheetObj); } catch { }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    captured = ex;
                    result = null;
                }
            }

            // Ensure STA for COM calls
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            {
                Probe();
            }
            else
            {
                var t = new Thread(Probe) { IsBackground = true };
                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();
            }

            if (result == null && captured != null)
                Debug.WriteLine("Excel not reachable via Interaction/AccessibleObject: " + captured.Message);

            return result; // null if no instance
        }

        // ---- Win32 + Accessibility interop helpers ----
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd, uint dwObjectID, ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        private static bool TryGetComFromHwnd(IntPtr hwnd, out object com)
        {
            com = null;
            try
            {
                // COPY readonly field ra biến local
                Guid iid = IID_IDispatch;

                int hr = AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, ref iid, out object obj);
                if (hr >= 0 && obj != null)
                {
                    com = obj;
                    return true;
                }
            }
            catch { }
            return false;
        }

        private static bool TryGetAppFromHwnd(IntPtr hwnd, out Excel.Application app)
        {
            app = null;
            if (TryGetComFromHwnd(hwnd, out object obj) && obj != null)
            {
                // Try cast directly
                app = obj as Excel.Application;
                if (app != null) return true;

                // Or get Application property
                try
                {
                    var appObj = obj.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, obj, null);
                    app = appObj as Excel.Application;
                    if (app != null) return true;
                }
                catch { }
                finally
                {
                    try { if (Marshal.IsComObject(obj)) Marshal.FinalReleaseComObject(obj); } catch { }
                }
            }
            return false;
        }

        /// <summary>
        /// Get a worksheet by name; create if missing. Returns a WS RCW owned by the caller.
        /// </summary>
        internal static Excel.Worksheet GetOrCreateWorksheet(Excel.Workbook wb, string sheetName)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

            Excel.Worksheet ws = null;
            try
            {
                // 1-based loop and release intermediates to avoid leaks
                for (int i = 1; i <= wb.Worksheets.Count; i++)
                {
                    Excel.Worksheet s = null;
                    try
                    {
                        s = (Excel.Worksheet)wb.Worksheets[i];
                        if (string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            ws = s; // transfer ownership to caller
                            s = null;
                            break;
                        }
                    }
                    finally
                    {
                        if (s != null) ReleaseCom(s);
                    }
                }
            }
            catch { /* ignore and try create */ }

            if (ws == null)
            {
                Excel.Sheets sheets = null;
                try
                {
                    sheets = wb.Worksheets;
                    ws = (Excel.Worksheet)sheets.Add(After: sheets[sheets.Count]);
                    try { ws.Name = sheetName; } catch { /* could clash, leave default */ }
                }
                finally
                {
                    ReleaseCom(sheets);
                }
            }

            return ws;
        }

        /// <summary>
        /// Using existing workbook/worksheet (no attach/open here). Writes a ragged column dictionary starting at (startRow,startColumn).
        /// </summary>
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

                // Build column order (explicit order preferred, otherwise dictionary order + append any missing keys)
                var columnKeys = new List<string>();
                if (columnOrder != null && columnOrder.Count > 0)
                    columnKeys.AddRange(columnOrder);
                foreach (var k in data.Keys)
                    if (!columnKeys.Contains(k)) columnKeys.Add(k);

                int colCount = columnKeys.Count;
                if (colCount == 0) return "Dictionary is empty.";

                int maxRows = 0;
                foreach (var k in columnKeys)
                {
                    if (data.TryGetValue(k, out var lst) && lst != null && lst.Count > maxRows)
                        maxRows = lst.Count;
                }

                int totalRows = Math.Max(1, maxRows + 1); // +1 for header row
                var values = new object[totalRows, colCount];

                for (int c = 0; c < colCount; c++)
                {
                    string key = columnKeys[c] ?? string.Empty;
                    string headerLabel = (headers != null && c < headers.Count && !string.IsNullOrWhiteSpace(headers[c]))
                        ? headers[c]
                        : key;

                    values[0, c] = headerLabel ?? string.Empty;

                    if (!data.TryGetValue(key, out var branch) || branch == null) continue;
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
                ReleaseCom(ws); // WS RCW created here; caller still owns the workbook
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
