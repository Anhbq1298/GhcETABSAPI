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

namespace MGT
{
    /// <summary>
    /// Excel COM helpers for .NET 8 (net8.0-windows).
    /// - No Microsoft.Office.Core reference required.
    /// - Hide Ribbon via Excel4 macro; Zoom-to-selection via reflection (ExecuteMso).
    /// - Kiosk view (no Table): focus a range, hide chrome, keep sheet tabs, no maximize by default.
    /// - Bring to front supported.
    /// </summary>
    internal class ExcelHelpers
    {
        // Prevent overlapping COM calls in multi-threaded contexts
        private static readonly object _excelLock = new object();

        // -------- UI STATE (INSIDE THIS CLASS) --------
        /// <summary>Snapshot of Excel UI state so we can restore later.</summary>
        internal class UiState
        {
            public bool? DisplayFullScreen;
            public bool? DisplayFormulaBar;
            public bool? DisplayStatusBar;
            public bool? ScreenUpdating;
            public bool? EnableEvents;

            // Per-window toggles
            public bool? DisplayGridlines;
            public bool? DisplayHeadings;
            public bool? DisplayHorizontalScrollBar;
            public bool? DisplayVerticalScrollBar;
            public bool? DisplayWorkbookTabs;
            public int? Zoom;
        }
        // ----------------------------------------------

        #region Path / COM helpers

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

        internal static string TrimOrEmpty(object value)
        {
            if (value == null)
            {
                return string.Empty;
            }

            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();
        }

        internal static bool IsNullOrEmpty(object value)
        {
            if (value == null)
            {
                return true;
            }

            if (value is string s)
            {
                return string.IsNullOrWhiteSpace(s);
            }

            if (value is double d)
            {
                return double.IsNaN(d) || double.IsInfinity(d);
            }

            return false;
        }

        internal static Excel.Worksheet FindWorksheet(Excel.Workbook workbook, string sheetName)
        {
            if (workbook == null)
            {
                return null;
            }

            Excel.Sheets sheets = null;

            try
            {
                sheets = TryGetSheets(workbook);
                int count = sheets?.Count ?? 0;

                for (int i = 1; i <= count; i++)
                {
                    Excel.Worksheet candidate = null;
                    try
                    {
                        candidate = sheets[i] as Excel.Worksheet;
                        if (candidate != null && string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.Worksheet result = candidate;
                            candidate = null;
                            return result;
                        }
                    }
                    finally
                    {
                        ReleaseCom(candidate);
                    }
                }

                return null;
            }
            finally
            {
                ReleaseCom(sheets);
            }
        }

        #region Worksheet reading helper

        internal sealed class ExcelSheetProfile
        {
            internal string ExpectedSheetName { get; init; }
            internal int StartColumn { get; init; } = 1;
            internal IReadOnlyList<string> ExpectedHeaders { get; init; }
            internal string ProgressLabel { get; init; } = "Reading Excel";
            internal string CompletionLabel { get; init; } = "Excel Done";
            internal string ProgressUnit { get; init; } = "row";
        }

        internal sealed class ExcelSheetReadResult
        {
            internal List<string> Headers { get; } = new List<string>();
            internal List<object[]> Rows { get; } = new List<object[]>();
        }

        internal static ExcelSheetReadResult ReadSheet(
            string fullPath,
            string requestedSheetName,
            ExcelSheetProfile profile,
            Action<int, int, string> progressCallback = null)
        {
            if (string.IsNullOrWhiteSpace(fullPath))
            {
                throw new ArgumentException("Excel path cannot be empty.", nameof(fullPath));
            }

            if (profile == null)
            {
                throw new ArgumentNullException(nameof(profile));
            }

            if (profile.ExpectedHeaders == null)
            {
                throw new ArgumentException("Profile must define ExpectedHeaders.", nameof(profile));
            }

            string sheetToRead = string.IsNullOrWhiteSpace(requestedSheetName)
                ? profile.ExpectedSheetName
                : requestedSheetName;

            if (!string.IsNullOrEmpty(profile.ExpectedSheetName) &&
                !string.Equals(sheetToRead, profile.ExpectedSheetName, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    $"Invalid workbook: expected sheet name '{profile.ExpectedSheetName}'.");
            }

            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range usedRange = null;

            try
            {
                app = new Excel.Application();

                TrySetApplicationBool(app, "Visible", false);
                TrySetApplicationBool(app, "DisplayAlerts", false);
                TrySetApplicationBool(app, "UserControl", false);

                books = TryGetWorkbooks(app);
                wb = TryOpenWorkbook(
                    books,
                    Filename: fullPath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                ws = FindWorksheet(wb, sheetToRead);
                if (ws == null)
                {
                    throw new InvalidOperationException(
                        $"Worksheet '{sheetToRead}' not found in '{Path.GetFileName(fullPath)}'.");
                }

                ExcelSheetReadResult result = new ExcelSheetReadResult();

                int startColumn = Math.Max(1, profile.StartColumn);
                int columnCount = profile.ExpectedHeaders.Count;

                for (int col = 0; col < columnCount; col++)
                {
                    Excel.Range headerCell = null;
                    try
                    {
                        headerCell = (Excel.Range)ws.Cells[1, startColumn + col];
                        string headerValue = TrimOrEmpty(headerCell?.Value2);
                        result.Headers.Add(headerValue);

                        string expectedHeader = profile.ExpectedHeaders[col];
                        if (!string.Equals(headerValue, expectedHeader, StringComparison.OrdinalIgnoreCase))
                        {
                            char columnLetter = (char)('A' + startColumn + col - 1);
                            throw new InvalidOperationException(
                                $"Invalid workbook: expected header '{expectedHeader}' in column {columnLetter}, found '{headerValue}'.");
                        }
                    }
                    finally
                    {
                        ReleaseCom(headerCell);
                    }
                }

                usedRange = ws.UsedRange;
                int lastRow = 1;
                if (usedRange != null)
                {
                    try
                    {
                        lastRow = Math.Max(lastRow, usedRange.Row + usedRange.Rows.Count - 1);
                    }
                    catch
                    {
                        lastRow = 1;
                    }
                }

                int totalRows = Math.Max(0, lastRow - 1);
                progressCallback?.Invoke(
                    0,
                    totalRows,
                    UiHelpers.FormatProgressStatus(0, totalRows, profile.ProgressLabel, profile.ProgressUnit));

                int processedRows = 0;

                for (int row = 2; row <= lastRow; row++)
                {
                    object[] rowValues = new object[columnCount];
                    bool hasData = false;

                    for (int col = 0; col < columnCount; col++)
                    {
                        Excel.Range cell = null;
                        try
                        {
                            cell = (Excel.Range)ws.Cells[row, startColumn + col];
                            object value = cell?.Value2;
                            rowValues[col] = value;
                            if (!IsNullOrEmpty(value))
                            {
                                hasData = true;
                            }
                        }
                        finally
                        {
                            ReleaseCom(cell);
                        }
                    }

                    processedRows++;
                    int current = totalRows > 0 ? Math.Min(processedRows, totalRows) : processedRows;
                    progressCallback?.Invoke(
                        current,
                        totalRows,
                        UiHelpers.FormatProgressStatus(current, totalRows, profile.ProgressLabel, profile.ProgressUnit));

                    if (!hasData)
                    {
                        continue;
                    }

                    result.Rows.Add(rowValues);
                }

                progressCallback?.Invoke(
                    result.Rows.Count,
                    result.Rows.Count,
                    UiHelpers.FormatCompletionStatus(result.Rows.Count, profile.CompletionLabel, profile.ProgressUnit));

                return result;
            }
            finally
            {
                ReleaseCom(usedRange);

                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                }

                ReleaseCom(ws);
                ReleaseCom(wb);
                ReleaseCom(books);

                if (app != null)
                {
                    try { app.Quit(); } catch { }
                }

                ReleaseCom(app);
            }
        }

        #endregion

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

        #endregion

        #region Attach / Open / Close Excel

        /// <summary>
        /// Force-restart Excel: close any running instance, then open workbook.
        /// If 'filePathOrRelative' is null/empty, creates a temp workbook (temp.xlsx).
        /// Returns true if a new Excel Application was created (always true on success here).
        /// </summary>
        internal static bool AttachOrOpenWorkbook(
            out Excel.Application app,
            out Excel.Workbook wb,
            string filePathOrRelative,
            bool visible = true,
            bool readOnly = false,
            bool maximizeWindow = false,   // default false (do not maximize)
            bool bringToFront = true)      // default true (bring window to front)
        {
            lock (_excelLock)
            {
                app = null;
                wb = null;

                bool createTemporaryWorkbook = string.IsNullOrWhiteSpace(filePathOrRelative);

                string fullPath = null;
                if (!createTemporaryWorkbook)
                {
                    fullPath = filePathOrRelative;
                    if (!File.Exists(fullPath)) return false;
                }

                // 0) Close any running Excel for a clean start
                var running = TryGetRunningExcelApplication();
                if (running != null) SafeQuitExcel(running);

                // 1) Create fresh Excel & open workbook (no prompts)
                bool createdApplication = false;
                try
                {
                    app = new Excel.Application();
                    createdApplication = true;

                    bool prevAlerts = TryGetApplicationBool(app, "DisplayAlerts") ?? false;
                    TrySetApplicationBool(app, "DisplayAlerts", false);

                    try
                    {
                        if (createTemporaryWorkbook)
                        {
                            wb = TryAddWorkbook(app);

                            string tempDirectory = Path.GetTempPath();
                            string tempFileName = "temp.xlsx";
                            string tempFullPath = Path.Combine(tempDirectory, tempFileName);

                            try { if (File.Exists(tempFullPath)) File.Delete(tempFullPath); } catch { }
                            TrySaveWorkbook(wb, tempFullPath, Excel.XlFileFormat.xlOpenXMLWorkbook);

                            try
                            {
                                Excel.Window window = null;
                                try
                                {
                                    window = TryGetActiveWindow(app);
                                    if (window != null) window.Caption = "temp";
                                }
                                finally { if (window != null) ReleaseCom(window); }
                            }
                            catch { }
                        }
                        else
                        {
                            wb = TryOpenWorkbook(
                                TryGetWorkbooks(app),
                                Filename: fullPath,
                                UpdateLinks: 0,
                                ReadOnly: readOnly,
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false
                            );
                        }
                    }
                    catch
                    {
                        if (!readOnly && !createTemporaryWorkbook)
                        {
                            wb = TryOpenWorkbook(
                                TryGetWorkbooks(app),
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
                        TrySetApplicationBool(app, "DisplayAlerts", prevAlerts);
                    }

                    // 2) UI: show; no maximize unless asked; bring to front if requested
                    try
                    {
                        if (visible)
                        {
                            TrySetApplicationBool(app, "Visible", true);
                            TrySetApplicationBool(app, "UserControl", true);
                        }
                        wb.Activate();

                        if (maximizeWindow) MaximizeExcelWindow(app);
                        if (bringToFront) BringExcelToFront(app);
                    }
                    catch { /* no-op */ }

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
        /// Bring the Excel main window to the foreground (no maximize).
        /// </summary>
        private static void BringExcelToFront(Excel.Application app)
        {
            if (app == null) return;
            try
            {
                // Make sure Excel is visible and the workbook window is active
                TrySetApplicationBool(app, "Visible", true);
                TrySetApplicationBool(app, "UserControl", true);
                try
                {
                    var aw = TryGetActiveWindow(app);
                    if (aw != null)
                    {
                        try { aw.Activate(); } catch { }
                        ReleaseCom(aw);
                    }
                }
                catch { }

                // Bring main HWND to front
                IntPtr hwnd = TryGetApplicationHwnd(app);
                if (hwnd != IntPtr.Zero)
                {
                    ShowWindow(hwnd, SW_RESTORE);
                    SetForegroundWindow(hwnd);
                }
            }
            catch { /* no-op */ }
        }

        /// <summary>
        /// Safely access Excel.Application.ActiveWindow even when interop definitions omit it.
        /// </summary>
        private static Excel.Window TryGetActiveWindow(Excel.Application app)
        {
            if (app == null) return null;
            try
            {
                object aw = app.GetType().InvokeMember(
                    "ActiveWindow",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    app,
                    null);
                return aw as Excel.Window;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Maximize Excel UI window(s) reliably (MDI/SDI): Application & ActiveWindow + Win32.
        /// </summary>
        private static void MaximizeExcelWindow(Excel.Application app)
        {
            if (app == null) return;
                try
                {
                    TrySetApplicationProperty(app, "WindowState", Excel.XlWindowState.xlMaximized);
                    Excel.Window aw = null;
                    try
                    {
                    aw = TryGetActiveWindow(app);
                    if (aw != null)
                    {
                        try { aw.WindowState = Excel.XlWindowState.xlMaximized; } catch { }
                    }
                }
                catch { }
                finally
                {
                    if (aw != null) ReleaseCom(aw);
                }

                IntPtr hwnd = TryGetApplicationHwnd(app);
                if (hwnd != IntPtr.Zero)
                {
                    ShowWindow(hwnd, SW_RESTORE);
                    ShowWindow(hwnd, SW_MAXIMIZE);
                    SetForegroundWindow(hwnd);
                }
            }
            catch { }
        }

        /// <summary>
        /// Late-bind a boolean property on Excel.Application (returns null if unavailable).
        /// Avoids compile-time dependency on interop members that may be missing.
        /// </summary>
        private static bool? TryGetApplicationBool(Excel.Application app, string propertyName)
        {
            if (app == null || string.IsNullOrWhiteSpace(propertyName)) return null;

            try
            {
                object value = app.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    app,
                    null);

                if (value is bool b) return b;
                if (value is int i) return i != 0;
            }
            catch { }

            return null;
        }

        /// <summary>
        /// Late-bind a boolean property setter on Excel.Application (safe no-op on failure).
        /// </summary>
        private static void TrySetApplicationBool(Excel.Application app, string propertyName, bool value)
        {
            if (app == null || string.IsNullOrWhiteSpace(propertyName)) return;

            try
            {
                app.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    app,
                    new object[] { value });
            }
            catch { }
        }

        /// <summary>
        /// Late-bind any property setter on Excel.Application (safe no-op on failure).
        /// </summary>
        private static void TrySetApplicationProperty(Excel.Application app, string propertyName, object value)
        {
            if (app == null || string.IsNullOrWhiteSpace(propertyName)) return;

            try
            {
                app.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.SetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    app,
                    new object[] { value });
            }
            catch { }
        }

        private static IntPtr TryGetApplicationHwnd(Excel.Application app)
        {
            if (app == null) return IntPtr.Zero;

            try
            {
                object value = app.GetType().InvokeMember(
                    "Hwnd",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    app,
                    null);

                if (value is int i) return new IntPtr(i);
                if (value is long l) return new IntPtr(l);
            }
            catch { }

            return IntPtr.Zero;
        }

        private static Excel.Workbooks TryGetWorkbooks(Excel.Application app)
        {
            if (app == null) return null;
            try
            {
                return app.GetType().InvokeMember(
                    "Workbooks",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    app,
                    null) as Excel.Workbooks;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Workbook TryAddWorkbook(Excel.Application app)
        {
            Excel.Workbooks books = null;
            try
            {
                books = TryGetWorkbooks(app);
                return books?.GetType().InvokeMember(
                    "Add",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    books,
                    new object[] { Type.Missing }) as Excel.Workbook;
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(books);
            }
        }

        private static Excel.Workbook TryOpenWorkbook(
            Excel.Workbooks books,
            string Filename,
            int UpdateLinks,
            bool ReadOnly,
            bool IgnoreReadOnlyRecommended,
            bool AddToMru)
        {
            if (books == null) return null;
            try
            {
                return books.GetType().InvokeMember(
                    "Open",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    books,
                    new object[] { Filename, UpdateLinks, ReadOnly, Type.Missing, Type.Missing, Type.Missing, true, "",
                        AddToMru, false, Type.Missing, IgnoreReadOnlyRecommended, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing }) as Excel.Workbook;
            }
            catch
            {
                return null;
            }
        }

        private static void TrySaveWorkbook(Excel.Workbook wb, string path, Excel.XlFileFormat format)
        {
            if (wb == null || string.IsNullOrWhiteSpace(path)) return;
            try
            {
                wb.GetType().InvokeMember(
                    "SaveAs",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    wb,
                    new object[] { path, format });
            }
            catch { }
        }

        private static Excel.Sheets TryGetSheets(Excel.Workbook workbook)
        {
            if (workbook == null) return null;
            try
            {
                return workbook.GetType().InvokeMember(
                    "Worksheets",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    workbook,
                    null) as Excel.Sheets;
            }
            catch
            {
                return null;
            }
        }

        private static int TryGetSheetsCount(Excel.Sheets sheets)
        {
            if (sheets == null) return 0;
            try
            {
                object value = sheets.GetType().InvokeMember(
                    "Count",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    sheets,
                    null);

                if (value is int i) return i;
                if (value is double d) return (int)d;
            }
            catch { }

            return 0;
        }

        private static Excel.Worksheet TryGetWorksheet(Excel.Sheets sheets, int index)
        {
            if (sheets == null) return null;
            try
            {
                return sheets.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    sheets,
                    new object[] { index }) as Excel.Worksheet;
            }
            catch
            {
                return null;
            }
        }

        private static Excel.Worksheet TryAddWorksheet(Excel.Sheets sheets)
        {
            if (sheets == null) return null;
            try
            {
                int count = TryGetSheetsCount(sheets);
                object after = sheets.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    sheets,
                    new object[] { count });

                return sheets.GetType().InvokeMember(
                    "Add",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null,
                    sheets,
                    new object[] { Type.Missing, after, Type.Missing, Type.Missing }) as Excel.Worksheet;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Close all workbooks (no prompts), quit the app, release RCWs, then double-GC.
        /// </summary>
        private static void SafeQuitExcel(Excel.Application excel)
        {
            if (excel == null) return;

            bool prevAlerts = TryGetApplicationBool(excel, "DisplayAlerts") ?? false;
            TrySetApplicationBool(excel, "DisplayAlerts", false);
            TrySetApplicationBool(excel, "ScreenUpdating", false);
            TrySetApplicationBool(excel, "UserControl", false);

            Excel.Workbooks books = null;
            try
            {
                books = TryGetWorkbooks(excel);
                for (int i = books.Count; i >= 1; i--)
                {
                    Excel.Workbook wb = null;
                    try { wb = books[i]; wb.Close(SaveChanges: false); }
                    catch { }
                    finally { ReleaseCom(wb); }
                }
            }
            catch { }
            finally { ReleaseCom(books); }

            try { excel.Quit(); } catch { }
            ReleaseCom(excel);

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

                    if (obj != null)
                    {
                        var appObj = obj.GetType().InvokeMember("Application", BindingFlags.GetProperty, null, obj, null);
                        result = appObj as Excel.Application;
                        if (result != null) return;
                    }
                }
                catch (Exception ex)
                {
                    captured = ex;
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
                                    catch { }
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

            return result;
        }

        // ---- Win32 + Accessibility interop helpers ----
        private const uint OBJID_NATIVEOM = 0xFFFFFFF0;
        private static readonly Guid IID_IDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        private const int SW_RESTORE = 9;
        private const int SW_MAXIMIZE = 3;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user232.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd, uint dwObjectID, ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        private static bool TryGetComFromHwnd(IntPtr hwnd, out object com)
        {
            com = null;
            try
            {
                Guid iid = IID_IDispatch; // local copy for ref param
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
                app = obj as Excel.Application;
                if (app != null) return true;

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

        #endregion

        #region Worksheet helpers

        /// <summary>
        /// Get a worksheet by name; create if missing. Returns a WS RCW owned by the caller.
        /// </summary>
        internal static Excel.Worksheet GetOrCreateWorksheet(Excel.Workbook wb, string sheetName)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

            Excel.Worksheet ws = null;
            Excel.Sheets sheets = null;
            try
            {
                sheets = TryGetSheets(wb);
                int sheetCount = TryGetSheetsCount(sheets);

                for (int i = 1; i <= sheetCount; i++)
                {
                    Excel.Worksheet s = null;
                    try
                    {
                        s = TryGetWorksheet(sheets, i);
                        if (string.Equals(s?.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            ws = s; // transfer ownership
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
            catch { }
            finally
            {
                ReleaseCom(sheets);
            }

            if (ws == null)
            {
                Excel.Sheets sheets = null;
                try
                {
                    sheets = TryGetSheets(wb);
                    ws = TryAddWorksheet(sheets);
                    try { ws.Name = sheetName; } catch { }
                }
                finally
                {
                    ReleaseCom(sheets);
                }
            }

            return ws;
        }

        /// <summary>
        /// Parse "A1" (or "$A$1") to (row, col). Returns false if invalid.
        /// </summary>
        private static bool TryParseA1Address(string a1, out int row, out int col)
        {
            row = 0; col = 0;
            if (string.IsNullOrWhiteSpace(a1)) return false;

            string s = a1.Trim().ToUpperInvariant();
            int i = 0;

            while (i < s.Length && (s[i] == '$' || char.IsLetter(s[i])))
            {
                if (char.IsLetter(s[i]))
                    col = col * 26 + (s[i] - 'A' + 1);
                i++;
            }
            if (col <= 0) return false;

            if (i < s.Length && s[i] == '$') i++;

            int start = i;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            if (start == i) return false;
            if (!int.TryParse(s.Substring(start, i - start), NumberStyles.None, CultureInfo.InvariantCulture, out row)) return false;

            return row > 0 && col > 0;
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

        #endregion

        #region Ribbon / CommandBars (reflection)

        /// <summary>
        /// Late-bind Application.CommandBars.ExecuteMso(controlId) to avoid referencing office.dll.
        /// Works on .NET 8 without Microsoft.Office.Core.
        /// </summary>
        private static void TryExecuteMso(object application, string controlId)
        {
            try
            {
                if (application == null || string.IsNullOrWhiteSpace(controlId)) return;

                // object cmdBars = application.CommandBars;
                var cmdBars = application.GetType().InvokeMember(
                    "CommandBars",
                    BindingFlags.GetProperty,
                    binder: null,
                    target: application,
                    args: null);

                if (cmdBars == null) return;

                // cmdBars.ExecuteMso(controlId);
                cmdBars.GetType().InvokeMember(
                    "ExecuteMso",
                    BindingFlags.InvokeMethod,
                    binder: null,
                    target: cmdBars,
                    args: new object[] { controlId });
            }
            catch
            {
                // Swallow: ExecuteMso may not exist in some Office builds or policies.
            }
        }

        #endregion

        #region Kiosk view (NO Table) + UI restore

        /// <summary>
        /// In-place "kiosk" view WITHOUT converting to ListObject:
        /// - Selects the given dataRange
        /// - Zoom-to-selection (via ExecuteMso reflection), fallback Zoom
        /// - Hides Ribbon, FormulaBar, StatusBar, gridlines, headings, scrollbars
        /// - Keeps sheet tabs visible
        /// - Does NOT maximize; FullScreen only if requested (default false)
        /// Returns captured UiState for optional restore.
        /// </summary>
        internal static UiState ApplyKioskViewNoTable(
            Excel.Application app,
            Excel.Worksheet ws,
            Excel.Range dataRange,
            bool makeFullScreen = false) // default false
        {
            if (app == null || ws == null || dataRange == null) return null;

            // Select focus range
            try { dataRange.Select(); } catch { }

            // Capture current UI state
            UiState state = new UiState();
            try
            {
                state.DisplayFullScreen = TryGetApplicationBool(app, "DisplayFullScreen");
                state.DisplayFormulaBar = TryGetApplicationBool(app, "DisplayFormulaBar");
                state.DisplayStatusBar = TryGetApplicationBool(app, "DisplayStatusBar");
                state.ScreenUpdating = TryGetApplicationBool(app, "ScreenUpdating");
                state.EnableEvents = TryGetApplicationBool(app, "EnableEvents");

                var win = TryGetActiveWindow(app);
                if (win != null)
                {
                    state.DisplayGridlines = win.DisplayGridlines;
                    state.DisplayHeadings = win.DisplayHeadings;
                    state.DisplayHorizontalScrollBar = win.DisplayHorizontalScrollBar;
                    state.DisplayVerticalScrollBar = win.DisplayVerticalScrollBar;
                    state.DisplayWorkbookTabs = win.DisplayWorkbookTabs;
                    state.Zoom = win.Zoom is int z ? z : (int?)null;
                }
            }
            catch { }

            // Hide chrome + zoom to selection
            try
            {
                TrySetApplicationBool(app, "ScreenUpdating", true);
                TrySetApplicationBool(app, "EnableEvents", false);

                // Hide Ribbon (Excel4 macro)
                try { app.ExecuteExcel4Macro(@"SHOW.TOOLBAR(""Ribbon"",False)"); } catch { }

                TrySetApplicationBool(app, "DisplayFormulaBar", false);
                TrySetApplicationBool(app, "DisplayStatusBar", false);

                var win = TryGetActiveWindow(app);
                if (win != null)
                {
                    // Hide most chrome but KEEP sheet tabs
                    win.DisplayGridlines = false;
                    win.DisplayHeadings = false;
                    win.DisplayHorizontalScrollBar = false;
                    win.DisplayVerticalScrollBar = false;
                    win.DisplayWorkbookTabs = true; // keep visible

                    // Zoom-to-selection without office.dll
                    TryExecuteMso(app, "ZoomToSelection");

                    // Fallback zoom if ExecuteMso not available
                    try { if (!(win.Zoom is int)) win.Zoom = 120; } catch { }
                }

                // Do NOT force FullScreen unless requested
                TrySetApplicationBool(app, "DisplayFullScreen", makeFullScreen);
            }
            catch { }
            finally
            {
                TrySetApplicationBool(app, "EnableEvents", true);
            }

            // Scroll to top-left of range
            try
            {
                var win = TryGetActiveWindow(app);
                if (win != null)
                {
                    win.ScrollRow = dataRange.Row;
                    win.ScrollColumn = dataRange.Column;
                }
            }
            catch { }

            return state;
        }

        /// <summary>
        /// Restore Excel UI back from kiosk mode and show Ribbon again.
        /// </summary>
        internal static void RestoreUi(Excel.Application app, UiState state)
        {
            if (app == null || state == null) return;

            TrySetApplicationBool(app, "ScreenUpdating", state.ScreenUpdating ?? true);
            TrySetApplicationBool(app, "EnableEvents", state.EnableEvents ?? true);
            TrySetApplicationBool(app, "DisplayFullScreen", state.DisplayFullScreen ?? false);
            TrySetApplicationBool(app, "DisplayFormulaBar", state.DisplayFormulaBar ?? true);
            TrySetApplicationBool(app, "DisplayStatusBar", state.DisplayStatusBar ?? true);

            // Show Ribbon back
            try { app.ExecuteExcel4Macro(@"SHOW.TOOLBAR(""Ribbon"",True)"); } catch { }

            try
            {
                var win = TryGetActiveWindow(app);
                if (win != null)
                {
                    if (state.DisplayGridlines != null) win.DisplayGridlines = state.DisplayGridlines.Value;
                    if (state.DisplayHeadings != null) win.DisplayHeadings = state.DisplayHeadings.Value;
                    if (state.DisplayHorizontalScrollBar != null) win.DisplayHorizontalScrollBar = state.DisplayHorizontalScrollBar.Value;
                    if (state.DisplayVerticalScrollBar != null) win.DisplayVerticalScrollBar = state.DisplayVerticalScrollBar.Value;
                    if (state.DisplayWorkbookTabs != null) win.DisplayWorkbookTabs = state.DisplayWorkbookTabs.Value;
                    if (state.Zoom != null) win.Zoom = state.Zoom.Value;
                }
            }
            catch { }
        }

        #endregion

        #region Write API (Dictionary -> Worksheet)

        /// <summary>
        /// Write data into a worksheet starting at (startRow,startColumn) or 'startAddress' (A1).
        /// Bold header row. Optionally clear only the target block. Optionally switch to a kiosk
        /// "range-only" view (hidden Ribbon/UI but keep sheet tabs). No Table/ListObject involved.
        /// </summary>
        internal static string WriteDictionaryToWorksheet(
            IDictionary<string, List<object>> data,
            IList<string> headers,
            IList<string> columnOrder,
            Excel.Workbook wb,
            string worksheetName,
            int startRow,
            int startColumn,
            string startAddress,
            bool saveAfterWrite,
            bool readOnly,
            // ---- Optional controls ----
            bool clearTargetBlockBeforeWrite = true,
            bool applyKioskView = true,     // show clean view
            bool makeFullScreen = false,    // keep normal window frame by default
            bool maximizeWindow = false,    // do NOT maximize by default
            bool bringToFront = true        // bring Excel to foreground after write
        )
        {
            if (data == null || data.Count == 0) return "Dictionary is empty.";
            if (wb == null) return "Workbook is null.";
            if (string.IsNullOrWhiteSpace(worksheetName)) worksheetName = "Sheet1";

            Excel.Worksheet ws = null;
            Excel.Range topLeft = null, bottomRight = null, fullBlock = null;
            Excel.Range rowsFromHeader = null, headerCell = null, lastCellInColumn = null;
            Excel.Range headerRight = null, headerRow = null;
            Excel.Range dataRange = null;

            try
            {
                // Resolve start address if provided
                if (!string.IsNullOrWhiteSpace(startAddress) &&
                    TryParseA1Address(startAddress, out int r, out int c))
                {
                    startRow = r;
                    startColumn = c;
                }

                ws = GetOrCreateWorksheet(wb, worksheetName);
                if (ws == null) return "Failed to access worksheet.";

                // Build column keys respecting explicit order first
                var columnKeys = new List<string>();
                if (columnOrder != null && columnOrder.Count > 0)
                    columnKeys.AddRange(columnOrder);
                foreach (var k in data.Keys)
                    if (!columnKeys.Contains(k)) columnKeys.Add(k);

                int colCount = columnKeys.Count;
                if (colCount == 0) return "Dictionary is empty.";

                // Max rows among branches
                int maxRows = 0;
                foreach (var k in columnKeys)
                    if (data.TryGetValue(k, out var lst) && lst != null && lst.Count > maxRows)
                        maxRows = lst.Count;

                // Total rows = header (1) + data
                int totalRows = Math.Max(1, maxRows + 1);

                // Prepare 2D array [rows, cols]
                var values = new object[totalRows, colCount];

                for (int c2 = 0; c2 < colCount; c2++)
                {
                    string key = columnKeys[c2] ?? string.Empty;
                    string headerLabel = (headers != null && c2 < headers.Count && !string.IsNullOrWhiteSpace(headers[c2]))
                        ? headers[c2]
                        : key;

                    values[0, c2] = headerLabel ?? string.Empty;

                    if (!data.TryGetValue(key, out var branch) || branch == null) continue;
                    for (int r2 = 0; r2 < branch.Count; r2++)
                        values[r2 + 1, c2] = branch[r2];
                }

                // Target block: exactly header + data (focus area)
                startRow = Math.Max(1, startRow);
                startColumn = Math.Max(1, startColumn);

                // Remove all existing rows starting from the header row before writing new data
                try
                {
                    headerCell = (Excel.Range)ws.Cells[startRow, 1];
                    lastCellInColumn = (Excel.Range)ws.Cells[ws.Rows.Count, 1];
                    rowsFromHeader = ws.Range[headerCell, lastCellInColumn].EntireRow;
                    rowsFromHeader.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }
                catch { /* ignore */ }

                topLeft = (Excel.Range)ws.Cells[startRow, startColumn];
                bottomRight = (Excel.Range)ws.Cells[startRow + totalRows - 1, startColumn + colCount - 1];
                fullBlock = ws.Range[topLeft, bottomRight];

                // Optional: clear ONLY the target block
                if (clearTargetBlockBeforeWrite)
                {
                    try { fullBlock.Clear(); } catch { /* ignore */ }
                }

                // Write values
                fullBlock.Value2 = values;

                // Header formatting (bold + light blue)
                try
                {
                    headerRight = (Excel.Range)ws.Cells[startRow, startColumn + colCount - 1];
                    headerRow = ws.Range[topLeft, headerRight];

                    headerRow.Font.Bold = true;
                    headerRow.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    headerRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(
                        System.Drawing.Color.FromArgb(221, 235, 247) // Excel light blue (#DDEBF7)
                    );
                }
                finally
                {
                    ReleaseCom(headerRow);
                    ReleaseCom(headerRight);
                }

                // Ensure ONLY header is formatted → clear formats on data rows
                if (totalRows > 1)
                {
                    dataRange = ws.Range[ws.Cells[startRow + 1, startColumn], bottomRight];
                    try { dataRange.ClearFormats(); } catch { /* ignore */ }
                }

                // Activate sheet; optional maximize; optional bring-to-front
                try
                {
                    ws.Activate();
                    if (maximizeWindow) MaximizeExcelWindow(ws.Application); // opt-in
                    try { ws.Application.Goto(topLeft, true); } catch { }
                    try { topLeft.Select(); } catch { }
                    if (bringToFront) BringExcelToFront(ws.Application);
                }
                catch { }

                // Kiosk view without table (keep sheet tabs; optional fullscreen)
                if (applyKioskView)
                {
                    try
                    {
                        ApplyKioskViewNoTable(ws.Application, ws, fullBlock, makeFullScreen);
                    }
                    catch { /* degrade gracefully */ }
                }

                if (saveAfterWrite && !readOnly)
                {
                    try { wb.Save(); } catch { }
                }

                string wsName = ws?.Name ?? worksheetName;
                string startLabel = !string.IsNullOrWhiteSpace(startAddress)
                    ? startAddress.ToUpperInvariant()
                    : ColumnNumberToLetters(startColumn) + Math.Max(1, startRow).ToString(CultureInfo.InvariantCulture);

                return string.Format(CultureInfo.InvariantCulture,
                    "Wrote {0} columns × {1} rows to '{2}' starting at {3}; header bolded (kiosk: {4}, fullscreen: {5}, maximized: {6}, front: {7}).",
                    colCount,
                    totalRows,
                    wsName,
                    startLabel,
                    applyKioskView ? "on" : "off",
                    makeFullScreen ? "on" : "off",
                    maximizeWindow ? "on" : "off",
                    bringToFront ? "on" : "off");
            }
            catch (Exception ex)
            {
                return "Failed: " + ex.Message;
            }
            finally
            {
                ReleaseCom(dataRange);
                ReleaseCom(rowsFromHeader);
                ReleaseCom(lastCellInColumn);
                ReleaseCom(headerCell);
                ReleaseCom(fullBlock);
                ReleaseCom(topLeft);
                ReleaseCom(bottomRight);
                ReleaseCom(ws);
            }
        }

        #endregion
    }
}
