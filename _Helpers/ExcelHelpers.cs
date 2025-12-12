using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

// Optional attach-to-running WITHOUT Marshal.GetActiveObject
using Microsoft.VisualBasic;

namespace MGT
{
    /// <summary>
    /// Excel COM helpers for .NET 8 (net8.0-windows) WITHOUT PIAs
    /// (no Microsoft.Office.Interop.Excel, no office.dll).
    ///
    /// - Late-binding via ProgID "Excel.Application"
    /// - STA-threaded COM execution
    /// - ReleaseCom(FinalReleaseComObject)
    /// - ExcelSheetProfile + ReadSheet (fast Value2 block read)
    /// - WriteDictionaryToWorksheet (fast 2D Value2 write)
    ///
    /// Notes:
    /// - Requires Microsoft Excel installed.
    /// - DO NOT reference Microsoft.Office.Interop.Excel in csproj.
    /// </summary>
    internal static class ExcelHelpers
    {
        private static readonly object _excelLock = new object();

        // =========================
        // UI STATE
        // =========================
        internal sealed class UiState
        {
            public bool? DisplayFullScreen;
            public bool? DisplayFormulaBar;
            public bool? DisplayStatusBar;
            public bool? ScreenUpdating;
            public bool? EnableEvents;

            public bool? DisplayGridlines;
            public bool? DisplayHeadings;
            public bool? DisplayHorizontalScrollBar;
            public bool? DisplayVerticalScrollBar;
            public bool? DisplayWorkbookTabs;
            public int? Zoom;
        }

        // =========================
        // PROFILES (ReadSheet)
        // =========================
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

        // =========================
        // PATH / VALUE HELPERS
        // =========================
        private static string NormalizeAbsoluteWorkbookPath(string absolutePath)
        {
            if (string.IsNullOrWhiteSpace(absolutePath)) return null;

            string p = absolutePath.Trim();
            if (!Path.IsPathRooted(p))
                throw new ArgumentException("Expected an absolute path (from Rhino).", nameof(absolutePath));

            return Path.GetFullPath(p);
        }

        internal static string TrimOrEmpty(object value)
        {
            if (value == null) return string.Empty;
            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(s) ? string.Empty : s.Trim();
        }

        internal static bool IsNullOrEmpty(object value)
        {
            if (value == null) return true;
            if (value is string s) return string.IsNullOrWhiteSpace(s);
            if (value is double d) return double.IsNaN(d) || double.IsInfinity(d);
            return false;
        }

        // =========================
        // COM CLEANUP
        // =========================
        internal static void ReleaseCom(object o)
        {
            if (o == null) return;
            try
            {
                if (Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch { }
        }

        // =========================
        // STA RUNNER
        // =========================
        private static T RunSta<T>(Func<T> work)
        {
            if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
                return work();

            T result = default;
            Exception error = null;

            var t = new Thread(() =>
            {
                try { result = work(); }
                catch (Exception ex) { error = ex; }
            });

            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            if (error != null) throw error;
            return result;
        }

        // =========================
        // EXCEL CREATE / ATTACH (NO Marshal.GetActiveObject)
        // =========================

        /// <summary>
        /// Attach to running Excel WITHOUT Marshal.GetActiveObject.
        /// Uses VB Interaction.GetObject (ROT under the hood).
        /// Returns false if no running instance.
        /// </summary>
        private static bool TryGetRunningExcel(out dynamic app)
        {
            app = null;
            try
            {
                object obj = Interaction.GetObject(string.Empty, "Excel.Application");
                if (obj == null) return false;
                app = obj;
                return true;
            }
            catch
            {
                app = null;
                return false;
            }
        }

        private static dynamic CreateExcelApplication(bool visible)
        {
            Type t = Type.GetTypeFromProgID("Excel.Application", throwOnError: true);
            dynamic app = Activator.CreateInstance(t);

            try { app.DisplayAlerts = false; } catch { }
            try { app.Visible = visible; } catch { }
            try { app.UserControl = visible; } catch { }

            return app;
        }

        /// <summary>
        /// Either attach to running Excel (optional), or create a new instance.
        /// Returns true if a NEW instance was created.
        /// </summary>
        private static bool GetOrCreateExcelApplication(bool visible, bool tryAttachToRunning, out dynamic app)
        {
            app = null;

            if (tryAttachToRunning)
            {
                if (TryGetRunningExcel(out app) && app != null)
                {
                    try { app.DisplayAlerts = false; } catch { }
                    if (visible)
                    {
                        try { app.Visible = true; } catch { }
                        try { app.UserControl = true; } catch { }
                    }
                    return false; // attached (not created)
                }
            }

            app = CreateExcelApplication(visible);
            return true; // created new
        }

        // =========================
        // WORKBOOK OPEN / REUSE
        // =========================
        private static dynamic TryFindOpenWorkbook(dynamic app, string fullPath)
        {
            if (app == null) return null;
            if (string.IsNullOrWhiteSpace(fullPath)) return null;

            dynamic books = null;
            try
            {
                books = app.Workbooks;
                int count = 0;
                try { count = (int)books.Count; } catch { count = 0; }

                string target = Path.GetFullPath(fullPath);

                for (int i = 1; i <= count; i++)
                {
                    dynamic wb = null;
                    bool isMatch = false;

                    try
                    {
                        wb = books[i];

                        string opened = null;
                        try { opened = (string)wb.FullName; } catch { opened = null; }
                        if (string.IsNullOrWhiteSpace(opened))
                        {
                            try { opened = (string)wb.Name; } catch { opened = null; }
                        }

                        if (!string.IsNullOrWhiteSpace(opened))
                        {
                            if (Path.IsPathRooted(opened))
                            {
                                if (string.Equals(Path.GetFullPath(opened), target, StringComparison.OrdinalIgnoreCase))
                                    isMatch = true;
                            }
                            else
                            {
                                if (string.Equals(opened, Path.GetFileName(target), StringComparison.OrdinalIgnoreCase))
                                    isMatch = true;
                            }
                        }

                        if (isMatch) return wb; // caller owns
                    }
                    finally
                    {
                        if (!isMatch) ReleaseCom(wb);
                    }
                }

                return null;
            }
            finally
            {
                ReleaseCom(books);
            }
        }

        /// <summary>
        /// Open workbook in this Excel instance.
        /// If already open -> reuse.
        /// </summary>
        private static dynamic OpenOrReuseWorkbook(dynamic app, string fullPath, bool readOnly)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));

            if (!string.IsNullOrWhiteSpace(fullPath))
            {
                if (!File.Exists(fullPath))
                    throw new FileNotFoundException(fullPath);

                dynamic reuse = TryFindOpenWorkbook(app, fullPath);
                if (reuse != null) return reuse;
            }

            dynamic books = null;
            try
            {
                books = app.Workbooks;

                if (string.IsNullOrWhiteSpace(fullPath))
                    return books.Add();

                return books.Open(
                    fullPath,
                    0,
                    readOnly,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    true,
                    Type.Missing,
                    Type.Missing,
                    false,
                    false,
                    Type.Missing,
                    false,
                    Type.Missing,
                    Type.Missing);
            }
            finally
            {
                ReleaseCom(books);
            }
        }

        // =========================
        // WORKSHEET FIND / CREATE
        // =========================

        /// <summary>
        /// IMPORTANT: internal (fix CS0122)
        /// Returns COM worksheet (caller owns).
        /// </summary>
        internal static dynamic FindWorksheet(dynamic wb, string sheetName)
        {
            if (wb == null) return null;
            if (string.IsNullOrWhiteSpace(sheetName)) return null;

            dynamic sheets = null;
            try
            {
                sheets = wb.Worksheets;
                int count = 0;
                try { count = (int)sheets.Count; } catch { count = 0; }

                for (int i = 1; i <= count; i++)
                {
                    dynamic ws = null;
                    bool isMatch = false;

                    try
                    {
                        ws = sheets[i];
                        string name = null;
                        try { name = (string)ws.Name; } catch { name = null; }

                        if (!string.IsNullOrWhiteSpace(name) &&
                            string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            isMatch = true;
                            return ws;
                        }
                    }
                    finally
                    {
                        if (!isMatch) ReleaseCom(ws);
                    }
                }

                return null;
            }
            finally
            {
                ReleaseCom(sheets);
            }
        }

        private static dynamic GetOrCreateWorksheet(dynamic wb, string sheetName)
        {
            if (wb == null) throw new ArgumentNullException(nameof(wb));
            if (string.IsNullOrWhiteSpace(sheetName)) sheetName = "Sheet1";

            // try find first
            dynamic found = FindWorksheet(wb, sheetName);
            if (found != null) return found;

            // create new
            dynamic sheets = null;
            dynamic afterSheet = null;
            dynamic created = null;

            try
            {
                sheets = wb.Worksheets;

                int count = 0;
                try { count = (int)sheets.Count; } catch { count = 0; }

                if (count >= 1)
                {
                    afterSheet = sheets[count];
                    created = sheets.Add(After: afterSheet);
                }
                else
                {
                    created = sheets.Add();
                }

                try { created.Name = sheetName; } catch { }
                return created;
            }
            finally
            {
                ReleaseCom(afterSheet);
                ReleaseCom(sheets);
            }
        }

        // =========================
        // SAFE QUIT (ONLY IF YOU OWN THE INSTANCE)
        // =========================
        internal static void SafeQuitExcel(dynamic app, dynamic wb)
        {
            try { if (app != null) app.DisplayAlerts = false; } catch { }
            try { if (app != null) app.ScreenUpdating = false; } catch { }
            try { if (app != null) app.UserControl = false; } catch { }

            try { if (wb != null) wb.Close(false); } catch { }
            try { if (app != null) app.Quit(); } catch { }

            ReleaseCom(wb);
            ReleaseCom(app);

            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        // =========================
        // ADDRESS HELPERS
        // =========================
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

            if (!int.TryParse(s.Substring(start, i - start), NumberStyles.None, CultureInfo.InvariantCulture, out row))
                return false;

            return row > 0 && col > 0;
        }

        private static string ColumnNumberToLetters(int column)
        {
            if (column < 1) column = 1;

            var sb = new StringBuilder();
            int dividend = column;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                sb.Insert(0, (char)('A' + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return sb.Length > 0 ? sb.ToString() : "A";
        }

        // =========================
        // UI STRING HELPERS
        // =========================
        private static string FormatProgress(int current, int total, string label, string unit)
        {
            if (total <= 0) return $"{label}: {current} {unit}";
            return $"{label}: {current}/{total} {unit}";
        }

        private static string FormatDone(int count, string label, string unit)
        {
            return $"{label}: {count} {unit}";
        }

        // =========================
        // KIOSK VIEW (NO TABLE)
        // =========================
        internal static UiState ApplyKioskViewNoTable(
            dynamic app,
            dynamic ws,
            dynamic dataRange,
            bool makeFullScreen = false,
            int fallbackZoom = 120)
        {
            if (app == null || ws == null || dataRange == null) return null;

            try { dataRange.Select(); } catch { }

            var state = new UiState();

            dynamic win = null;
            try
            {
                try { state.DisplayFullScreen = (bool)app.DisplayFullScreen; } catch { }
                try { state.DisplayFormulaBar = (bool)app.DisplayFormulaBar; } catch { }
                try { state.DisplayStatusBar = (bool)app.DisplayStatusBar; } catch { }
                try { state.ScreenUpdating = (bool)app.ScreenUpdating; } catch { }
                try { state.EnableEvents = (bool)app.EnableEvents; } catch { }

                try { win = app.ActiveWindow; } catch { win = null; }
                if (win != null)
                {
                    try { state.DisplayGridlines = (bool)win.DisplayGridlines; } catch { }
                    try { state.DisplayHeadings = (bool)win.DisplayHeadings; } catch { }
                    try { state.DisplayHorizontalScrollBar = (bool)win.DisplayHorizontalScrollBar; } catch { }
                    try { state.DisplayVerticalScrollBar = (bool)win.DisplayVerticalScrollBar; } catch { }
                    try { state.DisplayWorkbookTabs = (bool)win.DisplayWorkbookTabs; } catch { }
                    try { state.Zoom = (int)win.Zoom; } catch { }
                }
            }
            catch { }
            finally
            {
                ReleaseCom(win);
            }

            try
            {
                try { app.ScreenUpdating = true; } catch { }
                try { app.EnableEvents = false; } catch { }

                try { app.ExecuteExcel4Macro(@"SHOW.TOOLBAR(""Ribbon"",False)"); } catch { }

                try { app.DisplayFormulaBar = false; } catch { }
                try { app.DisplayStatusBar = false; } catch { }

                dynamic w2 = null;
                try
                {
                    try { w2 = app.ActiveWindow; } catch { w2 = null; }
                    if (w2 != null)
                    {
                        try { w2.DisplayGridlines = false; } catch { }
                        try { w2.DisplayHeadings = false; } catch { }
                        try { w2.DisplayHorizontalScrollBar = false; } catch { }
                        try { w2.DisplayVerticalScrollBar = false; } catch { }
                        try { w2.DisplayWorkbookTabs = true; } catch { }
                        try { w2.Zoom = fallbackZoom; } catch { }
                    }
                }
                finally
                {
                    ReleaseCom(w2);
                }

                try { app.DisplayFullScreen = makeFullScreen; } catch { }
            }
            catch { }
            finally
            {
                try { app.EnableEvents = true; } catch { }
            }

            return state;
        }

        // =========================
        // WINDOW HELPERS
        // =========================
        private const int SW_RESTORE = 9;
        private const int SW_MAXIMIZE = 3;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static void BringExcelToFront(dynamic app)
        {
            if (app == null) return;

            try
            {
                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = new IntPtr((int)app.Hwnd); } catch { hwnd = IntPtr.Zero; }

                if (hwnd != IntPtr.Zero)
                {
                    ShowWindow(hwnd, SW_RESTORE);
                    SetForegroundWindow(hwnd);
                }
            }
            catch { }
        }

        private static void MaximizeExcelWindow(dynamic app)
        {
            if (app == null) return;

            try
            {
                try { app.WindowState = -4137; } catch { } // xlMaximized

                IntPtr hwnd = IntPtr.Zero;
                try { hwnd = new IntPtr((int)app.Hwnd); } catch { hwnd = IntPtr.Zero; }

                if (hwnd != IntPtr.Zero)
                {
                    ShowWindow(hwnd, SW_RESTORE);
                    ShowWindow(hwnd, SW_MAXIMIZE);
                    SetForegroundWindow(hwnd);
                }
            }
            catch { }
        }

        // =========================
        // PUBLIC: AttachOrOpenWorkbook  (FIX CS1628)
        // =========================
        /// <summary>
        /// Returns true if a NEW Excel instance was created.
        /// If tryAttachToRunningExcel=true, may attach via Interaction.GetObject (NO Marshal).
        /// </summary>
        internal static bool AttachOrOpenWorkbook(
            out dynamic app,
            out dynamic wb,
            string absolutePathFromRhino,
            bool visible = true,
            bool readOnly = false,
            bool maximizeWindow = false,
            bool bringToFront = true,
            bool tryAttachToRunningExcel = false)
        {
            // MUST NOT assign out params inside lambda -> use locals (fix CS1628)
            dynamic appLocal = null;
            dynamic wbLocal = null;

            bool createdNewLocal = RunSta(() =>
            {
                lock (_excelLock)
                {
                    string fullPath = null;
                    if (!string.IsNullOrWhiteSpace(absolutePathFromRhino))
                        fullPath = NormalizeAbsoluteWorkbookPath(absolutePathFromRhino);

                    dynamic a = null;
                    dynamic w = null;
                    bool createdNew = false;

                    try
                    {
                        createdNew = GetOrCreateExcelApplication(visible, tryAttachToRunningExcel, out a);
                        w = OpenOrReuseWorkbook(a, fullPath, readOnly);

                        if (visible)
                        {
                            try { a.Visible = true; } catch { }
                            try { a.UserControl = true; } catch { }

                            if (maximizeWindow) MaximizeExcelWindow(a);
                            if (bringToFront) BringExcelToFront(a);
                            try { w.Activate(); } catch { }
                        }

                        appLocal = a;
                        wbLocal = w;
                        return createdNew;
                    }
                    catch
                    {
                        // If we created it -> quit. If attached -> DO NOT quit Excel.
                        if (createdNew)
                        {
                            SafeQuitExcel(a, w);
                        }
                        else
                        {
                            ReleaseCom(w);
                            ReleaseCom(a);
                        }

                        appLocal = null;
                        wbLocal = null;
                        return false;
                    }
                }
            });

            app = appLocal;
            wb = wbLocal;
            return createdNewLocal;
        }

        // =========================
        // PUBLIC: ReadSheet
        // =========================
        internal static ExcelSheetReadResult ReadSheet(
            string absolutePathFromRhino,
            string requestedSheetName,
            ExcelSheetProfile profile,
            Action<int, int, string> progressCallback = null)
        {
            if (string.IsNullOrWhiteSpace(absolutePathFromRhino))
                throw new ArgumentException("Excel path cannot be empty.", nameof(absolutePathFromRhino));

            if (profile == null) throw new ArgumentNullException(nameof(profile));
            if (profile.ExpectedHeaders == null)
                throw new ArgumentException("Profile must define ExpectedHeaders.", nameof(profile));

            return RunSta(() =>
            {
                lock (_excelLock)
                {
                    string fullPath = NormalizeAbsoluteWorkbookPath(absolutePathFromRhino);

                    string sheetToRead = string.IsNullOrWhiteSpace(requestedSheetName)
                        ? profile.ExpectedSheetName
                        : requestedSheetName;

                    if (!string.IsNullOrEmpty(profile.ExpectedSheetName) &&
                        !string.Equals(sheetToRead, profile.ExpectedSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            $"Invalid workbook: expected sheet name '{profile.ExpectedSheetName}'.");
                    }

                    dynamic app = null;
                    dynamic wb = null;
                    dynamic ws = null;
                    dynamic usedRange = null;

                    try
                    {
                        app = CreateExcelApplication(visible: false);
                        wb = OpenOrReuseWorkbook(app, fullPath, readOnly: true);

                        // strict exists:
                        ws = FindWorksheet(wb, sheetToRead);
                        if (ws == null)
                            throw new InvalidOperationException($"Sheet '{sheetToRead}' not found.");

                        var result = new ExcelSheetReadResult();

                        int startCol = Math.Max(1, profile.StartColumn);
                        int colCount = profile.ExpectedHeaders.Count;

                        // headers row
                        dynamic hTL = null, hBR = null, hRng = null;
                        try
                        {
                            hTL = ws.Cells[1, startCol];
                            hBR = ws.Cells[1, startCol + colCount - 1];
                            hRng = ws.Range(hTL, hBR);

                            object hv = null;
                            try { hv = hRng.Value2; } catch { hv = null; }

                            if (hv is object[,] hv2)
                            {
                                for (int c = 1; c <= colCount; c++)
                                {
                                    string headerValue = TrimOrEmpty(hv2[1, c]);
                                    result.Headers.Add(headerValue);

                                    string expected = profile.ExpectedHeaders[c - 1];
                                    if (!string.Equals(headerValue, expected, StringComparison.OrdinalIgnoreCase))
                                    {
                                        string colLetter = ColumnNumberToLetters(startCol + (c - 1));
                                        throw new InvalidOperationException(
                                            $"Invalid workbook: expected header '{expected}' in column {colLetter}, found '{headerValue}'.");
                                    }
                                }
                            }
                            else
                            {
                                string headerValue = TrimOrEmpty(hv);
                                result.Headers.Add(headerValue);

                                string expected = profile.ExpectedHeaders[0];
                                if (!string.Equals(headerValue, expected, StringComparison.OrdinalIgnoreCase))
                                {
                                    string colLetter = ColumnNumberToLetters(startCol);
                                    throw new InvalidOperationException(
                                        $"Invalid workbook: expected header '{expected}' in column {colLetter}, found '{headerValue}'.");
                                }
                            }
                        }
                        finally
                        {
                            ReleaseCom(hRng);
                            ReleaseCom(hTL);
                            ReleaseCom(hBR);
                        }

                        // last row (UsedRange)
                        int lastRow = 1;
                        try
                        {
                            usedRange = ws.UsedRange;
                            if (usedRange != null)
                            {
                                int firstRow = 1;
                                int rowsCount = 1;
                                try { firstRow = (int)usedRange.Row; } catch { firstRow = 1; }
                                try { rowsCount = (int)usedRange.Rows.Count; } catch { rowsCount = 1; }
                                lastRow = Math.Max(1, firstRow + rowsCount - 1);
                            }
                        }
                        catch { lastRow = 1; }

                        int totalDataRows = Math.Max(0, lastRow - 1);
                        progressCallback?.Invoke(0, totalDataRows, FormatProgress(0, totalDataRows, profile.ProgressLabel, profile.ProgressUnit));

                        if (lastRow <= 1)
                        {
                            progressCallback?.Invoke(0, 0, FormatDone(0, profile.CompletionLabel, profile.ProgressUnit));
                            return result;
                        }

                        // data block rows 2..lastRow
                        dynamic dTL = null, dBR = null, dRng = null;
                        try
                        {
                            dTL = ws.Cells[2, startCol];
                            dBR = ws.Cells[lastRow, startCol + colCount - 1];
                            dRng = ws.Range(dTL, dBR);

                            object dv = null;
                            try { dv = dRng.Value2; } catch { dv = null; }

                            if (dv is object[,] dv2)
                            {
                                int rows = dv2.GetLength(0);
                                for (int r = 1; r <= rows; r++)
                                {
                                    var rowArr = new object[colCount];
                                    bool hasData = false;

                                    for (int c = 1; c <= colCount; c++)
                                    {
                                        object v = dv2[r, c];
                                        rowArr[c - 1] = v;
                                        if (!IsNullOrEmpty(v)) hasData = true;
                                    }

                                    progressCallback?.Invoke(
                                        r,
                                        totalDataRows,
                                        FormatProgress(r, totalDataRows, profile.ProgressLabel, profile.ProgressUnit));

                                    if (hasData) result.Rows.Add(rowArr);
                                }
                            }
                            else
                            {
                                var rowArr = new object[colCount];
                                rowArr[0] = dv;
                                if (!IsNullOrEmpty(dv)) result.Rows.Add(rowArr);
                            }
                        }
                        finally
                        {
                            ReleaseCom(dRng);
                            ReleaseCom(dTL);
                            ReleaseCom(dBR);
                        }

                        progressCallback?.Invoke(
                            result.Rows.Count,
                            result.Rows.Count,
                            FormatDone(result.Rows.Count, profile.CompletionLabel, profile.ProgressUnit));

                        return result;
                    }
                    finally
                    {
                        ReleaseCom(usedRange);
                        ReleaseCom(ws);
                        SafeQuitExcel(app, wb); // ReadSheet always owns hidden instance
                    }
                }
            });
        }

        // =========================
        // PUBLIC: WriteDictionaryToWorksheet
        // =========================
        internal static string WriteDictionaryToWorksheet(
    IDictionary<string, List<object>> data,
    IList<string> headers,
    IList<string> columnOrder,
    dynamic wb,                 // EXISTING workbook instance (reuse)
    string worksheetName,
    int startRow,
    int startColumn,
    string startAddress,
    bool saveAfterWrite,
    bool readOnly,
    bool clearTargetBlockBeforeWrite = true,
    bool applyKioskView = true,
    bool makeFullScreen = false)
        {
            if (data == null || data.Count == 0) return "Dictionary is empty.";
            if (wb == null) return "Workbook is null.";
            if (string.IsNullOrWhiteSpace(worksheetName)) worksheetName = "Sheet1";

            return RunSta(() =>
            {
                lock (_excelLock)
                {
                    if (!string.IsNullOrWhiteSpace(startAddress) &&
                        TryParseA1Address(startAddress, out int rr, out int cc))
                    {
                        startRow = rr;
                        startColumn = cc;
                    }

                    startRow = Math.Max(1, startRow);
                    startColumn = Math.Max(1, startColumn);

                    dynamic app = null;
                    dynamic ws = null;

                    dynamic topLeft = null;
                    dynamic bottomRight = null;
                    dynamic fullBlock = null;
                    dynamic headerRow = null;
                    dynamic dataRange = null;

                    try
                    {
                        // Get Excel.Application from the existing workbook
                        try { app = wb.Application; } catch { app = null; }

                        // Get/Create sheet in THIS workbook
                        ws = GetOrCreateWorksheet(wb, worksheetName);

                        // Build column keys (explicit order first)
                        var columnKeysLocal = new List<string>();
                        if (columnOrder != null && columnOrder.Count > 0)
                            columnKeysLocal.AddRange(columnOrder);

                        foreach (var k in data.Keys)
                            if (!columnKeysLocal.Contains(k)) columnKeysLocal.Add(k);

                        int colCount = columnKeysLocal.Count;
                        if (colCount == 0) return "Dictionary is empty.";

                        // Max rows among columns
                        int maxRows = 0;
                        foreach (var k in columnKeysLocal)
                        {
                            if (data.TryGetValue(k, out var lst) && lst != null && lst.Count > maxRows)
                                maxRows = lst.Count;
                        }

                        int totalRows = Math.Max(1, maxRows + 1); // header + data

                        // Prepare 2D array for fast Value2 write
                        var values = new object[totalRows, colCount];

                        for (int c2 = 0; c2 < colCount; c2++)
                        {
                            string key = columnKeysLocal[c2] ?? string.Empty;

                            string headerLabel = (headers != null &&
                                                  c2 < headers.Count &&
                                                  !string.IsNullOrWhiteSpace(headers[c2]))
                                ? headers[c2]
                                : key;

                            values[0, c2] = headerLabel ?? string.Empty;

                            if (!data.TryGetValue(key, out var branch) || branch == null) continue;
                            for (int r2 = 0; r2 < branch.Count; r2++)
                                values[r2 + 1, c2] = branch[r2];
                        }

                        // Target range
                        topLeft = ws.Cells[startRow, startColumn];
                        bottomRight = ws.Cells[startRow + totalRows - 1, startColumn + colCount - 1];
                        fullBlock = ws.Range(topLeft, bottomRight);

                        if (clearTargetBlockBeforeWrite)
                        {
                            try { fullBlock.Clear(); } catch { }
                        }

                        fullBlock.Value2 = values;

                        // Header formatting (row 1 of block)
                        try
                        {
                            dynamic headerRight = null;
                            try
                            {
                                headerRight = ws.Cells[startRow, startColumn + colCount - 1];
                                headerRow = ws.Range(topLeft, headerRight);

                                try { headerRow.Font.Bold = true; } catch { }
                                try { headerRow.Interior.Pattern = 1; } catch { }       // xlPatternSolid
                                try { headerRow.Interior.Color = 16247773; } catch { }  // light blue
                            }
                            finally
                            {
                                ReleaseCom(headerRight);
                            }
                        }
                        catch { }

                        // Clear formats for data rows only
                        if (totalRows > 1)
                        {
                            try
                            {
                                dataRange = ws.Range(ws.Cells[startRow + 1, startColumn], bottomRight);
                                dataRange.ClearFormats();
                            }
                            catch { }
                        }

                        // Optional kiosk focus (only if we can access app)
                        if (applyKioskView && app != null)
                        {
                            try
                            {
                                ws.Activate();
                                ApplyKioskViewNoTable(app, ws, fullBlock, makeFullScreen, 120);
                            }
                            catch { }
                        }

                        // Save
                        if (saveAfterWrite && !readOnly)
                        {
                            try { wb.Save(); } catch { }
                        }

                        string startLabel = !string.IsNullOrWhiteSpace(startAddress)
                            ? startAddress.ToUpperInvariant()
                            : ColumnNumberToLetters(startColumn) + startRow.ToString(CultureInfo.InvariantCulture);

                        return string.Format(CultureInfo.InvariantCulture,
                            "Wrote {0} columns × {1} rows to '{2}' at {3} (reused workbook).",
                            colCount, totalRows, worksheetName, startLabel);
                    }
                    catch (Exception ex)
                    {
                        return "Failed: " + ex.Message;
                    }
                    finally
                    {
                        ReleaseCom(dataRange);
                        ReleaseCom(headerRow);
                        ReleaseCom(fullBlock);
                        ReleaseCom(topLeft);
                        ReleaseCom(bottomRight);
                        ReleaseCom(ws);

                        // DO NOT quit Excel here (we don't own wb/app).
                        ReleaseCom(app);

                        try
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                        catch { }
                    }
                }
            });
        }
    }
}
