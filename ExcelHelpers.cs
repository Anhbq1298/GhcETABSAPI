using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
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

                // Resolve and validate path
                string path = ProjectRelative(filePathOrRelative);
                if (string.IsNullOrWhiteSpace(path)) return false;
                string fullPath = Path.GetFullPath(path);

                // 0) Try to attach to any running Excel instances
                var runningApps = GetRunningExcelApplications();

                Excel.Application matchedApp = null;
                Excel.Workbook matchedWorkbook = null;

                // 1) If Excel is running, see if the target workbook is already open there
                if (runningApps != null && runningApps.Count > 0)
                {
                    foreach (var existingApp in runningApps)
                    {
                        if (existingApp == null) continue;

                        Excel.Workbooks openWorkbooks = null;
                        try
                        {
                            openWorkbooks = existingApp.Workbooks;
                            foreach (Excel.Workbook candidate in openWorkbooks)
                            {
                                if (candidate == null) continue;

                                bool matched = false;
                                string candidatePath = null;
                                try { candidatePath = candidate.FullName; } catch { candidatePath = null; }
                                if (!string.IsNullOrWhiteSpace(candidatePath))
                                {
                                    string candidateFullPath;
                                    try { candidateFullPath = Path.GetFullPath(candidatePath); }
                                    catch { candidateFullPath = candidatePath; }

                                    matched = string.Equals(candidateFullPath, fullPath, StringComparison.OrdinalIgnoreCase);
                                }

                                if (matched)
                                {
                                    matchedApp = existingApp;
                                    matchedWorkbook = candidate;
                                    break;
                                }

                                ReleaseCom(candidate);
                            }
                        }
                        catch { /* ignore and continue */ }
                        finally
                        {
                            ReleaseCom(openWorkbooks);
                        }

                        if (matchedWorkbook != null)
                        {
                            break;
                        }
                    }
                }

                if (matchedWorkbook != null && matchedApp != null)
                {
                    app = matchedApp;
                    wb = matchedWorkbook;

                    try
                    {
                        if (visible) { app.Visible = true; app.UserControl = true; }
                        wb.Activate();
                    }
                    catch { /* no-op */ }

                    ReleaseOtherExcelApplications(runningApps, matchedApp);
                    return false; // did not create a new Excel application
                }

                // 2) No open copy found → reuse running Excel or create a new instance
                if (runningApps != null && runningApps.Count > 0)
                {
                    app = runningApps[0];
                    ReleaseOtherExcelApplications(runningApps, app);
                }
                else
                {
                    app = new Excel.Application();
                    createdApplication = true;
                }

                // 3) Open the workbook (avoid prompts & read-only recommendation)
                //    NOTE: if the file doesn’t exist, this will throw; handle as you prefer.
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
                    // Optional fallback: if not explicitly readOnly and open failed (e.g., locked),
                    // try opening as read-only once.
                    if (!readOnly)
                    {
                        try
                        {
                            wb = app.Workbooks.Open(
                                Filename: fullPath,
                                UpdateLinks: 0,
                                ReadOnly: true,
                                IgnoreReadOnlyRecommended: true,
                                AddToMru: false
                            );
                        }
                        catch
                        {
                            // Clean up partially created app if we made it
                            try { if (createdApplication) app.Quit(); } catch { }
                            try { ReleaseCom(wb); } catch { }
                            try { ReleaseCom(app); } catch { }
                            wb = null; app = null;
                            return false;
                        }
                    }
                    else
                    {
                        try { if (createdApplication) app.Quit(); } catch { }
                        try { ReleaseCom(wb); } catch { }
                        try { ReleaseCom(app); } catch { }
                        wb = null; app = null;
                        return false;
                    }
                }

                // 4) Final UI touches
                try
                {
                    if (visible) { app.Visible = true; app.UserControl = true; }
                    wb.Activate();
                }
                catch { /* no-op */ }

                return createdApplication; // true if we spawned a new Excel.exe, else false
            }
        }

        private static List<Excel.Application> GetRunningExcelApplications()
        {
            List<Excel.Application> results = new List<Excel.Application>();
            Exception capturedError = null;

            IRunningObjectTable rot = null;
            IEnumMoniker enumerator = null;
            IBindCtx bindCtx = null;

            try
            {
                Marshal.ThrowExceptionForHR(GetRunningObjectTable(0, out rot));
                Marshal.ThrowExceptionForHR(CreateBindCtx(0, out bindCtx));

                rot.EnumRunning(out enumerator);
                var monikers = new IMoniker[1];

                while (enumerator != null && enumerator.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    var moniker = monikers[0];
                    monikers[0] = null;

                    if (moniker == null)
                        continue;

                    try
                    {
                        string displayName = null;
                        try
                        {
                            moniker.GetDisplayName(bindCtx, null, out displayName);
                        }
                        catch (Exception ex)
                        {
                            capturedError = ex;
                            displayName = null;
                        }

                        if (string.IsNullOrWhiteSpace(displayName) ||
                            displayName.IndexOf("Excel.Application", StringComparison.OrdinalIgnoreCase) < 0)
                        {
                            continue;
                        }

                        object candidate = null;
                        try
                        {
                            rot.GetObject(moniker, out candidate);
                        }
                        catch (Exception ex)
                        {
                            capturedError = ex;
                            candidate = null;
                        }

                        if (candidate is Excel.Application excelApp)
                        {
                            if (!results.Contains(excelApp))
                                results.Add(excelApp);
                        }
                        else
                        {
                            ReleaseCom(candidate);
                        }
                    }
                    finally
                    {
                        ReleaseCom(moniker);
                    }
                }
            }
            catch (Exception ex)
            {
                capturedError = ex;
            }
            finally
            {
                ReleaseCom(enumerator);
                ReleaseCom(rot);
                ReleaseCom(bindCtx);
            }

            if (results.Count == 0 && capturedError != null)
            {
                System.Diagnostics.Debug.WriteLine($"Unable to attach to running Excel instance: {capturedError}");
            }

            return results;
        }

        [DllImport("ole32.dll")]
        private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable pprot);

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        private static void ReleaseOtherExcelApplications(List<Excel.Application> apps, Excel.Application keep)
        {
            if (apps == null) return;

            foreach (var instance in apps)
            {
                if (instance == null) continue;
                if (keep != null && ReferenceEquals(instance, keep)) continue;
                ReleaseCom(instance);
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
