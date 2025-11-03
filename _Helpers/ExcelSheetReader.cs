using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace MGT
{
    internal static class ExcelSheetReader
    {
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
                app = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false,
                    UserControl = false
                };

                books = app.Workbooks;
                wb = books.Open(
                    Filename: fullPath,
                    UpdateLinks: 0,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    AddToMru: false);

                ws = ExcelHelpers.FindWorksheet(wb, sheetToRead);
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
                        string headerValue = ExcelHelpers.TrimOrEmpty(headerCell?.Value2);
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
                        ExcelHelpers.ReleaseCom(headerCell);
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
                            if (!ExcelHelpers.IsNullOrEmpty(value))
                            {
                                hasData = true;
                            }
                        }
                        finally
                        {
                            ExcelHelpers.ReleaseCom(cell);
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
                ExcelHelpers.ReleaseCom(usedRange);

                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                }

                ExcelHelpers.ReleaseCom(ws);
                ExcelHelpers.ReleaseCom(wb);
                ExcelHelpers.ReleaseCom(books);

                if (app != null)
                {
                    try { app.Quit(); } catch { }
                }

                ExcelHelpers.ReleaseCom(app);
            }
        }
    }
}
