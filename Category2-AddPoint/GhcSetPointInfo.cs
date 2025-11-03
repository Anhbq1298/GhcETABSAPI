// -------------------------------------------------------------
// Component : GhcSetPointInfo
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET 8.0 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)
// Panel     : "MGT" / "2.0 Point Object Modelling"
// -------------------------------------------------------------
// Inputs (ordered):
//   0) run        (bool, item)        – Rising-edge trigger (executes on False→True).
//   1) sapModel   (ETABSv1.cSapModel) – ETABS model handle from Attach component.
//   2) excelPath  (string, item)      – Absolute or project-relative workbook path.
//   3) sheetName  (string, item)      – Worksheet name (default: "Point Info").
//   4) baseline   (tree, optional)    – Baseline tree from GhcGetPointInfo; used
//                                       only to detect safe deletions (reference-only).
//
// Outputs:
//   0) values   (tree) – Echo of Excel data in 4 branches: UniqueName / X / Y / Z.
//   1) actions  (list) – Human-readable log of Create / Move / Delete / Rename operations.
//   2) messages (list) – Summaries, skips, and error diagnostics.
//
// Behavior Notes:
//   • Baseline is reference-only and index-agnostic (row order in Excel does not matter).
//   • CREATE  → Only if UniqueName is not in ETABS AND no existing point occupies the
//               same coordinates (±1e-6). Missing XYZ → skip create.
//   • MOVE    → If UniqueName exists and |Δ|≥tol between current and target XYZ
//               (blank Excel cells keep current coordinate).
//   • RENAME  → (a) Case-only mismatch → ChangeName(oldExact, desiredExcelName).
//               (b) If desired name not found but another point sits at target coords,
//                   ChangeName(thatPointName, desiredExcelName) instead of creating.
//   • DELETE  → If a name exists in baseline but is missing from Excel AND has no
//               connectivity; use PointObj.DeleteSpecialPoint(name).
//   • Reader skips fully blank rows and processes until UsedRange end.
//   • Dual progress bars (Excel + Assignment) always reach 100%.
//   • After apply, this component **auto-refreshes GhcGetPointInfo**.
// -------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using ETABSv1;
using Grasshopper.Kernel;
using Grasshopper.Kernel.Data;
using Grasshopper.Kernel.Types;
using static MGT.ComponentShared;

namespace MGT
{
    public sealed class GhcSetPointInfo : GH_Component
    {
        // ===== Constants =====
        private const string DefaultSheet = "Point Info";   // Default worksheet name
        private const int StartColumn = 2;                  // Column B (A=1)
        private const int StartRow = 2;                     // Data starts at row 2 (row 1 = headers)
        private const double Tolerance = 1e-6;              // Tolerance for move & coordinate signature

        // ===== Sticky replay (non-rising) =====
        private bool _lastRun;
        private GH_Structure<GH_ObjectWrapper> _lastValues = new GH_Structure<GH_ObjectWrapper>();
        private readonly List<string> _lastActions = new();
        private readonly List<string> _lastMessages = new() { "No previous run. Toggle 'run' to execute." };

        public GhcSetPointInfo()
            : base(
                "Set Point Info",
                "SetPointInfo",
                "Reads points (UniqueName, X, Y, Z) from Excel and applies minimal Create / Move / Delete / Rename changes to ETABS.\n" +
                "Baseline is reference-only and index-agnostic; Excel row order does not matter.",
                "MGT",
                "2.0 Point Object Modelling")
        { }

        public override Guid ComponentGuid => new Guid("A9B6F07F-7D5E-4A25-AD2A-6F0A7AE12C47");
        protected override Bitmap Icon => null;

        // ===== Inputs =====
        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("run", "run", "Rising-edge trigger (executes when toggled True).", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel object from Attach component.", GH_ParamAccess.item);
            p.AddTextParameter("excelPath", "excelPath", "Full or project-relative path to the workbook.", GH_ParamAccess.item, string.Empty);
            p.AddTextParameter("sheetName", "sheetName", "Target worksheet name (default: 'Point Info').", GH_ParamAccess.item, DefaultSheet);

            int idx = p.AddGenericParameter(
                "baseline",
                "baseline",
                "Optional baseline tree (reference-only). Used to detect deletions safely.",
                GH_ParamAccess.tree);
            p[idx].Optional = true;
        }

        // ===== Outputs =====
        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddGenericParameter("values", "values", "Echo of Excel data (UniqueName / X / Y / Z).", GH_ParamAccess.tree);
            p.AddTextParameter("actions", "actions", "Logs of Create / Move / Delete / Rename operations in ETABS.", GH_ParamAccess.list);
            p.AddTextParameter("messages", "messages", "Status, warnings, and error details.", GH_ParamAccess.list);
        }

        // ===== Execution =====
        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sap = null;
            string path = null;
            string sheet = DefaultSheet;

            da.GetData(0, ref run);
            da.GetData(1, ref sap);
            da.GetData(2, ref path);
            da.GetData(3, ref sheet);
            da.GetDataTree(4, out GH_Structure<IGH_Goo> baselineTree);

            // Rising-edge gate (execute only on False→True)
            bool rising = !_lastRun && run;
            if (!rising)
            {
                da.SetDataTree(0, _lastValues.Duplicate());
                da.SetDataList(1, _lastActions);
                da.SetDataList(2, _lastMessages);
                _lastRun = run;
                return;
            }

            var actions = new List<string>();
            var messages = new List<string>();
            var valueTree = new GH_Structure<GH_ObjectWrapper>();

            try
            {
                // --- Validate inputs ---
                if (sap == null)
                    throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");

                string fullPath = ExcelHelpers.ProjectRelative(path);
                if (string.IsNullOrWhiteSpace(fullPath))
                    throw new InvalidOperationException("excelPath is empty.");
                if (!File.Exists(fullPath))
                    throw new FileNotFoundException("Excel workbook not found.", fullPath);
                if (string.IsNullOrWhiteSpace(sheet))
                    sheet = DefaultSheet;

                // --- Progress UI: Excel + Assignment ---
                UiHelpers.ShowDualProgressBar(
                    "Set Point Info",
                    "Reading Excel...",
                    0,
                    "Updating points...",
                    0);

                // --- Parse baseline (reference-only) ---
                PointBaseline baseline = PointBaseline.FromTree(baselineTree);

                // --- Read Excel → row list (reader drives its own 100% progress) ---
                List<PointRow> rows = ReadExcelSheet(
                    fullPath,
                    sheet,
                    (cur, total, msg) => UiHelpers.UpdateExcelProgressBar(cur, total, msg));

                // Echo read values to GH output
                valueTree = BuildValueTree(rows);

                // --- Apply to ETABS ---
                if (rows.Count == 0)
                {
                    messages.Add("Excel sheet contained no valid data rows.");
                    UiHelpers.UpdateAssignmentProgressBar(1, 1, "No rows to update (100%).");
                }
                else
                {
                    EnsureModelUnlocked(sap);

                    UiHelpers.UpdateAssignmentProgressBar(
                        0,
                        rows.Count,
                        FormatAssignStatus(0, rows.Count));

                    ProcessRows_CreateMoveDelete(
                        sap,
                        rows,
                        baseline,
                        actions,
                        messages,
                        (c, t, s) => UiHelpers.UpdateAssignmentProgressBar(c, t, s));

                    UiHelpers.UpdateAssignmentProgressBar(rows.Count, rows.Count, FormatAssignStatus(rows.Count, rows.Count));
                }

                // Non-fatal view refresh
                try { sap.View.RefreshView(0, false); } catch { /* ignore */ }

                // === Auto-refresh GhcGetPointInfo after applying updates ===
                int refreshed = TriggerGetPointInfoRefresh();
                if (refreshed > 0)
                    messages.Add($"Scheduled refresh for {refreshed} Get Point Info component(s).");
            }
            catch (Exception ex)
            {
                messages.Add("Error: " + ex.Message);
                AddRuntimeMessage(GH_RuntimeMessageLevel.Error, ex.Message);
            }
            finally
            {
                UiHelpers.CloseProgressBar();

                // Outputs
                da.SetDataTree(0, valueTree);
                da.SetDataList(1, actions);
                da.SetDataList(2, messages);

                // Sticky replay
                _lastValues = valueTree.Duplicate();
                _lastActions.Clear(); _lastActions.AddRange(actions);
                _lastMessages.Clear(); _lastMessages.AddRange(messages);
                _lastRun = run;
            }
        }

        // =========================================================
        // CORE: Compare current Excel rows vs ETABS vs Baseline
        // =========================================================
        private static void ProcessRows_CreateMoveDelete(
            cSapModel sap,
            List<PointRow> rows,
            PointBaseline baseline,
            List<string> actions,
            List<string> messages,
            Action<int, int, string> progress)
        {
            // Build ETABS indices once (names, coordinates, coordinate signatures, and reverse map)
            BuildEtabsPointIndices(
                sap,
                out var nameSet,
                out var coordSigSet,
                out var nameToCoord,
                out var coordSigToName /* reverse index for rename-by-coordinate */);

            var excelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var seenExcelNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            int createCount = 0, moveCount = 0, deleteCount = 0, renameCount = 0;
            int processed = 0, total = rows.Count;

            void Tick() => progress?.Invoke(++processed, total, FormatAssignStatus(processed, total));

            // ===== PASS 1: RENAME(Case) + MOVE + CREATE / RENAME(By-Coordinate) =====
            foreach (var (row, idx) in rows.WithIndex())
            {
                string rowLabel = $"Row {StartRow + idx}";
                string desiredName = row.UniqueName?.Trim();

                if (string.IsNullOrWhiteSpace(desiredName))
                {
                    messages.Add($"{rowLabel}: UniqueName is empty. Skipped.");
                    Tick();
                    continue;
                }

                // Prevent duplicate names within the same Excel import
                if (!seenExcelNames.Add(desiredName))
                {
                    messages.Add($"{rowLabel}: Duplicate UniqueName in Excel ('{desiredName}'). Subsequent occurrence skipped.");
                    Tick();
                    continue;
                }

                excelNames.Add(desiredName);

                // Name already exists? → (A) Case-sync rename, then (B) move
                if (nameSet.Contains(desiredName))
                {
                    // (A) Case-only mismatch → enforce exact casing from Excel
                    if (TryFindExactNameIgnoreCase(nameToCoord, desiredName, out string exactExisting) &&
                        !string.Equals(exactExisting, desiredName, StringComparison.Ordinal))
                    {
                        int rn = sap.PointObj.ChangeName(exactExisting, desiredName);
                        if (rn == 0)
                        {
                            renameCount++;
                            actions.Add($"{rowLabel}: Renamed '{exactExisting}' -> '{desiredName}' (case sync).");

                            if (nameToCoord.TryGetValue(exactExisting, out var curC))
                            {
                                nameToCoord.Remove(exactExisting);
                                nameToCoord[desiredName] = curC;
                                coordSigSet.Add(CoordSig(curC.x, curC.y, curC.z));
                                coordSigToName[CoordSig(curC.x, curC.y, curC.z)] = desiredName;
                            }
                        }
                        else
                        {
                            messages.Add($"{rowLabel}: ChangeName failed (code {rn}).");
                        }
                    }

                    // (B) Move if XYZ changed (blank Excel cell keeps current)
                    if (!nameToCoord.TryGetValue(desiredName, out var cur))
                    {
                        messages.Add($"{rowLabel}: Cannot read current coordinates for '{desiredName}'.");
                        Tick();
                        continue;
                    }

                    double tx = row.X ?? cur.x;
                    double ty = row.Y ?? cur.y;
                    double tz = row.Z ?? cur.z;

                    double dx = tx - cur.x, dy = ty - cur.y, dz = tz - cur.z;
                    if (Math.Abs(dx) < Tolerance && Math.Abs(dy) < Tolerance && Math.Abs(dz) < Tolerance)
                    {
                        Tick();
                        continue; // No move needed
                    }

                    int sel = sap.PointObj.SetSelected(desiredName, true);
                    if (sel != 0)
                    {
                        messages.Add($"{rowLabel}: Failed to select '{desiredName}' (code {sel}).");
                        Tick();
                        continue;
                    }

                    try
                    {
                        int mv = sap.EditGeneral.Move(dx, dy, dz);
                        if (mv == 0)
                        {
                            moveCount++;
                            actions.Add($"{rowLabel}: Moved '{desiredName}' to ({tx:0.###}, {ty:0.###}, {tz:0.###}).");
                            nameToCoord[desiredName] = (tx, ty, tz);
                            coordSigSet.Add(CoordSig(tx, ty, tz));
                            coordSigToName[CoordSig(tx, ty, tz)] = desiredName;
                        }
                        else
                        {
                            messages.Add($"{rowLabel}: Move failed for '{desiredName}' (code {mv}).");
                        }
                    }
                    finally
                    {
                        try { sap.PointObj.SetSelected(desiredName, false); } catch { /* ignore */ }
                    }

                    Tick();
                    continue;
                }

                // Name not present in ETABS
                // If full XYZ provided → try rename-by-coordinate (prefer rename over create)
                if (row.X != null && row.Y != null && row.Z != null)
                {
                    string sig = CoordSig(row.X.Value, row.Y.Value, row.Z.Value);

                    if (coordSigSet.Contains(sig) && coordSigToName.TryGetValue(sig, out string oldNameAtSpot))
                    {
                        if (!string.Equals(oldNameAtSpot, desiredName, StringComparison.Ordinal))
                        {
                            int rn = sap.PointObj.ChangeName(oldNameAtSpot, desiredName);
                            if (rn == 0)
                            {
                                renameCount++;
                                actions.Add($"{rowLabel}: Renamed '{oldNameAtSpot}' -> '{desiredName}' (via coordinate match).");

                                if (nameToCoord.TryGetValue(oldNameAtSpot, out var curC))
                                {
                                    nameToCoord.Remove(oldNameAtSpot);
                                    nameToCoord[desiredName] = curC;
                                    coordSigToName[sig] = desiredName;
                                }
                                nameSet.Add(desiredName);
                                Tick();
                                continue; // Done (rename only)
                            }
                            else
                            {
                                messages.Add($"{rowLabel}: ChangeName failed (code {rn}). Falling back to create.");
                                // fall through to create
                            }
                        }
                    }

                    // No occupant at coords → proceed to CREATE (guard against duplicates again)
                    if (coordSigSet.Contains(sig))
                    {
                        messages.Add($"{rowLabel}: Coordinates already occupied; skipped create.");
                        Tick();
                        continue;
                    }

                    string newName = desiredName;
                    int add = sap.PointObj.AddCartesian(row.X.Value, row.Y.Value, row.Z.Value, ref newName);
                    if (add == 0)
                    {
                        if (!string.Equals(newName, desiredName, StringComparison.Ordinal))
                        {
                            int rn = sap.PointObj.ChangeName(newName, desiredName);
                            if (rn != 0)
                            {
                                messages.Add($"{rowLabel}: ChangeName failed after create (code {rn}). Kept '{newName}'.");
                                desiredName = newName;
                            }
                        }

                        createCount++;
                        actions.Add($"{rowLabel}: Created point '{desiredName}' at ({row.X:0.###}, {row.Y:0.###}, {row.Z:0.###}).");

                        nameSet.Add(desiredName);
                        nameToCoord[desiredName] = (row.X.Value, row.Y.Value, row.Z.Value);
                        coordSigSet.Add(sig);
                        coordSigToName[sig] = desiredName;
                    }
                    else
                    {
                        messages.Add($"{rowLabel}: AddCartesian failed for '{desiredName}' (code {add}).");
                    }

                    Tick();
                    continue;
                }

                // No XYZ for new name → cannot safely create/rename-by-coordinate
                messages.Add($"{rowLabel}: '{desiredName}' not found and XYZ missing. Skipped.");
                Tick();
            }

            // ===== PASS 2: DELETE (baseline names missing from Excel) =====
            if (baseline != null && baseline.HasData)
            {
                foreach (var entry in baseline.Entries)
                {
                    string baseName = entry.UniqueName;
                    if (string.IsNullOrWhiteSpace(baseName)) continue;
                    if (excelNames.Contains(baseName)) continue;
                    if (!TryGetPoint(sap, baseName, out _, out _, out _)) continue;

                    if (!TryCheckConnectivity(sap, baseName, out bool hasConn))
                    {
                        messages.Add($"Unable to determine connectivity for '{baseName}'. Skipped delete.");
                        continue;
                    }
                    if (hasConn)
                    {
                        messages.Add($"Skipped deleting '{baseName}' because it still has connectivity.");
                        continue;
                    }

                    int del = sap.PointObj.DeleteSpecialPoint(baseName);
                    if (del == 0)
                    {
                        deleteCount++;
                        actions.Add($"Deleted point '{baseName}' because it was removed from Excel.");
                    }
                    else
                    {
                        messages.Add($"Failed to delete point '{baseName}' (code {del}).");
                    }
                }
            }

            // Summary
            messages.Add($"Processed {rows.Count} row(s).");
            if (createCount > 0) messages.Add($"Created {createCount} point(s).");
            if (moveCount > 0) messages.Add($"Moved {moveCount} point(s).");
            if (renameCount > 0) messages.Add($"Renamed {renameCount} point(s).");
            if (deleteCount > 0) messages.Add($"Deleted {deleteCount} point(s) removed from Excel.");

            // Final 100%
            progress?.Invoke(total, total, FormatAssignStatus(total, total));
        }

        // =========================================================
        // Excel Reader (header-based)
        // =========================================================
        private static List<PointRow> ReadExcelSheet(
            string workbookPath,
            string sheetName,
            Action<int, int, string> progress)
        {
            var profile = new ExcelHelpers.ExcelSheetProfile
            {
                ExpectedSheetName = sheetName,
                StartColumn = StartColumn,
                ExpectedHeaders = new[] { "UniqueName", "X", "Y", "Z" }
            };

            var result = ExcelHelpers.ReadSheet(workbookPath, sheetName, profile, progress);

            // Build header→index map
            var col = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 0; c < result.Headers.Count; c++)
            {
                string h = Convert.ToString(result.Headers[c], CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(h)) col[h] = c;
            }

            if (!col.TryGetValue("UniqueName", out int iName))
                throw new InvalidOperationException("Missing 'UniqueName' header in sheet.");
            if (!col.TryGetValue("X", out int iX))
                throw new InvalidOperationException("Missing 'X' header in sheet.");
            if (!col.TryGetValue("Y", out int iY))
                throw new InvalidOperationException("Missing 'Y' header in sheet.");
            if (!col.TryGetValue("Z", out int iZ))
                throw new InvalidOperationException("Missing 'Z' header in sheet.");

            var rows = new List<PointRow>(result.Rows.Count);

            foreach (object[] r in result.Rows)
            {
                string name = ReadText(SafeAt(r, iName));
                double? x = ReadNullableDouble(SafeAt(r, iX));
                double? y = ReadNullableDouble(SafeAt(r, iY));
                double? z = ReadNullableDouble(SafeAt(r, iZ));

                bool allEmpty = string.IsNullOrWhiteSpace(name) && x == null && y == null && z == null;
                if (allEmpty) continue;

                rows.Add(new PointRow { UniqueName = name, X = x, Y = y, Z = z });
            }

            return rows;

            static object SafeAt(object[] arr, int i) => (i >= 0 && i < arr.Length) ? arr[i] : null;
        }

        // =========================================================
        // Helpers & Utilities
        // =========================================================

        private static string FormatAssignStatus(int processed, int total)
        {
            if (total <= 0) return "Updating ETABS points 0 of 0 row(s) (100%).";
            int p = Math.Max(0, Math.Min(processed, total));
            double percent = (p * 100.0) / total;
            return $"Updating ETABS points {p} of {total} row(s) ({percent:0.##}%).";
        }

        private static GH_Structure<GH_ObjectWrapper> BuildValueTree(List<PointRow> rows)
        {
            var tree = new GH_Structure<GH_ObjectWrapper>();
            GH_Path pName = new GH_Path(0);
            GH_Path pX = new GH_Path(1);
            GH_Path pY = new GH_Path(2);
            GH_Path pZ = new GH_Path(3);

            foreach (var r in rows)
            {
                tree.Append(new GH_ObjectWrapper(r.UniqueName), pName);
                tree.Append(new GH_ObjectWrapper(r.X), pX);
                tree.Append(new GH_ObjectWrapper(r.Y), pY);
                tree.Append(new GH_ObjectWrapper(r.Z), pZ);
            }
            return tree;
        }

        /// <summary>
        /// Scan ETABS once and build fast lookup indices:
        ///  - nameSet:       case-insensitive set of names
        ///  - coordSigSet:   set of rounded XYZ signatures
        ///  - nameToCoord:   exact coordinates by exact name (key preserves casing)
        ///  - coordSigToName:reverse index: signature → representative name
        /// </summary>
        private static void BuildEtabsPointIndices(
            cSapModel sap,
            out HashSet<string> nameSet,
            out HashSet<string> coordSigSet,
            out Dictionary<string, (double x, double y, double z)> nameToCoord,
            out Dictionary<string, string> coordSigToName)
        {
            nameSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            coordSigSet = new HashSet<string>(StringComparer.Ordinal);
            nameToCoord = new Dictionary<string, (double x, double y, double z)>(StringComparer.OrdinalIgnoreCase);
            coordSigToName = new Dictionary<string, string>(StringComparer.Ordinal);

            int n = 0;
            string[] names = null;
            if (sap.PointObj.GetNameList(ref n, ref names) != 0 || names == null) return;

            for (int i = 0; i < n; i++)
            {
                string nm = names[i];
                if (string.IsNullOrWhiteSpace(nm)) continue;

                double x = 0, y = 0, z = 0;
                if (sap.PointObj.GetCoordCartesian(nm, ref x, ref y, ref z) == 0)
                {
                    string exact = nm.Trim();
                    string sig = CoordSig(x, y, z);

                    nameSet.Add(exact);
                    nameToCoord[exact] = (x, y, z);
                    coordSigSet.Add(sig);

                    if (!coordSigToName.ContainsKey(sig))
                        coordSigToName[sig] = exact;
                }
            }
        }

        private static bool TryCheckConnectivity(cSapModel sap, string name, out bool hasConnectivity)
        {
            hasConnectivity = false;
            int count = 0;
            int[] types = null;
            string[] objNames = null;
            int[] pointNumbers = null;

            int ret = sap.PointObj.GetConnectivity(name, ref count, ref types, ref objNames, ref pointNumbers);
            if (ret != 0) return false;

            hasConnectivity = count > 0;
            return true;
        }

        private static bool TryGetPoint(cSapModel sap, string name, out double x, out double y, out double z)
        {
            x = y = z = 0;
            return sap.PointObj.GetCoordCartesian(name, ref x, ref y, ref z) == 0;
        }

        private static bool TryFindExactNameIgnoreCase(
            Dictionary<string, (double x, double y, double z)> nameToCoord,
            string desired,
            out string exactExisting)
        {
            foreach (var k in nameToCoord.Keys)
            {
                if (string.Equals(k, desired, StringComparison.OrdinalIgnoreCase))
                {
                    exactExisting = k;
                    return true;
                }
            }
            exactExisting = null;
            return false;
        }

        private static string CoordSig(double x, double y, double z)
            => $"{Math.Round(x, 6)}|{Math.Round(y, 6)}|{Math.Round(z, 6)}";

        private static string ReadText(object value)
        {
            if (value == null) return string.Empty;
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
            return text.Trim();
        }

        private static double? ReadNullableDouble(object value)
        {
            if (value == null) return null;
            if (value is double direct) return direct;

            string text = ReadText(value);
            if (string.IsNullOrEmpty(text)) return null;

            if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out double result))
                return result;

            return null;
        }

        // ===== Trigger GhcGetPointInfo refresh (post-update) =====
        private int TriggerGetPointInfoRefresh()
        {
            GH_Document doc = OnPingDocument();
            if (doc == null) return 0;

            // Collect strong-typed GhcGetPointInfo first (fast & exact)
            var targets = new List<GhcGetPointInfo>();
            foreach (IGH_DocumentObject obj in doc.Objects)
            {
                if (obj is GhcGetPointInfo g && !g.Locked && !g.Hidden)
                    targets.Add(g);
            }

            // Fallback: also refresh any GH_Component named "Get Point Info" (defensive)
            var fallback = new List<GH_Component>();
            foreach (IGH_DocumentObject obj in doc.Objects)
            {
                if (obj is GH_Component c && !c.Locked && !c.Hidden &&
                    string.Equals(c.Name, "Get Point Info", StringComparison.OrdinalIgnoreCase))
                {
                    // Avoid duplicates if GhcGetPointInfo already captured
                    if (c is not GhcGetPointInfo) fallback.Add(c);
                }
            }

            int total = targets.Count + fallback.Count;
            if (total == 0) return 0;

            // Schedule a small deferred refresh so Grasshopper UI can repaint cleanly
            doc.ScheduleSolution(5, _ =>
            {
                foreach (var g in targets) g.ExpireSolution(false);
                foreach (var c in fallback) c.ExpireSolution(false);
            });

            return total;
        }

        // ===== Data structures =====
        private sealed class PointRow
        {
            public string UniqueName { get; set; }
            public double? X { get; set; }
            public double? Y { get; set; }
            public double? Z { get; set; }
        }

        private sealed class PointBaseline
        {
            private readonly List<Entry> _entries = new();
            private bool _has;

            private PointBaseline() { }

            public static PointBaseline FromTree(GH_Structure<IGH_Goo> tree)
            {
                var b = new PointBaseline();
                if (tree == null) return b;

                var names = new List<string>(ReadStringBranch(tree, 0));
                var xs = new List<double?>(ReadDoubleBranch(tree, 1));
                var ys = new List<double?>(ReadDoubleBranch(tree, 2));
                var zs = new List<double?>(ReadDoubleBranch(tree, 3));

                int n = Math.Max(Math.Max(names.Count, xs.Count), Math.Max(ys.Count, zs.Count));
                for (int i = 0; i < n; i++)
                {
                    string name = i < names.Count ? (names[i] ?? string.Empty).Trim() : string.Empty;
                    double? x = i < xs.Count ? xs[i] : null;
                    double? y = i < ys.Count ? ys[i] : null;
                    double? z = i < zs.Count ? zs[i] : null;
                    b._entries.Add(new Entry(name, x, y, z));
                }

                b._has = tree.DataCount > 0;
                return b;
            }

            public bool HasData => _has;
            public IReadOnlyList<Entry> Entries => _entries;

            private static IEnumerable<string> ReadStringBranch(GH_Structure<IGH_Goo> tree, int index)
            {
                if (index < 0 || index >= tree.PathCount) yield break;
                IList branch = tree.get_Branch(index);
                if (branch == null) yield break;

                foreach (object item in branch)
                {
                    if (item is IGH_Goo goo && GH_Convert.ToString(goo, out string s, GH_Conversion.Both))
                        yield return s?.Trim() ?? string.Empty;
                    else
                        yield return string.Empty;
                }
            }

            private static IEnumerable<double?> ReadDoubleBranch(GH_Structure<IGH_Goo> tree, int index)
            {
                if (index < 0 || index >= tree.PathCount) yield break;
                IList branch = tree.get_Branch(index);
                if (branch == null) yield break;

                foreach (object item in branch)
                {
                    if (item is IGH_Goo goo && GH_Convert.ToDouble(goo, out double v, GH_Conversion.Both))
                        yield return v;
                    else
                        yield return null;
                }
            }

            internal readonly record struct Entry(string UniqueName, double? X, double? Y, double? Z);
        }
    }

    // ===== Small utilities (local) =====
    internal static class LocalUtil
    {
        /// <summary>Enumerate with 0-based indices: foreach (var (item, i) in list.WithIndex())</summary>
        public static IEnumerable<(T item, int index)> WithIndex<T>(this IEnumerable<T> source)
        {
            int i = 0;
            foreach (var it in source) yield return (it, i++);
        }
    }
}
