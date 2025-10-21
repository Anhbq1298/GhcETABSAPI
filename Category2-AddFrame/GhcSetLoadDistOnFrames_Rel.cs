// -------------------------------------------------------------
// Component : Set Frame Distributed Loads (relative distances)
// Author    : Anh Bui
// Target    : Rhino 7/8 + Grasshopper, .NET Framework 4.8 (x64)
// Depends   : Grasshopper, ETABSv1 (COM)  [Embed Interop Types = False]
// Panel     : "ETABS API" / "2.0 Frame Object Modelling"
// -------------------------------------------------------------
// Inputs (ordered):
//   0) run         (bool, item)    Rising-edge trigger.
//   1) sapModel    (ETABSv1.cSapModel, item)  ETABS model from Attach component.
//   2) frameNames  (string, list)  Names of frame objects to receive the load (order preserved).
//   3) loadPattern (string, item)  Target load pattern name.
//   4) myType      (int, item)     1 = Uniform, 2 = Trapezoidal.
//   5) dirCode     (int, item)     Direction code 1..11.
//   6) dist1       (double, list)  Relative start distance (0..1). Missing/invalid skips the frame.
//   7) dist2       (double, list)  Relative end distance   (0..1). Missing/invalid skips the frame.
//   8) val1        (double, list)  Start intensity [F/L].  Missing/invalid skips the frame.
//   9) val2        (double, list)  End intensity   [F/L].  Missing/invalid skips the frame.
//  10) replaceMode (bool, item)    True = replace, False = add.
//
// Outputs:
//   0) messages    (string, list)  Summary + optional failure/skip diagnostics.
//
// Behavior Notes:
//   • Uses FrameObj.SetLoadDistributed with IsRelativeDist = true and eItemType.Objects.
//   • Attempts to unlock the model automatically before assignment.
//   • Frame existence pre-checked via FrameObj.GetNameList when available.
//   • Direction codes 1..3 use the Local coordinate system; others use Global.
//   • Distances are clamped to [0,1]; swapped when start > end.
//   • Any frame with missing/invalid distances or values is reported as "skipped".
//   • When run is false or not toggled, the component replays the last output messages.
// -------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Drawing;
using ETABSv1;
using Grasshopper.Kernel;

namespace GhcETABSAPI
{
    public class GhcSetLoadDistOnFrames_Rel : GH_Component
    {
        private bool _lastRun;
        private readonly List<string> _lastMessages = new List<string> { "No previous run. Toggle 'run' to assign." };

        public GhcSetLoadDistOnFrames_Rel()
          : base(
                "Set Frame Distributed Loads (Rel)",
                "SetFrameUDLRel",
                "Assign distributed loads to ETABS frame objects using relative distances (0..1).",
                "ETABS API",
                "2.0 Frame Object Modelling")
        {
        }

        public override Guid ComponentGuid => new Guid("8AA2C49F-30B1-4E21-803C-5F85AB4A0C5B");

        protected override Bitmap Icon => null;

        protected override void RegisterInputParams(GH_InputParamManager p)
        {
            p.AddBooleanParameter("run", "run", "Rising-edge trigger; executes when this turns True.", GH_ParamAccess.item, false);
            p.AddGenericParameter("sapModel", "sapModel", "ETABS cSapModel from the Attach component.", GH_ParamAccess.item);
            p.AddTextParameter("frameNames", "frameNames", "Frame object names to assign (order preserved).", GH_ParamAccess.list);
            p.AddTextParameter("loadPattern", "loadPattern", "Target ETABS load pattern name.", GH_ParamAccess.item, string.Empty);
            p.AddIntegerParameter("myType", "myType", "Distributed load type: 1 = Uniform, 2 = Trapezoidal.", GH_ParamAccess.item, 1);
            p.AddIntegerParameter("dirCode", "dirCode", "Direction code 1..11.", GH_ParamAccess.item, 10);
            p.AddNumberParameter("dist1", "dist1", "Relative start distance (0..1). Missing/invalid skips the frame.", GH_ParamAccess.list);
            p.AddNumberParameter("dist2", "dist2", "Relative end distance (0..1). Missing/invalid skips the frame.", GH_ParamAccess.list);
            p.AddNumberParameter("val1", "val1", "Distributed load start intensity (F/L). Missing/invalid skips the frame.", GH_ParamAccess.list);
            p.AddNumberParameter("val2", "val2", "Distributed load end intensity (F/L). Missing/invalid skips the frame.", GH_ParamAccess.list);
            p.AddBooleanParameter("replaceMode", "replace", "True = replace existing values, False = add.", GH_ParamAccess.item, true);
        }

        protected override void RegisterOutputParams(GH_OutputParamManager p)
        {
            p.AddTextParameter("messages", "messages", "Summary and diagnostic messages.", GH_ParamAccess.list);
        }

        protected override void SolveInstance(IGH_DataAccess da)
        {
            bool run = false;
            cSapModel sapModel = null;
            List<string> frameNames = new List<string>();
            string loadPattern = null;
            int myType = 1;
            int dirCode = 10;
            List<double> dist1 = new List<double>();
            List<double> dist2 = new List<double>();
            List<double> val1 = new List<double>();
            List<double> val2 = new List<double>();
            bool replaceMode = true;

            da.GetData(0, ref run);
            da.GetData(1, ref sapModel);
            da.GetDataList(2, frameNames);
            da.GetData(3, ref loadPattern);
            da.GetData(4, ref myType);
            da.GetData(5, ref dirCode);
            da.GetDataList(6, dist1);
            da.GetDataList(7, dist2);
            da.GetDataList(8, val1);
            da.GetDataList(9, val2);
            da.GetData(10, ref replaceMode);

            bool rising = !_lastRun && run;
            if (!rising)
            {
                da.SetDataList(0, _lastMessages);
                _lastRun = run;
                return;
            }

            List<string> messages = new List<string>();

            try
            {
                if (sapModel == null)
                {
                    throw new InvalidOperationException("sapModel is null. Wire it from the Attach component.");
                }

                EnsureModelUnlocked(sapModel);

                if (frameNames == null || frameNames.Count == 0)
                {
                    throw new InvalidOperationException("frameNames is empty.");
                }

                string pattern = (loadPattern ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(pattern))
                {
                    throw new InvalidOperationException("loadPattern is empty.");
                }

                int loadType = (myType == 2) ? 2 : 1;
                int direction = ClampDirCode(dirCode);
                string coordinateSystem = ResolveCoordinateSystem(null, direction);
                bool replaceFlag = replaceMode;

                int frameCount = frameNames.Count;
                List<string> cleanedNames = new List<string>(frameCount);
                for (int i = 0; i < frameCount; i++)
                {
                    string nm = frameNames[i];
                    cleanedNames.Add(string.IsNullOrWhiteSpace(nm) ? string.Empty : nm.Trim());
                }

                HashSet<string> existingNames = TryGetExistingFrameNames(sapModel);
                bool[] existsMask = new bool[frameCount];

                int assignedCount = 0;
                int failedCount = 0;
                List<string> failedPairs = new List<string>();
                List<string> skippedPairs = new List<string>();
                HashSet<string> skipSet = new HashSet<string>();

                if (existingNames != null)
                {
                    for (int i = 0; i < frameCount; i++)
                    {
                        string nm = cleanedNames[i];
                        if (string.IsNullOrEmpty(nm))
                        {
                            string pair = $"{i}:{nm}";
                            if (skipSet.Add(pair))
                            {
                                skippedPairs.Add(pair);
                            }
                            continue;
                        }

                        bool exists = existingNames.Contains(nm);
                        existsMask[i] = exists;
                        if (!exists)
                        {
                            failedCount++;
                            failedPairs.Add($"{i}:{nm}");
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < frameCount; i++)
                    {
                        existsMask[i] = true;
                    }
                }

                for (int i = 0; i < frameCount; i++)
                {
                    string frameName = cleanedNames[i];

                    if (string.IsNullOrEmpty(frameName))
                    {
                        string pair = $"{i}:{frameName}";
                        if (skipSet.Add(pair))
                        {
                            skippedPairs.Add(pair);
                        }
                        continue;
                    }

                    if (!existsMask[i])
                    {
                        continue;
                    }

                    double? rawD1 = TryGet(dist1, i);
                    double? rawD2 = TryGet(dist2, i);
                    double? rawV1 = TryGet(val1, i);
                    double? rawV2 = TryGet(val2, i);

                    if (!rawD1.HasValue || !rawD2.HasValue || !rawV1.HasValue || !rawV2.HasValue)
                    {
                        string pair = $"{i}:{frameName}";
                        if (skipSet.Add(pair))
                        {
                            skippedPairs.Add(pair);
                        }
                        continue;
                    }

                    if (IsInvalidNumber(rawD1.Value) || IsInvalidNumber(rawD2.Value) || IsInvalidNumber(rawV1.Value) || IsInvalidNumber(rawV2.Value))
                    {
                        string pair = $"{i}:{frameName}";
                        if (skipSet.Add(pair))
                        {
                            skippedPairs.Add(pair);
                        }
                        continue;
                    }

                    double d1 = Clamp01(rawD1.Value);
                    double d2 = Clamp01(rawD2.Value);
                    double v1 = rawV1.Value;
                    double v2 = rawV2.Value;

                    if (d1 > d2)
                    {
                        double tmp = d1;
                        d1 = d2;
                        d2 = tmp;
                    }

                    int ret = sapModel.FrameObj.SetLoadDistributed(
                        frameName,
                        pattern,
                        loadType,
                        direction,
                        d1,
                        d2,
                        v1,
                        v2,
                        coordinateSystem,
                        true,
                        replaceFlag,
                        (int)eItemType.Objects);

                    if (ret == 0)
                    {
                        assignedCount++;
                    }
                    else
                    {
                        failedCount++;
                        failedPairs.Add($"{i}:{frameName}");
                    }
                }

                string summary = $"{Plural(assignedCount, "member")} successfully assigned, {Plural(failedCount, "member")} unsuccessful.";
                messages.Add(summary);

                if (failedPairs.Count > 0)
                {
                    messages.Add("Unsuccessful members (0-based index:name): " + string.Join(", ", failedPairs));
                }

                if (skippedPairs.Count > 0)
                {
                    messages.Add("Skipped members (0-based index:name): " + string.Join(", ", skippedPairs));
                }

                try
                {
                    sapModel.View.RefreshView(0, false);
                }
                catch
                {
                    // ignored
                }
            }
            catch (Exception ex)
            {
                string summary = $"{Plural(0, "member")} successfully assigned, {Plural(1, "member")} unsuccessful.";
                messages.Add(summary);
                messages.Add("Error: " + ex.Message);
            }

            da.SetDataList(0, messages);

            _lastMessages.Clear();
            _lastMessages.AddRange(messages);
            _lastRun = run;
        }

        private static void EnsureModelUnlocked(cSapModel model)
        {
            if (model == null)
            {
                return;
            }

            try
            {
                bool isLocked = false;

                isLocked = model.GetModelIsLocked();
                if (isLocked)
                {
                    model.SetModelIsLocked(false);
                }
            }
            catch
            {
                // ignored
            }
        }

        private static HashSet<string> TryGetExistingFrameNames(cSapModel model)
        {
            if (model == null)
            {
                return null;
            }

            try
            {
                int count = 0;
                string[] names = null;
                int ret = model.FrameObj.GetNameList(ref count, ref names);
                if (ret != 0)
                {
                    return null;
                }

                HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (names != null)
                {
                    for (int i = 0; i < names.Length; i++)
                    {
                        string nm = names[i];
                        if (!string.IsNullOrWhiteSpace(nm))
                        {
                            result.Add(nm.Trim());
                        }
                    }
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        private static double Clamp01(double value)
        {
            if (value < 0.0) return 0.0;
            if (value > 1.0) return 1.0;
            return value;
        }

        private static int ClampDirCode(int dirCode)
        {
            if (dirCode < 1 || dirCode > 11)
            {
                return 10;
            }

            return dirCode;
        }

        private static string ResolveCoordinateSystem(string coordinateSystem, int direction)
        {
            string directionReference = Math.Abs(direction) < 10 ? "Local" : "Global";

            if (string.IsNullOrWhiteSpace(coordinateSystem))
            {
                return directionReference;
            }

            string trimmed = coordinateSystem.Trim();

            if (string.Equals(trimmed, "Local", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Local", StringComparison.OrdinalIgnoreCase))
            {
                return "Local";
            }

            if (string.Equals(trimmed, "Global", StringComparison.OrdinalIgnoreCase) ||
                trimmed.StartsWith("Global", StringComparison.OrdinalIgnoreCase))
            {
                return "Global";
            }

            return directionReference;
        }

        private static bool IsInvalidNumber(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value);
        }

        private static double? TryGet(IList<double> source, int index)
        {
            if (source == null)
            {
                return null;
            }

            if (index < 0 || index >= source.Count)
            {
                return null;
            }

            return source[index];
        }

        private static string Plural(int count, string word)
        {
            return count == 1 ? $"{count} {word}" : $"{count} {word}s";
        }
    }
}
