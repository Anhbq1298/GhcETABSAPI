using System;
using System.Collections.Generic;
using ETABSv1;

namespace MGT
{
    internal static class ComponentShared
    {
        internal static void EnsureModelUnlocked(cSapModel model)
        {
            if (model == null)
            {
                return;
            }

            try
            {
                if (model.GetModelIsLocked())
                {
                    model.SetModelIsLocked(false);
                }
            }
            catch
            {
                // ignored
            }
        }

        internal static HashSet<string> TryGetExistingFrameNames(cSapModel model)
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
                if (names == null)
                {
                    return result;
                }

                for (int i = 0; i < names.Length; i++)
                {
                    string nm = names[i];
                    if (!string.IsNullOrWhiteSpace(nm))
                    {
                        result.Add(nm.Trim());
                    }
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        internal static double Clamp01(double value)
        {
            if (value < 0.0)
            {
                return 0.0;
            }

            if (value > 1.0)
            {
                return 1.0;
            }

            return value;
        }

        internal static int ClampDirCode(int dirCode)
        {
            if (dirCode < 1 || dirCode > 11)
            {
                return 10;
            }

            return dirCode;
        }

        internal static string ResolveDirectionReference(int direction)
        {
            return Math.Abs(direction) < 4 ? "Local" : "Global";
        }

        internal static bool IsInvalidNumber(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value);
        }

        internal static string ResolveDirectionReferenceArea(int direction)
        {
            return Math.Abs(direction) < 4 ? "Local" : "Global";
        }


        internal static double? TryGet(IList<double> source, int index)
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

        internal static string Plural(int count, string word)
        {
            return count == 1 ? $"{count} {word}" : $"{count} {word}s";
        }

        internal static List<string> NormalizeDistinct(IList<string> source)
        {
            List<string> result = new List<string>();
            if (source == null)
            {
                return result;
            }

            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < source.Count; i++)
            {
                string entry = source[i];
                if (string.IsNullOrWhiteSpace(entry))
                {
                    continue;
                }

                string trimmed = entry.Trim();
                if (trimmed.Length == 0)
                {
                    continue;
                }

                if (seen.Add(trimmed))
                {
                    result.Add(trimmed);
                }
            }

            return result;
        }

        internal static HashSet<string> TryGetExistingAreaNames(cSapModel model)
        {
            if (model == null)
            {
                return null;
            }

            try
            {
                int count = 0;
                string[] names = null;
                int ret = model.AreaObj.GetNameList(ref count, ref names);
                if (ret != 0)
                {
                    return null;
                }

                HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (names == null)
                {
                    return result;
                }

                for (int i = 0; i < names.Length; i++)
                {
                    string nm = names[i];
                    if (!string.IsNullOrWhiteSpace(nm))
                    {
                        result.Add(nm.Trim());
                    }
                }

                return result;
            }
            catch
            {
                return null;
            }
        }

        internal static void TryRefreshView(cSapModel model)
        {
            if (model == null)
            {
                return;
            }

            try
            {
                model.View.RefreshView(0, false);
            }
            catch
            {
                // ignored
            }
        }

        internal static string FormatLoadPatternSummary(IReadOnlyList<string> filters)
        {
            if (filters == null || filters.Count == 0)
            {
                return string.Empty;
            }

            if (filters.Count == 1)
            {
                return $"load pattern \"{filters[0]}\"";
            }

            return $"load patterns ({string.Join(", ", filters)})";
        }
    }
}
