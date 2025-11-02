using System;
using System.Collections.Generic;
using ETABSv1;

//testttttttt
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

    /// <summary>
    /// Maintains an insertion-ordered list of entries together with a lookup by key.
    /// </summary>
    /// <typeparam name="TKey">Key type used for dictionary lookups.</typeparam>
    /// <typeparam name="TValue">Value type stored for each entry.</typeparam>
    internal sealed class OrderedLookup<TKey, TValue>
    {
        // Keep an ordered backing list so callers can iterate entries in the
        // exact sequence they were recorded from Excel/baseline captures.
        private readonly List<TValue> _entries = new List<TValue>();
        // Maintain a dictionary side-car to expose O(1) lookups without
        // disturbing the insertion order stored in _entries.
        private readonly Dictionary<TKey, TValue> _lookup;

        internal OrderedLookup()
            : this(EqualityComparer<TKey>.Default)
        {
        }

        internal OrderedLookup(IEqualityComparer<TKey> comparer)
        {
            _lookup = new Dictionary<TKey, TValue>(comparer ?? EqualityComparer<TKey>.Default);
        }

        internal int Count => _entries.Count;

        internal IReadOnlyList<TValue> Entries => _entries;

        internal void Add(TKey key, TValue value)
        {
            // Always append to the ordered list first; this guarantees
            // enumeration reflects the original capture order even when the
            // same key appears multiple times (dictionary ignores duplicates).
            _entries.Add(value);

            if (key == null)
            {
                return;
            }

            // Only seed the lookup when we see a key for the first time; later
            // duplicates keep their place in _entries but should not overwrite
            // the first occurrence used for keyed access.
            if (!_lookup.ContainsKey(key))
            {
                _lookup.Add(key, value);
            }
        }

        internal bool TryGetValue(TKey key, out TValue value)
        {
            if (key == null)
            {
                value = default;
                return false;
            }

            return _lookup.TryGetValue(key, out value);
        }
    }
}
