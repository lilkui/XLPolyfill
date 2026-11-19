using System.Buffers;
using System.Globalization;
using ExcelDna.Integration;

namespace XLPolyfill;

public static class LookupFunctions
{
    private const int WorksheetBlockSize = 512;
    private const int WorksheetCellWorkingBytes = 32;
    private const int PositionBudgetExceeded = -2;
    private const int DictionaryEntryBytes = 64;
    // shortcut: estimate text work in 64-character units until Excel timings can guide the budget.
    private const int WildcardWorkLimit = 10_000_000;
    private static readonly object NotFound = ExcelError.ExcelErrorNA;
    private static readonly object InvalidValue = ExcelError.ExcelErrorValue;
    private static readonly object Zero = 0d;

    [ExcelFunction(Name = "XLOOKUP", Category = "Lookup & Reference",
        Description = "Searches a row or column and returns the corresponding result.")]
    public static object XLookup(
        [ExcelArgument(Name = "lookup_value", Description = "The value to find. Omit to match a blank cell.", AllowReference = true)]
        object lookupValue,
        [ExcelArgument(Name = "lookup_array", Description = "A single row or column to search.", AllowReference = true)]
        object lookupArray,
        [ExcelArgument(Name = "return_array", Description = "The corresponding range or array to return from.", AllowReference = true)]
        object returnArray,
        [ExcelArgument(Name = "[if_not_found]", Description = "The result for a missing match. Defaults to #N/A.", AllowReference = true)]
        object ifNotFound,
        [ExcelArgument(Name = "[match_mode]", Description = "0 exact, -1 next smaller, 1 next larger, or 2 wildcard.", AllowReference = true)]
        object matchMode,
        [ExcelArgument(Name = "[search_mode]", Description = "1 forward, -1 reverse, 2 binary ascending, or -2 binary descending.", AllowReference = true)]
        object searchMode) =>
        XLookup(lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode,
            static block => block.GetValue(), Environment.Is64BitProcess);

    internal static object XLookup(object lookupValue, object lookupArray, object returnArray, object ifNotFound,
        object matchMode, object searchMode, Func<ExcelReference, object> readBlock, bool is64BitProcess,
        long? remainingWorkingBudget = null, Func<long>? getWorksheetTextBudget = null, bool leadingCellOnly = false)
    {
        // Disjoint keys and modes contribute a scalar error to the output shape.
        lookupValue = DisjointAsValueError(lookupValue);
        matchMode = DisjointAsValueError(matchMode);
        searchMode = DisjointAsValueError(searchMode);
        if (IsModeArray(matchMode) || IsModeArray(searchMode))
            return XLookupWithModeArrays(lookupValue, lookupArray, returnArray, ifNotFound, matchMode, searchMode,
                readBlock, is64BitProcess);

        long workingBudget = remainingWorkingBudget ?? ((is64BitProcess ? 512L : 128L) << 20);
        ExcelArray? keyRange = lookupValue is ExcelReference reference ? new ExcelArray(reference, readBlock) : null;
        if (keyRange is { Rows: 1, Columns: 1 } singleCell)
        {
            lookupValue = singleCell.At(0, 0);
            keyRange = null;
        }
        if (lookupValue is ExcelError)
            return lookupValue;
        if (LookupArgumentError(lookupArray) is { } lookupError)
            return ReturnFailure(lookupError);
        if (LookupArgumentError(returnArray) is { } returnError)
            return ReturnFailure(returnError);
        if (matchMode is ExcelReference matchCell)
            matchMode = readBlock(matchCell);
        if (matchMode is ExcelError)
            return ReturnFailure(matchMode);
        double matchNumber = ReadMode(matchMode, 0);
        if (double.IsNaN(matchNumber))
            return ReturnFailure(InvalidValue);
        if (searchMode is ExcelReference searchCell)
            searchMode = readBlock(searchCell);
        if (searchMode is ExcelError)
            return ReturnFailure(searchMode);

        if (!TryValidateModes(matchNumber, ReadMode(searchMode, 1), out var modes))
            return ReturnFailure(InvalidValue);
        var (match, search) = modes;
        bool binary = search is 2 or -2;

        var lookup = new ExcelArray(lookupArray, readBlock);
        var returns = new ExcelArray(returnArray, readBlock);
        if (!lookup.IsValid || !returns.IsValid)
            return ReturnFailure(lookup.IsDisjointReference || returns.IsDisjointReference ? NotFound : InvalidValue);
        if (lookup.Rows > 1 && lookup.Columns > 1)
            return ReturnFailure(InvalidValue);

        bool vertical = lookup.Columns == 1 && (lookup.Rows > 1 || returns.Rows == 1);
        int length = vertical ? lookup.Rows : lookup.Columns;
        if (length != (vertical ? returns.Rows : returns.Columns))
            return ReturnFailure(InvalidValue);
        object fallback = ifNotFound switch
        {
            ExcelMissing => NotFound,
            ExcelEmpty => Zero,
            _ => ifNotFound
        };

        string localeName = CultureInfo.CurrentCulture.Name;
        if (keyRange is null && lookupValue is not object[,])
        {
            int index = Find(lookup, new LookupKey(lookupValue, localeName), length, vertical, match, search, localeName);
            if (index == PositionBudgetExceeded)
                return ExcelArrayResult.TooLarge;
            if (index >= 0 && returns.IsReference && getWorksheetTextBudget?.Invoke() < 0)
                return ExcelArrayResult.TooLarge;
            if (index >= 0 && leadingCellOnly)
                return returns.At(index, vertical);
            object result = index < 0 ? fallback : returns.GetResult(index, vertical);
            return result is object[,] array ? ExcelArrayResult.Check(array, is64BitProcess) : result;
        }

        var keyShape = keyRange ?? new ExcelArray(lookupValue, readBlock);
        if (!keyShape.IsValid)
            return InvalidValue;
        long cells = (long)keyShape.Rows * keyShape.Columns;
        long textBudget = ExcelArrayResult.TextBudget(cells, is64BitProcess);
        if (textBudget < 0)
            return ExcelArrayResult.TooLarge;
        long worksheetCopyAllowance = textBudget -
            ExcelArrayResult.TextBudget(cells, is64BitProcess, copiesWorksheetValues: true);
        object keyValues = keyRange is { } range ? range.ReadValues(ref workingBudget) : lookupValue;
        if (keyValues is not object[,] keys)
            return keyValues;

        // Excel takes the leading cell of each matched slice when given multiple lookup values.
        LookupKeyMap<int>? positions = null;
        Dictionary<string, LookupKey>? scanTextKeys = null;
        Dictionary<object, int> searched = [];
        LookupKeyMap<int>? searchedText = null;
        if (match == 2)
        {
            var wildcardPositions = FindWildcardPositions(lookup, keys, length, vertical, search == 1, ref workingBudget);
            if (wildcardPositions is null)
                return ExcelArrayResult.TooLarge;
            searched = wildcardPositions;
        }
        else if (!binary)
        {
            scanTextKeys = new(StringComparer.Ordinal);
            positions = FindScanPositions(lookup, keys, length, vertical, match, search == 1, localeName,
                scanTextKeys, ref workingBudget);
            if (positions is null)
                return ExcelArrayResult.TooLarge;
        }
        Func<int, object>? readProbe = null;
        Dictionary<int, (int Count, object? Value)>? returnValues = null;
        if (returns.IsReference)
        {
            returnValues = [];
            foreach (object key in keys)
            {
                if (key is ExcelError)
                    continue;
                int index = FindPosition(key);
                if (index == PositionBudgetExceeded)
                    return ExcelArrayResult.TooLarge;
                if (index >= 0)
                {
                    if (!returnValues.TryGetValue(index, out var entry) &&
                        !Reserve(ref workingBudget, DictionaryEntryBytes + sizeof(int)))
                        return ExcelArrayResult.TooLarge;
                    returnValues[index] = (entry.Count + 1, null);
                }
            }
            if (returnValues.Count > 0 &&
                (!UseWorksheetResultLimit() || !returns.ReadMatches(returnValues, vertical, ref textBudget)))
                return ExcelArrayResult.TooLarge;
        }

        var fallbacks = new ExcelArray(fallback, readBlock);
        int rows = keys.GetLength(0);
        int columns = keys.GetLength(1);
        bool shapedFallback = fallbacks.Rows == rows && fallbacks.Columns == columns;
        bool readFallbackInBlocks = shapedFallback && fallbacks.IsReference;
        bool hasMissingFallbacks = false;
        bool useFirstFallback = (fallback is object[,] && fallbacks.IsValid) ||
            (fallbacks.Rows == 1 && fallbacks.Columns == 1);
        object? scalarFallback = null;
        var results = new object[rows, columns];
        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                object key = keys[row, column];
                if (key is ExcelError)
                {
                    results[row, column] = key;
                    continue;
                }

                int index = FindPosition(key);
                if (index == PositionBudgetExceeded)
                    return ExcelArrayResult.TooLarge;
                if (index >= 0 && returnValues is not null)
                {
                    results[row, column] = returnValues[index].Value!;
                    continue;
                }

                object value;
                if (index >= 0)
                    value = returns.At(index, vertical);
                else if (readFallbackInBlocks)
                {
                    hasMissingFallbacks = true;
                    continue;
                }
                else if (shapedFallback)
                    value = fallbacks.At(row, column);
                else
                    value = useFirstFallback ? scalarFallback ??= fallbacks.At(0, 0) : InvalidValue;

                textBudget -= ExcelArrayResult.TextBytes(value);
                if (textBudget < 0)
                    return ExcelArrayResult.TooLarge;
                results[row, column] = index >= 0 ? value : ValueOrZero(value);
            }
        }
        if (hasMissingFallbacks &&
            (!UseWorksheetResultLimit() || !fallbacks.ReadMissingFallbacks(results, ref textBudget)))
            return ExcelArrayResult.TooLarge;
        return results;

        object ReturnFailure(object error)
        {
            if (keyRange is null && lookupValue is not object[,])
                return error;

            var shape = keyRange ?? new ExcelArray(lookupValue, readBlock);
            if (!shape.IsValid)
                return InvalidValue;
            if (ExcelArrayResult.TextBudget((long)shape.Rows * shape.Columns, is64BitProcess) < 0)
                return ExcelArrayResult.TooLarge;

            if (keyRange is { } range)
                return range.ReadValues(ref workingBudget, error);

            var input = (object[,])lookupValue;
            var output = new object[shape.Rows, shape.Columns];
            for (int row = 0; row < shape.Rows; row++)
                for (int column = 0; column < shape.Columns; column++)
                    output[row, column] = input[row, column] is ExcelError ? input[row, column] : error;
            return output;
        }

        bool UseWorksheetResultLimit()
        {
            textBudget -= worksheetCopyAllowance;
            worksheetCopyAllowance = 0;
            if (getWorksheetTextBudget is not null)
                textBudget = Math.Min(textBudget, getWorksheetTextBudget());
            return textBudget >= 0;
        }

        int FindPosition(object key)
        {
            if (positions is not null)
                return positions[key is string text ? scanTextKeys![text] : new LookupKey(key, localeName)];
            key = BlankLookupValue(key);
            if (!searched.TryGetValue(key, out int index))
            {
                if (!Reserve(ref workingBudget, DictionaryEntryBytes))
                    return PositionBudgetExceeded;
                var lookupKey = new LookupKey(key, localeName);
                if (key is string)
                {
                    searchedText ??= new(localeName);
                    if (searchedText.TryGetValue(lookupKey, out index))
                    {
                        searched.Add(key, index);
                        return index;
                    }
                    if (!searchedText.AddIfMissing(lookupKey, PositionBudgetExceeded, ref workingBudget))
                        return PositionBudgetExceeded;
                }
                if (lookup.IsReference && readProbe is null && searched.Count > 0)
                    readProbe = lookup.CacheProbes(length, vertical, bytes => Reserve(ref workingBudget, bytes));
                index = FindBinary(lookup, lookupKey, length, vertical, match, search == 2,
                    localeName, readProbe);
                searched.Add(key, index);
                if (key is string)
                    searchedText![lookupKey] = index;
            }
            return index;
        }
    }

    private static object? LookupArgumentError(object argument) => argument switch
    {
        ExcelError => argument,
        ExcelMissing => InvalidValue,
        _ => null
    };

    private static object DisjointAsValueError(object argument) =>
        argument is ExcelReference { InnerReferences.Count: > 1 } ? InvalidValue : argument;

    private static object XLookupWithModeArrays(object lookupValue, object lookupArray, object returnArray,
        object ifNotFound, object matchMode, object searchMode, Func<ExcelReference, object> readBlock, bool is64BitProcess)
    {
        var keys = new ExcelArray(BlankLookupValue(lookupValue), readBlock);
        var matches = new ExcelArray(matchMode is ExcelMissing ? 0d : matchMode, readBlock);
        var searches = new ExcelArray(searchMode is ExcelMissing ? 1d : searchMode, readBlock);
        if (!keys.IsValid || !matches.IsValid || !searches.IsValid)
            return InvalidValue;

        int rows = Math.Max(keys.Rows, Math.Max(matches.Rows, searches.Rows));
        int columns = Math.Max(keys.Columns, Math.Max(matches.Columns, searches.Columns));
        long cells = (long)rows * columns;
        long textBudget = ExcelArrayResult.TextBudget(cells, is64BitProcess);
        long worksheetCopyAllowance = textBudget -
            ExcelArrayResult.TextBudget(cells, is64BitProcess, copiesWorksheetValues: true);
        // Account for the spare capacity in each list of result coordinates and keys.
        long workingBudget = ((is64BitProcess ? 512L : 128L) << 20) - 32 * cells;
        if (textBudget < 0 || workingBudget < 0)
            return ExcelArrayResult.TooLarge;
        object? lookupError = LookupArgumentError(lookupArray) ?? LookupArgumentError(returnArray);
        object? readError = ReadWorksheetRange(ref keys, lookupError);
        if (readError is null && lookupError is null && keys.HasNonErrorValue())
        {
            readError = ReadWorksheetRange(ref matches);
            if (readError is null && matches.HasNonErrorValue())
                readError = ReadWorksheetRange(ref searches);
        }
        if (readError is not null)
            return readError;

        var results = new object[rows, columns];
        var groups = new Dictionary<(int Match, int Search), List<(int Index, object Key)>>();
        for (int row = 0; row < rows; row++)
        {
            for (int column = 0; column < columns; column++)
            {
                object? key = keys.BroadcastAt(row, column);
                if (lookupError is not null || key is null or ExcelError)
                {
                    results[row, column] = key switch
                    {
                        ExcelError => key,
                        null => NotFound,
                        _ => lookupError!
                    };
                    continue;
                }

                object? match = matches.BroadcastAt(row, column);
                if (match is null or ExcelError)
                {
                    results[row, column] = match ?? NotFound;
                    continue;
                }
                double matchNumber = ReadMode(match, 0);
                if (double.IsNaN(matchNumber))
                {
                    results[row, column] = InvalidValue;
                    continue;
                }
                object? search = searches.BroadcastAt(row, column);
                if (search is null or ExcelError)
                {
                    results[row, column] = search ?? NotFound;
                    continue;
                }
                if (!TryValidateModes(matchNumber, ReadMode(search, 1), out var modes))
                {
                    results[row, column] = InvalidValue;
                    continue;
                }

                if (!groups.TryGetValue(modes, out var group))
                    groups.Add(modes, group = []);
                group.Add((row * columns + column, key));
            }
        }

        // Excel returns #VALUE! for a reference fallback when a mode is an array.
        object fallback = ifNotFound switch
        {
            ExcelReference => InvalidValue,
            object[,] { Length: > 0 } values => values[0, 0],
            _ => ifNotFound
        };
        bool usedWorksheetLimit = false;
        foreach (var (modes, group) in groups)
        {
            object groupKeys = group[0].Key;
            long groupBudget = workingBudget;
            if (group.Count > 1)
            {
                if (!Reserve(ref groupBudget, group.Count * (long)IntPtr.Size))
                    return ExcelArrayResult.TooLarge;
                var keyColumn = new object[group.Count, 1];
                for (int i = 0; i < group.Count; i++)
                    keyColumn[i, 0] = group[i].Key;
                groupKeys = keyColumn;
            }

            usedWorksheetLimit = false;
            object found = XLookup(groupKeys, lookupArray, returnArray, fallback,
                (double)modes.Match, (double)modes.Search, readBlock, is64BitProcess, groupBudget, GetWorksheetTextBudget,
                leadingCellOnly: true);
            // A group that exhausts the shared worksheet allowance makes the whole result too large.
            if (usedWorksheetLimit && ReferenceEquals(found, ExcelArrayResult.TooLarge))
                return ExcelArrayResult.TooLarge;
            var values = new ExcelArray(found, readBlock);
            for (int i = 0; i < group.Count; i++)
            {
                object value = values.At(values.Rows == 1 ? 0 : i, 0);
                textBudget -= ExcelArrayResult.TextBytes(value);
                if (textBudget < 0)
                    return ExcelArrayResult.TooLarge;
                int index = group[i].Index;
                results[index / columns, index % columns] = value;
            }
        }
        return results;

        object? ReadWorksheetRange(ref ExcelArray array, object? replaceNonErrors = null)
        {
            if (!array.IsReference)
                return null;
            object values = array.ReadValues(ref workingBudget, replaceNonErrors);
            if (values is not object[,] cells)
                return values;
            array = new ExcelArray(cells, readBlock);
            return null;
        }

        long GetWorksheetTextBudget()
        {
            usedWorksheetLimit = true;
            textBudget -= worksheetCopyAllowance;
            worksheetCopyAllowance = 0;
            return textBudget;
        }
    }

    private static bool IsModeArray(object mode) => mode switch
    {
        object[,] => true,
        ExcelReference reference => reference.RowFirst != reference.RowLast || reference.ColumnFirst != reference.ColumnLast,
        _ => false
    };

    private static bool Reserve(ref long budget, long bytes)
    {
        if (bytes > budget)
            return false;
        budget -= bytes;
        return true;
    }

    private static Dictionary<object, int>? FindWildcardPositions(ExcelArray lookup, object[,] keys, int length,
        bool vertical, bool forward, ref long workingBudget)
    {
        // shortcut: cap substring fragments at 12 characters until longer fragments improve worksheet timings.
        const int maxFragmentLength = 12;
        const int fragmentIndexEntryBytes = 80 + 2 * (DictionaryEntryBytes + 24 + maxFragmentLength * sizeof(char));
        int remainingWork = WildcardWorkLimit;
        var positions = new Dictionary<object, int>();
        var patternsByMinimumLength = new SortedDictionary<int, List<(string Key, ExcelWildcard Pattern)>>();
        var anchored = new Dictionary<(bool FromEnd, char Edge, int Width), Dictionary<string, List<(string Key, ExcelWildcard Pattern)>>>();
        var substrings = new Dictionary<int, Dictionary<string, WildcardBucket>>();
        var literalIndexes = new List<(string Key, ExcelWildcard Pattern, string Primary, string? Alternate, bool? AnchorFromEnd)>();
        var fragmentCounts = new Dictionary<string, int>(StringComparer.Ordinal);
        int remainingText = 0;
        foreach (object value in keys)
        {
            object key = BlankLookupValue(value);
            if (value is ExcelError || positions.ContainsKey(key))
                continue;
            if (!Reserve(ref workingBudget, DictionaryEntryBytes))
                return null;
            positions.Add(key, -1);
            if (value is not string pattern)
                continue;
            var wildcard = new ExcelWildcard(pattern);
            if (!Reserve(ref workingBudget, wildcard.EstimatedBytes))
                return null;
            remainingText++;
            if (wildcard.RequiredLiteral is not { } required)
            {
                if (!patternsByMinimumLength.TryGetValue(wildcard.MinimumLength, out var patterns))
                    patternsByMinimumLength.Add(wildcard.MinimumLength, patterns = []);
                patterns.Add((pattern, wildcard));
                continue;
            }

            if (!Reserve(ref workingBudget, fragmentIndexEntryBytes))
                return null;
            string? prefix = wildcard.StartingLiteral;
            string? anchor = prefix ?? wildcard.EndingLiteral;
            string primary = anchor ?? GetFragment(required);
            string? other = anchor is not null && anchor != required ? required : wildcard.AlternateLiteral;
            string? alternate = other is null ? null : GetFragment(other);
            if (anchor is null && (alternate is null || alternate == primary))
                alternate = required[..primary.Length];
            if (alternate == primary)
                alternate = null;
            CountFragment(primary);
            if (alternate is not null)
                CountFragment(alternate);
            literalIndexes.Add((pattern, wildcard, primary, alternate, anchor is null ? null : prefix is null));
        }
        int remaining = positions.Count;
        if (remaining == 0)
            return positions;

        foreach (var (key, wildcard, primary, alternate, anchorFromEnd) in literalIndexes)
        {
            // A literal shared by fewer patterns usually gives each matching cell fewer candidates.
            bool useAlternate = alternate is not null && fragmentCounts[alternate] < fragmentCounts[primary];
            if (!useAlternate && anchorFromEnd is { } fromEnd)
            {
                var anchorGroup = (fromEnd, fromEnd ? primary[^1] : primary[0], primary.Length);
                if (!anchored.TryGetValue(anchorGroup, out var literals))
                    anchored[anchorGroup] = literals = new(StringComparer.Ordinal);
                if (!literals.TryGetValue(primary, out var anchoredBucket))
                    literals[primary] = anchoredBucket = [];
                anchoredBucket.Add((key, wildcard));
                continue;
            }

            string fragment = useAlternate ? alternate! : primary;
            if (!substrings.TryGetValue(fragment.Length, out var fragments))
                substrings[fragment.Length] = fragments = new(StringComparer.Ordinal);
            if (!fragments.TryGetValue(fragment, out var bucket))
                fragments[fragment] = bucket = new();
            bucket.Patterns.Add((key, wildcard));
        }

        var substringGroups = substrings.Select(group => (Width: group.Key,
            Search: SearchValues.Create(group.Value.Keys.ToArray(), StringComparison.Ordinal),
            Buckets: group.Value, StaleHits: 0)).ToArray();
        var anchoredGroups = anchored.GroupBy(group => (group.Key.FromEnd, group.Key.Edge))
            .ToDictionary(group => group.Key, group => group.OrderBy(entry => entry.Key.Width).ToList());
        var minimumLengthGroups = new LinkedList<(int MinimumLength, List<(string Key, ExcelWildcard Pattern)> Patterns)>(
            patternsByMinimumLength.Select(group => (group.Key, group.Value)));

        var checkedTexts = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (index, value) in lookup.Scan(0, length, vertical, forward, probeFirstCell: remaining == 1))
        {
            if (value is string text)
            {
                if (remainingText == 0)
                    continue;
                if (!SpendWildcardScanWork(text, checkedTexts, ref remainingWork, out bool alreadyChecked))
                    return null;
                if (alreadyChecked)
                    continue;
                ReadOnlySpan<char> folded = WindowsText.FoldWildcardText(text);
                for (var group = minimumLengthGroups.First; group is not null && group.Value.MinimumLength <= folded.Length;)
                {
                    var next = group.Next;
                    var candidates = group.Value.Patterns;
                    if (!MatchPatterns(candidates, folded, index))
                        return null;
                    if (candidates.Count == 0)
                        minimumLengthGroups.Remove(group);
                    group = next;
                }
                if (!folded.IsEmpty && (!MatchAnchors(folded, index, fromEnd: false) || !MatchAnchors(folded, index, fromEnd: true)))
                    return null;
                foreach (ref var group in substringGroups.AsSpan())
                {
                    int uncheckedBuckets = group.Buckets.Count;
                    int width = group.Width;
                    if (uncheckedBuckets == 0 || folded.Length < width)
                        continue;
                    var fragments = group.Buckets.GetAlternateLookup<ReadOnlySpan<char>>();
                    ReadOnlySpan<char> rest = folded;
                    while (uncheckedBuckets > 0 && rest.Length >= width)
                    {
                        int offset = rest.IndexOfAny(group.Search);
                        if (!SpendWork(offset < 0 ? rest.Length : offset + width))
                            return null;
                        if (offset < 0)
                            break;
                        var fragment = rest.Slice(offset, width);
                        if (fragments.TryGetValue(fragment, out var bucket))
                        {
                            if (bucket.LastChecked != index)
                            {
                                bucket.LastChecked = index;
                                uncheckedBuckets--;
                                if (!MatchPatterns(bucket.Patterns, folded, index))
                                    return null;
                                if (bucket.Patterns.Count == 0)
                                    fragments.Remove(fragment);
                            }
                        }
                        // Rebuild once retired hits cost as much as visiting the remaining literals.
                        else if (++group.StaleHits >= group.Buckets.Count)
                        {
                            group.Search = SearchValues.Create(group.Buckets.Keys.ToArray(), StringComparison.Ordinal);
                            group.StaleHits = 0;
                        }
                        rest = rest[(offset + 1)..];
                    }
                }
            }
            else if (value is not ExcelError)
            {
                object key = BlankLookupValue(value);
                if (positions.TryGetValue(key, out int position) && position < 0)
                {
                    positions[key] = index;
                    remaining--;
                }
            }
            if (remaining == 0)
                break;
        }
        return positions;

        static string GetFragment(string literal) => literal[^Math.Min(literal.Length, maxFragmentLength)..];

        void CountFragment(string fragment) => fragmentCounts[fragment] = fragmentCounts.GetValueOrDefault(fragment) + 1;

        bool SpendWork(int characters) => ExcelWildcard.SpendCharacters(ref remainingWork, characters);

        bool MatchAnchors(ReadOnlySpan<char> text, int index, bool fromEnd)
        {
            char edge = fromEnd ? text[^1] : text[0];
            if (!anchoredGroups.TryGetValue((fromEnd, edge), out var groups))
                return true;

            bool retiredGroup = false;
            foreach (var (group, literals) in groups)
            {
                int width = group.Width;
                if (width > text.Length)
                    break;
                if (!SpendWork(width))
                    return false;
                var anchors = literals.GetAlternateLookup<ReadOnlySpan<char>>();
                ReadOnlySpan<char> anchor = fromEnd ? text[^width..] : text[..width];
                if (!anchors.TryGetValue(anchor, out var bucket))
                    continue;
                if (!MatchPatterns(bucket, text, index))
                    return false;
                if (bucket.Count == 0)
                {
                    anchors.Remove(anchor);
                    retiredGroup |= literals.Count == 0;
                }
            }
            if (retiredGroup)
                groups.RemoveAll(static group => group.Value.Count == 0);
            return true;
        }

        bool MatchPatterns(List<(string Key, ExcelWildcard Pattern)> candidates, ReadOnlySpan<char> text, int index)
        {
            for (int i = candidates.Count - 1; i >= 0; i--)
            {
                if (!SpendWork(text.Length))
                    return false;
                bool matches = candidates[i].Pattern.MatchesFolded(text, ref remainingWork);
                if (remainingWork < 0)
                    return false;
                if (!matches)
                    continue;
                positions[candidates[i].Key] = index;
                int last = candidates.Count - 1;
                candidates[i] = candidates[last];
                candidates.RemoveAt(last);
                remaining--;
                remainingText--;
            }
            return true;
        }
    }

    private sealed class WildcardBucket
    {
        public List<(string Key, ExcelWildcard Pattern)> Patterns { get; } = [];
        public int LastChecked { get; set; } = -1;
    }

    private static bool SpendWildcardScanWork(string text, HashSet<string> checkedTexts, ref int remainingWork,
        out bool alreadyChecked)
    {
        // shortcut: reset after 256 distinct cells until worksheet timings justify an eviction policy.
        const int cacheSize = 256;
        alreadyChecked = !checkedTexts.Add(text);
        if (checkedTexts.Count > cacheSize)
        {
            checkedTexts.Clear();
            checkedTexts.Add(text);
        }

        // Cache hits still hash and compare the text. They skip folding and matching.
        return ExcelWildcard.SpendCharacters(ref remainingWork, alreadyChecked ? text.Length / 4 : text.Length);
    }

    private static LookupKeyMap<int>? FindScanPositions(ExcelArray lookup, object[,] keys, int length,
        bool vertical, int match, bool forward, string localeName, Dictionary<string, LookupKey> textKeys, ref long workingBudget)
    {
        var positions = new LookupKeyMap<int>(localeName);
        foreach (object value in keys)
        {
            if (value is ExcelError || value is string existingText && textKeys.ContainsKey(existingText))
                continue;

            var key = new LookupKey(value, localeName);
            if (value is string text)
            {
                if (!Reserve(ref workingBudget, DictionaryEntryBytes + key.AdditionalTextBytes))
                    return null;
                textKeys.Add(text, key);
            }
            if (!positions.AddIfMissing(key, -1, ref workingBudget, countNormalizedText: false))
                return null;
        }

        // The ordered keys and nearest candidates are only retained during the scan.
        long scanBudget = workingBudget;
        if (match != 0 && !Reserve(ref scanBudget, 96L * positions.Count))
            return null;
        LookupKey[] ordered = match == 0 ? [] : positions.Keys.ToArray();
        Array.Sort(ordered, (left, right) => left.CompareTo(right, localeName));
        var nearest = new (LookupKey Key, int Index)?[ordered.Length];
        int remaining = positions.Count;
        if (remaining == 0)
            return positions;

        bool compareText = match != 0 || textKeys.Count != 0;
        foreach (var (index, value) in lookup.Scan(0, length, vertical, forward, probeFirstCell: remaining == 1))
        {
            if (!compareText && value is string)
                continue;
            var candidate = new LookupKey(value, localeName);
            if (positions.TryGetValue(candidate, out int position) && position < 0)
            {
                positions[candidate] = index;
                if (--remaining == 0)
                    break;
            }

            if (ordered.Length == 0)
                continue;

            // Each candidate competes only in the gap next to its position among the requested keys.
            int first = 0;
            int end = ordered.Length;
            while (first < end)
            {
                int middle = first + (end - first) / 2;
                int order = ordered[middle].CompareTo(candidate, localeName);
                if (order < 0 || (match == -1 && order == 0))
                    first = middle + 1;
                else
                    end = middle;
            }
            int gap = match == 1 ? first - 1 : first;
            if (gap < 0 || gap >= nearest.Length)
                continue;
            var current = nearest[gap];
            if (current is { } existing && Math.Sign(existing.Key.CompareTo(candidate, localeName)) != match)
                continue;
            if (!Reserve(ref scanBudget, candidate.RetainedValueBytes - (current?.Key.RetainedValueBytes ?? 0)))
                return null;
            nearest[gap] = (candidate, index);
        }

        int closest = -1;
        int slot = match == 1 ? ordered.Length - 1 : 0;
        int stop = match == 1 ? -1 : ordered.Length;
        // A candidate in a closer gap wins. Equal sort positions never replace an exact match.
        for (; slot != stop; slot -= match)
        {
            if (nearest[slot] is { } matchInGap)
                closest = matchInGap.Index;
            if (positions[ordered[slot]] < 0)
                positions[ordered[slot]] = closest;
        }
        return positions;
    }

    private sealed class LookupKeyMap<T>(string localeName)
    {
        private readonly Dictionary<object, T> values = [];
        // Windows can compare two strings as equal even when it generates different sort keys.
        private readonly SortedDictionary<(int Symbols, LookupKey Key), T> text = new(
            Comparer<(int Symbols, LookupKey Key)>.Create((left, right) =>
            {
                int order = left.Symbols.CompareTo(right.Symbols);
                return order != 0 ? order : left.Key.CompareTo(right.Key, localeName);
            }));

        public int Count => values.Count + text.Count;
        public IEnumerable<LookupKey> Keys => values.Keys.Select(value => new LookupKey(value, localeName))
            .Concat(text.Keys.Select(entry => entry.Key));

        public bool AddIfMissing(LookupKey key, T value, ref long budget, bool countNormalizedText = true)
        {
            if (TryGetValue(key, out _))
                return true;
            long entryBytes = key.Value is string ? 112 : DictionaryEntryBytes;
            if (!Reserve(ref budget, entryBytes + (countNormalizedText ? key.AdditionalTextBytes : 0)))
                return false;
            this[key] = value;
            return true;
        }

        public bool TryGetValue(LookupKey key, out T value) => key.Value is string
            ? text.TryGetValue((key.GetTextSymbolHashCode(), key), out value!)
            : values.TryGetValue(key.Value, out value!);

        public T this[LookupKey key]
        {
            get => key.Value is string ? text[(key.GetTextSymbolHashCode(), key)] : values[key.Value];
            set
            {
                if (key.Value is string)
                    text[(key.GetTextSymbolHashCode(), key)] = value;
                else
                    values[key.Value] = value;
            }
        }
    }

    private static int Find(ExcelArray lookup, LookupKey key, int length, bool vertical, int match, int search,
        string localeName)
    {
        if (search is 2 or -2)
            return FindBinary(lookup, key, length, vertical, match, search == 2, localeName);

        var wildcard = match == 2 && key.Value is string pattern ? new ExcelWildcard(pattern) : null;
        HashSet<string>? checkedTexts = wildcard is null ? null : new(StringComparer.Ordinal);
        bool skipText = match is 0 or 2 && key.Value is not string;
        int remainingWork = WildcardWorkLimit;

        int bestIndex = -1;
        LookupKey best = default;
        foreach (var (index, value) in lookup.Scan(0, length, vertical, search == 1, probeFirstCell: true))
        {
            if (wildcard is not null)
            {
                if (value is string text)
                {
                    if (!SpendWildcardScanWork(text, checkedTexts!, ref remainingWork, out bool alreadyChecked))
                        return PositionBudgetExceeded;
                    if (alreadyChecked)
                        continue;
                    if (wildcard.MatchesFolded(WindowsText.FoldWildcardText(text), ref remainingWork))
                        return index;
                    if (remainingWork < 0)
                        return PositionBudgetExceeded;
                }
                continue;
            }

            if (skipText && value is string)
                continue;
            var candidate = new LookupKey(value, localeName);
            int order = candidate.CompareTo(key, localeName);
            if (order == 0 && (value is string || Equals(candidate.Value, key.Value)))
                return index;
            if (match is -1 or 1 && Math.Sign(order) == match &&
                (bestIndex < 0 || Math.Sign(best.CompareTo(candidate, localeName)) == match))
            {
                bestIndex = index;
                best = candidate;
            }
        }
        return bestIndex;
    }

    private static int FindBinary(ExcelArray lookup, LookupKey key, int length, bool vertical, int match,
        bool ascending, string localeName, Func<int, object>? readProbe = null)
    {
        int exact = -1;
        int best = -1;
        int low = 0;
        int high = length - 1;
        while (low <= high)
        {
            int middle = low + (high - low) / 2;
            object value = readProbe is null ? lookup.At(middle, vertical) : readProbe(middle);
            var candidate = new LookupKey(value, localeName);
            int order = candidate.CompareTo(key, localeName);
            if (order == 0)
            {
                if (value is not string && !Equals(candidate.Value, key.Value))
                    return -1;
                exact = middle;
                if (ascending)
                    high = middle - 1;
                else
                    low = middle + 1;
                continue;
            }

            if (Math.Sign(order) == match)
                best = middle;

            if (!ascending)
                order = -order;
            if (order < 0)
                low = middle + 1;
            else
                high = middle - 1;
        }
        return exact >= 0 ? exact : best;
    }

    private readonly struct LookupKey(object value, string localeName)
    {
        public object Value { get; } = BlankLookupValue(value);
        private ExcelTextOrder Text { get; } = value is string text ? new ExcelTextOrder(text, localeName) : default;

        public int GetTextSymbolHashCode() => Text.GetSymbolHashCode();

        public long AdditionalTextBytes => Value is string text ? Text.AdditionalTextBytes(text) : 0;

        public long RetainedValueBytes => Value is string text ? ExcelTextOrder.StringBytes(text) + AdditionalTextBytes : 24;

        public int CompareTo(LookupKey other, string localeName) => (Value, other.Value) switch
        {
            (double a, double b) => a.CompareTo(b),
            (string, string) => Text.CompareTo(other.Text, localeName),
            (bool a, bool b) => a.CompareTo(b),
            (ExcelError a, ExcelError b) => ((int)a).CompareTo((int)b),
            _ => TypeOrder(Value) - TypeOrder(other.Value)
        };

        private static int TypeOrder(object value) => value switch
        {
            double => 0,
            string => 1,
            bool => 2,
            ExcelEmpty => 4,
            _ => 3
        };
    }

    private static bool TryValidateModes(double match, double search, out (int Match, int Search) modes)
    {
        match = Math.Floor(match);
        search = Math.Floor(search);
        if (match is not (-1 or 0 or 1 or 2) || search is not (-2 or -1 or 1 or 2) || (match == 2 && search is 2 or -2))
        {
            modes = default;
            return false;
        }

        modes = ((int)match, (int)search);
        return true;
    }

    private static double ReadMode(object value, int defaultMode) => value switch
    {
        ExcelMissing => defaultMode,
        ExcelEmpty => 0,
        double number => number,
        bool flag => flag ? 1 : 0,
        string text when XlCall.TryExcel(XlCall.xlfValue, out object parsed, text) == XlCall.XlReturn.XlReturnSuccess &&
            parsed is double number => number,
        _ => double.NaN
    };

    private static object ValueOrZero(object value) => value is ExcelEmpty ? Zero : value;

    private static object BlankLookupValue(object value) => value is ExcelMissing ? ExcelEmpty.Value : value;

    private readonly struct ExcelArray
    {
        private readonly object source;
        private readonly Func<ExcelReference, object> readBlock;

        public ExcelArray(object source, Func<ExcelReference, object> readBlock)
        {
            this.source = source;
            this.readBlock = readBlock;
            (Rows, Columns) = source switch
            {
                ExcelReference reference when reference.InnerReferences.Count == 1 =>
                    (reference.RowLast - reference.RowFirst + 1, reference.ColumnLast - reference.ColumnFirst + 1),
                ExcelReference or ExcelMissing => (0, 0),
                object[,] values => (values.GetLength(0), values.GetLength(1)),
                _ => (1, 1)
            };
        }

        public int Rows { get; }
        public int Columns { get; }
        public bool IsValid => Rows > 0 && Columns > 0;
        public bool IsReference => source is ExcelReference;
        public bool IsDisjointReference => !IsValid && source is ExcelReference { InnerReferences.Count: > 1 };

        public bool HasNonErrorValue() => source is object[,] values
            ? values.Cast<object>().Any(static value => value is not ExcelError)
            : source is not ExcelError;

        public object At(int index, bool vertical) => vertical ? At(index, 0) : At(0, index);

        public object? BroadcastAt(int row, int column)
        {
            row = Rows == 1 ? 0 : row;
            column = Columns == 1 ? 0 : column;
            return row < Rows && column < Columns ? At(row, column) : null;
        }

        public object At(int row, int column) => source switch
        {
            object[,] values => values[row, column],
            ExcelReference => ReadReferenceBlock(row, column, 1, 1).At(0, 0),
            _ => source
        };

        public object ReadValues(ref long remainingBytes, object? replaceNonErrors = null)
        {
            // Allow for the array slot and the value's managed object overhead.
            remainingBytes -= (long)Rows * Columns * WorksheetCellWorkingBytes;
            if (remainingBytes < 0)
                return ExcelArrayResult.TooLarge;

            var result = new object[Rows, Columns];
            int columnsPerBlock = Math.Min(Columns, WorksheetBlockSize);
            int rowsPerBlock = WorksheetBlockSize / columnsPerBlock;
            for (int row = 0; row < Rows; row += rowsPerBlock)
            {
                int rows = Math.Min(rowsPerBlock, Rows - row);
                for (int column = 0; column < Columns; column += columnsPerBlock)
                {
                    int columns = Math.Min(columnsPerBlock, Columns - column);
                    var values = ReadReferenceBlock(row, column, rows, columns);
                    if (values.Rows != rows || values.Columns != columns)
                        return values.source is ExcelError ? values.source : InvalidValue;

                    for (int i = 0; i < rows; i++)
                        for (int j = 0; j < columns; j++)
                        {
                            object value = values.At(i, j);
                            if (replaceNonErrors is not null && value is not ExcelError)
                                value = replaceNonErrors;
                            if (value is string text && (remainingBytes -= (long)text.Length * sizeof(char)) < 0)
                                return ExcelArrayResult.TooLarge;
                            result[row + i, column + j] = value;
                        }
                }
            }
            return result;
        }

        public Func<int, object> CacheProbes(int length, bool vertical, Func<long, bool> tryAdjustBudget)
        {
            const int blockSize = 64;
            // A full cache holds at most 64 MiB of worksheet text, even for maximum-length cells.
            const int capacity = 16;
            var lookup = this;
            var blocks = new Dictionary<int, (ExcelArray Values, int First, long Bytes, LinkedListNode<int> Recency)>(capacity);
            var recency = new LinkedList<int>();
            return index =>
            {
                int start = index / blockSize * blockSize;
                int first = index;
                int count = 1;
                if (blocks.TryGetValue(start, out var cached))
                {
                    if (cached.First == index || cached.Values.Rows > 1 || cached.Values.Columns > 1)
                    {
                        recency.Remove(cached.Recency);
                        recency.AddLast(cached.Recency);
                        return cached.Values.At(index - cached.First, vertical);
                    }

                    // A second cell in the same block makes prefetching useful.
                    Evict(cached.Recency);
                    first = start;
                    count = Math.Min(blockSize, length - start);
                }

                long bytes = DictionaryEntryBytes + count * WorksheetCellWorkingBytes;
                // Reserve the worst case before Excel copies the block.
                long maxBytes = bytes + count * (long)ExcelArrayResult.MaxCellTextBytes;
                while (blocks.Count == capacity || !tryAdjustBudget(maxBytes))
                {
                    if (recency.First is not { } oldest)
                        return lookup.At(index, vertical);
                    Evict(oldest);
                }

                var values = lookup.ReadReferenceBlock(vertical ? first : 0, vertical ? 0 : first,
                    vertical ? count : 1, vertical ? 1 : count);
                for (int offset = 0; offset < count; offset++)
                    if (values.At(offset, vertical) is string text)
                        bytes += text.Length * (long)sizeof(char);
                if (tryAdjustBudget(bytes - maxBytes))
                    blocks.Add(start, (values, first, bytes, recency.AddLast(start)));
                else
                    tryAdjustBudget(-maxBytes);

                return values.At(index - first, vertical);
            };

            void Evict(LinkedListNode<int> entry)
            {
                tryAdjustBudget(-blocks[entry.Value].Bytes);
                blocks.Remove(entry.Value);
                recency.Remove(entry);
            }
        }

        public bool ReadMatches(Dictionary<int, (int Count, object? Value)> matches, bool vertical, ref long textBudget)
        {
            int[] indices = matches.Keys.Order().ToArray();
            for (int start = 0; start < indices.Length;)
            {
                int first = indices[start];
                int end = start + 1;
                // Read across up to two unused cells to keep regularly spaced matches batched.
                while (end < indices.Length && indices[end] - first < WorksheetBlockSize &&
                    indices[end] - indices[end - 1] <= 3)
                    end++;

                int cells = indices[end - 1] - first + 1;
                var values = ReadReferenceBlock(vertical ? first : 0, vertical ? 0 : first,
                    vertical ? cells : 1, vertical ? 1 : cells);
                for (; start < end; start++)
                {
                    int index = indices[start];
                    object value = values.At(index - first, vertical);
                    int count = matches[index].Count;
                    textBudget -= (long)count * ExcelArrayResult.TextBytes(value);
                    if (textBudget < 0)
                        return false;
                    matches[index] = (count, value);
                }
            }
            return true;
        }

        public bool ReadMissingFallbacks(object[,] results, ref long textBudget)
        {
            // Follow the longer axis so tall, narrow ranges can batch reads across rows.
            bool vertical = Rows > Columns;
            int vectors = vertical ? Columns : Rows;
            int length = vertical ? Rows : Columns;
            for (int slice = 0; slice < vectors; slice++)
            {
                ExcelArray? vector = null;
                for (int start = 0; start < length;)
                {
                    if (!IsMissing(slice, start))
                    {
                        start++;
                        continue;
                    }

                    int last = start;
                    for (int next = start + 1; next < length && next - start < WorksheetBlockSize && next - last <= 2; next++)
                        if (IsMissing(slice, next))
                            last = next;
                    vector ??= vectors == 1 ? this : new ExcelArray(Slice((ExcelReference)source,
                        vertical ? 0 : slice, vertical ? slice : 0, vertical ? length : 1, vertical ? 1 : length), readBlock);
                    foreach (var (index, value) in vector.Value.Scan(start, last + 1, vertical, forward: true))
                    {
                        if (!IsMissing(slice, index))
                            continue;
                        textBudget -= ExcelArrayResult.TextBytes(value);
                        if (textBudget < 0)
                            return false;
                        results[vertical ? index : slice, vertical ? slice : index] = ValueOrZero(value);
                    }
                    start = last + 1;
                }
            }
            return true;

            bool IsMissing(int slice, int index) => results[vertical ? index : slice, vertical ? slice : index] is null;
        }

        public IEnumerable<(int Index, object Value)> Scan(int first, int end, bool vertical, bool forward,
            bool probeFirstCell = false)
        {
            int step = forward ? 1 : -1;
            int blockSize = probeFirstCell ? 1 : WorksheetBlockSize;
            while (first < end)
            {
                // Bound the amount of worksheet data copied at once, including in 32-bit Excel.
                int count = IsReference ? Math.Min(end - first, blockSize) : end - first;
                blockSize = WorksheetBlockSize;
                int start = forward ? first : end - count;
                int stop = start + count;
                var values = this;
                int origin = 0;
                if (IsReference)
                {
                    values = ReadReferenceBlock(vertical ? start : 0, vertical ? 0 : start,
                        vertical ? count : 1, vertical ? 1 : count);
                    origin = start;
                }

                int index = forward ? start : stop - 1;
                int sentinel = forward ? stop : start - 1;
                for (; index != sentinel; index += step)
                    yield return (index, values.At(index - origin, vertical));

                if (forward)
                    first = stop;
                else
                    end = start;
            }
        }

        public object GetResult(int index, bool vertical)
        {
            int row = vertical ? index : 0;
            int column = vertical ? 0 : index;
            int rows = vertical ? 1 : Rows;
            int columns = vertical ? Columns : 1;
            if (source is ExcelReference reference)
                return Slice(reference, row, column, rows, columns);
            if (source is not object[,] values)
                return ValueOrZero(source);
            if (rows == 1 && columns == 1)
                return ValueOrZero(values[row, column]);

            var result = new object[rows, columns];
            for (int i = 0; i < rows; i++)
                for (int j = 0; j < columns; j++)
                    result[i, j] = values[row + i, column + j];
            return result;
        }

        private ExcelArray ReadReferenceBlock(int row, int column, int rows, int columns) =>
            new(readBlock(Slice((ExcelReference)source, row, column, rows, columns)), readBlock);

        private static ExcelReference Slice(ExcelReference reference, int row, int column, int rows, int columns)
        {
            row += reference.RowFirst;
            column += reference.ColumnFirst;
            return new ExcelReference(row, row + rows - 1, column, column + columns - 1, reference.SheetId);
        }
    }
}
