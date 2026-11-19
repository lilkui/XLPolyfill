using System.Globalization;
using ExcelDna.Integration;
using Xunit;

namespace XLPolyfill.Tests;

public class LookupFunctionsTests
{
    private static object Lookup(object key, object lookup, object result, object? fallback = null,
        double match = 0, double search = 1) =>
        LookupFunctions.XLookup(key, lookup, result, fallback ?? ExcelMissing.Value, match, search);

    [Fact]
    public void DefaultsFindFirstExactMatchIgnoringCase()
    {
        object result = LookupFunctions.XLookup("KEY", new object[,] { { "key" }, { "Key" } },
            new object[,] { { "first" }, { "last" } }, ExcelError.ExcelErrorDiv0,
            ExcelMissing.Value, ExcelMissing.Value);

        Assert.Equal("first", result);
    }

    [Fact]
    public void ExactMatchKeepsNumbersTextAndBooleansDistinct()
    {
        var keys = new object[,] { { "1", true, 1d } };
        var values = new object[,] { { "text", "logical", "number" } };

        Assert.Equal("text", Lookup("1", keys, values));
        Assert.Equal("logical", Lookup(true, keys, values));
        Assert.Equal("number", Lookup(1d, keys, values));
        Assert.Equal(ExcelError.ExcelErrorNA, Lookup(false, keys, values));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    public void NumericLookupsSkipNormalizingTextCandidates(bool batched, double match)
    {
        const int rows = 2_048;
        object keys = batched ? new object[,] { { 2d, 3d } } : 2d;
        string prefix = new('\uFF11', 512);
        var lookup = new object[rows, 1];
        var returns = new object[rows, 1];
        for (int row = 0; row < rows - 1; row++)
            lookup[row, 0] = prefix + row.ToString(CultureInfo.InvariantCulture);
        lookup[rows - 1, 0] = 2d;
        returns[rows - 1, 0] = "found";

        _ = Lookup(keys, lookup, returns, "missing", match);
        long before = GC.GetAllocatedBytesForCurrentThread();
        object result = Lookup(keys, lookup, returns, "missing", match);
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        if (batched)
            Assert.Equal(new object[,] { { "found", "missing" } }, result);
        else
            Assert.Equal("found", result);
        Assert.InRange(allocated, 0, 512 * 1024);
    }

    [Fact]
    public void ReverseSearchFindsLastDuplicateAndKeepsApproximateTies()
    {
        var keys = new object[,] { { 2d }, { 5d }, { 2d }, { 1d } };
        var values = new object[,] { { "first" }, { "big" }, { "last" }, { "small" } };

        Assert.Equal("last", Lookup(2d, keys, values, search: -1));
        Assert.Equal("last", Lookup(3d, keys, values, match: -1, search: -1));
        Assert.Equal("big", Lookup(3d, keys, values, match: 1));
    }

    [Theory]
    [InlineData(1, 30d)]
    [InlineData(-1, 40d)]
    [InlineData(2, 30d)]
    [InlineData(-2, 40d)]
    public void NextLargerLookupCanMatchAnErrorCell(double search, double expected)
    {
        var lookup = search == -2
            ? new object[,] { { ExcelError.ExcelErrorNA, ExcelError.ExcelErrorDiv0, ExcelError.ExcelErrorDiv0, 5d, 1d } }
            : new object[,] { { 1d, 5d, ExcelError.ExcelErrorDiv0, ExcelError.ExcelErrorDiv0, ExcelError.ExcelErrorNA } };
        var values = search == -2 ? new object[,] { { 50d, 30d, 40d, 20d, 10d } }
            : new object[,] { { 10d, 20d, 30d, 40d, 50d } };

        Assert.Equal(expected, Lookup(6d, lookup, values, "missing", match: 1, search: search));
        Assert.Equal(new object[,] { { 20d, expected } },
            Lookup(new object[,] { { 2d, 6d } }, lookup, values, "missing", match: 1, search: search));
        Assert.Equal("missing", Lookup(6d, lookup, values, "missing", search: search));
    }

    [Theory]
    [InlineData(1, "first blank")]
    [InlineData(-1, "last blank")]
    [InlineData(2, "first blank")]
    [InlineData(-2, "missing")]
    public void NextLargerLookupUsesBlanksInExcelSearchOrder(double search, string expected)
    {
        double first = search == -2 ? 3d : 1d;
        double second = search == -2 ? 1d : 3d;
        var lookup = new object[,] { { first }, { second }, { ExcelEmpty.Value }, { ExcelEmpty.Value } };
        var values = new object[,] { { first }, { second }, { "first blank" }, { "last blank" } };

        Assert.Equal(expected, Lookup(4d, lookup, values, "missing", match: 1, search: search));
        Assert.Equal(new object[,] { { expected, 3d } },
            Lookup(new object[,] { { 4d, 2d } }, lookup, values, "missing", match: 1, search: search));
        Assert.Equal("missing", Lookup(4d, lookup, values, "missing", search: search));
    }

    [Theory]
    [InlineData(7, 0, 7)]
    [InlineData(6, -1, 4)]
    [InlineData(6, 1, 7)]
    [InlineData(0, 1, 1)]
    [InlineData(10, -1, 9)]
    [InlineData(0, -1, -99)]
    [InlineData(10, 1, -99)]
    [InlineData(6, 0, -99)]
    public void BinarySearchFindsExpectedValueInBothSortOrders(double key, double match, double expected)
    {
        var ascending = new object[,] { { 1d }, { 4d }, { 7d }, { 9d } };
        var descending = new object[,] { { 9d, 7d, 4d, 1d } };

        Assert.Equal(expected, Lookup(key, ascending, ascending, -99d, match, 2));
        Assert.Equal(expected, Lookup(key, descending, descending, -99d, match, -2));
    }

    [Theory]
    [InlineData(2, -1, 21)]
    [InlineData(2, 0, 21)]
    [InlineData(2, 1, 21)]
    [InlineData(-2, -1, 23)]
    [InlineData(-2, 0, 23)]
    [InlineData(-2, 1, 23)]
    public void BinarySearchPicksTheExcelDuplicateInScalarAndBatch(double search, double match, double expected)
    {
        var keys = new object[,] { { search == 2 ? 1d : 4d }, { 3d }, { 3d }, { 3d }, { search == 2 ? 4d : 1d } };
        var values = new object[,] { { 11d }, { 21d }, { 22d }, { 23d }, { 33d } };

        Assert.Equal(expected, Lookup(3d, keys, values, "missing", match, search));
        Assert.Equal(new object[,] { { expected }, { expected } },
            Lookup(new object[,] { { 3d }, { 3d } }, keys, values, "missing", match, search));
    }

    [Theory]
    [InlineData(0.9, 1.9, -99d)]
    [InlineData(-0.9, 1.9, 11d)]
    [InlineData(1.9, 2.9, 33d)]
    [InlineData(1.9, 0.9, null)]
    [InlineData(-1.9, 1.9, null)]
    public void RoundsFractionalModesDownBeforeValidation(double match, double search, double? expected)
    {
        var keys = new object[,] { { 1d }, { 3d } };
        var values = new object[,] { { 11d }, { 33d } };

        Assert.Equal(expected is double number ? (object)number : ExcelError.ExcelErrorValue,
            Lookup(2d, keys, values, -99d, match, search));
    }

    [Fact]
    public void ModeArraysBroadcastAcrossKeysAndReturnTheLeadingCell()
    {
        var keys = new object[,] { { 1d }, { 3d }, { 2d } };
        var lookup = new object[,] { { 1d }, { 1d }, { 3d } };
        var returns = new object[,] { { 10d, "first" }, { 11d, "last" }, { 30d, "third" } };
        var fallback = new object[,] { { 90d }, { 80d }, { 70d } };
        var searches = new object[,] { { 1d, -1d } };

        Assert.Equal(new object[,] { { 10d, 11d }, { 30d, 30d }, { 90d, 90d } },
            LookupFunctions.XLookup(keys, lookup, returns, fallback, ExcelMissing.Value, searches));
        Assert.Equal(new object[,] { { 11d } },
            LookupFunctions.XLookup(1d, lookup, returns, fallback, 0d, new object[,] { { -1d } }));
    }

    [Fact]
    public void ModeArraysPreserveErrorsAndFillNonoverlappingCells()
    {
        var lookup = new object[,] { { 1d, 3d } };
        var returns = new object[,] { { 10d, 30d } };
        var matches = new object[,] { { -1d, 1d, 2d, 0d, 0d } };
        var searches = new object[,] { { 1d, -1d, 2d, ExcelError.ExcelErrorRef } };

        Assert.Equal(new object[,] {
            { 10d, 30d, ExcelError.ExcelErrorValue, ExcelError.ExcelErrorRef, ExcelError.ExcelErrorNA }
        }, LookupFunctions.XLookup(2d, lookup, returns, 0d, matches, searches));
    }

    [Fact]
    public void ModeArraysPreserveEarlierArgumentErrorsOutsideTheSearchModeShape()
    {
        var lookup = new object[,] { { 1d }, { 2d } };
        var keys = new object[,] { { 1d }, { 2d }, { ExcelError.ExcelErrorDiv0 }, { 4d } };
        var match = new object[,] { { 0d }, { 0d }, { 0d }, { ExcelError.ExcelErrorNum } };
        var search = new object[,] { { 1d }, { -1d } };

        Assert.Equal(new object[,] { { 1d }, { 2d }, { ExcelError.ExcelErrorDiv0 }, { ExcelError.ExcelErrorNum } },
            LookupFunctions.XLookup(keys, lookup, lookup, 0d, match, search));
    }

    [Theory]
    [InlineData(2_100, false, false)]
    [InlineData(8_300, true, true)]
    public void ModeArraysBroadcastAScalarKeyErrorWithoutReadingModes(int rows, bool is64BitProcess, bool keyIsReference)
    {
        var sheet = new IntPtr(42);
        var keyCell = new ExcelReference(0, 0, 0, 0, sheet);
        var searchModes = new ExcelReference(0, rows - 1, 1, 1, sheet);
        var matchModes = new ExcelReference(0, 0, 2, 3, sheet);
        const ExcelError error = ExcelError.ExcelErrorNA;
        var reads = new List<ExcelReference>();

        object result = LookupFunctions.XLookup(keyIsReference ? keyCell : error, 1d, 10d,
            ExcelMissing.Value, matchModes, searchModes, ReadBlock, is64BitProcess);

        var cells = Assert.IsType<object[,]>(result);
        Assert.Equal((rows, 2), (cells.GetLength(0), cells.GetLength(1)));
        Assert.All(cells.Cast<object>(), cell => Assert.Equal(error, cell));
        ExcelReference[] expectedReads = keyIsReference ? [keyCell] : [];
        Assert.Equal(expectedReads, reads);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            return error;
        }
    }

    [Fact]
    public void ModeArraysSkipUnusedWorksheetReadsWhenEveryKeyOrMatchIsAnError()
    {
        const ExcelError first = ExcelError.ExcelErrorNA;
        const ExcelError second = ExcelError.ExcelErrorDiv0;
        var errors = new object[,] { { first }, { second } };
        var modeRange = new ExcelReference(0, 0, 1, 2, new IntPtr(42));
        var expected = new object[,] { { first, first }, { second, second } };

        Assert.Equal(expected, LookupFunctions.XLookup(errors, 1d, 10d, ExcelMissing.Value,
            modeRange, 1d, UnexpectedRead, is64BitProcess: true));
        Assert.Equal(expected, LookupFunctions.XLookup(1d, 1d, 10d, ExcelMissing.Value,
            errors, modeRange, UnexpectedRead, is64BitProcess: true));
    }

    [Fact]
    public void DisjointReferencesAreScalarErrorsWhenUsedAsKeysOrModes()
    {
        var disjoint = new ExcelReference([[0, 1, 0, 0], [3, 3, 0, 0]], new IntPtr(42));
        const ExcelError error = ExcelError.ExcelErrorValue;

        Assert.Equal(error, LookupFunctions.XLookup(1d, 1d, 10d, "missing",
            disjoint, ExcelError.ExcelErrorDiv0, UnexpectedRead, is64BitProcess: true));
        Assert.Equal(error, LookupFunctions.XLookup(1d, 1d, 10d, "missing",
            0d, disjoint, UnexpectedRead, is64BitProcess: true));
        Assert.Equal(new object[,] { { error, error } }, LookupFunctions.XLookup(disjoint, 1d, 10d, "missing",
            0d, new object[,] { { 1d, -1d } }, UnexpectedRead, is64BitProcess: true));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DisjointModeReferencesPreserveTheBatchShape(bool disjointMatch)
    {
        var disjoint = new ExcelReference([[0, 0, 0, 0], [2, 2, 0, 0]], new IntPtr(42));
        var keys = new object[,] { { 1d }, { 2d } };
        var validModes = new object[,] { { 1d, 1d } };
        object match = disjointMatch ? disjoint : validModes;
        object search = disjointMatch ? validModes : disjoint;
        const ExcelError error = ExcelError.ExcelErrorValue;

        Assert.Equal(new object[,] { { error, error }, { error, error } },
            LookupFunctions.XLookup(keys, keys, keys, "missing", match, search, UnexpectedRead, is64BitProcess: true));
    }

    [Fact]
    public void ModeArraysReturnValueErrorForWorksheetFallbackOnlyOnMisses()
    {
        var keys = new object[,] { { 9d, 2d, 8d } };
        var lookup = new object[,] { { 1d, 2d } };
        var returns = new object[,] { { 10d, 20d } };
        var fallback = new ExcelReference(0, 0, 0, 0, new IntPtr(42));
        var searches = new object[,] { { 1d, -1d, 1d } };

        Assert.Equal(new object[,] { { ExcelError.ExcelErrorValue, 20d, ExcelError.ExcelErrorValue } },
            LookupFunctions.XLookup(keys, lookup, returns, fallback, 0d, searches, UnexpectedRead, is64BitProcess: true));
        Assert.Equal(new object[,] { { ExcelError.ExcelErrorValue } },
            LookupFunctions.XLookup(9d, lookup, returns, fallback, new object[,] { { 0d } }, 1d, UnexpectedRead,
                is64BitProcess: true));
        Assert.Equal(new object[,] { { "fallback", 20d, "fallback" } },
            LookupFunctions.XLookup(keys, lookup, returns, "fallback", 0d, searches));
    }

    [Fact]
    public void ModeArraysShareWorksheetReadsForKeysWithTheSameModes()
    {
        var keys = new object[,] { { -1d, -2d } };
        var lookup = new ExcelReference(0, 512, 0, 0, new IntPtr(42));
        var returns = new object[513, 1];
        var searches = new object[,] { { 1d, 1d } };
        int reads = 0;

        Assert.Equal(new object[,] { { "missing", "missing" } },
            LookupFunctions.XLookup(keys, lookup, returns, "missing", 0d, searches, ReadBlock, is64BitProcess: true));
        Assert.Equal(2, reads);

        object ReadBlock(ExcelReference block)
        {
            reads++;
            int rows = block.RowLast - block.RowFirst + 1;
            if (rows == 1)
                return 1d;

            var values = new object[rows, 1];
            for (int row = 0; row < rows; row++)
                values[row, 0] = 1d;
            return values;
        }
    }

    [Fact]
    public void ModeArraysRejectOversizedBroadcastBeforeReadingKeys()
    {
        var keys = new ExcelReference(0, 1_048_575, 0, 0, new IntPtr(42));
        var searches = new object[1, 64];

        Assert.Equal(ExcelError.ExcelErrorNum,
            LookupFunctions.XLookup(keys, 1d, 10d, "missing", 0d, searches, UnexpectedRead, is64BitProcess: true));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OversizedWorksheetModeRangesAreRejectedBeforeReading(bool is64BitProcess)
    {
        var modes = new ExcelReference(0, 1_048_575, 3, 26, new IntPtr(42));

        Assert.Equal(ExcelError.ExcelErrorNum,
            LookupFunctions.XLookup(1d, 1d, 10d, "missing", modes, 1d, UnexpectedRead, is64BitProcess));
        Assert.Equal(ExcelError.ExcelErrorNum,
            LookupFunctions.XLookup(1d, 1d, 10d, "missing", 0d, modes, UnexpectedRead, is64BitProcess));
    }

    [Fact]
    public void WorksheetModeRangesAreReadInBlocksAndKeepTheirShape()
    {
        var sheet = new IntPtr(42);
        var matches = new ExcelReference(0, 512, 0, 0, sheet);
        var searches = new ExcelReference(0, 0, 1, 1, sheet);
        var reads = new List<ExcelReference>();
        var lookup = new object[,] { { 1d, 3d } };
        var values = new object[,] { { 10d, 30d } };

        object result = LookupFunctions.XLookup(2d, lookup, values, "missing", matches, searches, ReadBlock,
            is64BitProcess: true);

        var cells = Assert.IsType<object[,]>(result);
        Assert.Equal(513, cells.GetLength(0));
        Assert.Equal(1, cells.GetLength(1));
        Assert.All(cells.Cast<object>().Take(512), cell => Assert.Equal(10d, cell));
        Assert.Equal(30d, cells[512, 0]);
        Assert.Equal(new ExcelReference[] { new(0, 511, 0, 0, sheet), new(512, 512, 0, 0, sheet), searches }, reads);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            if (block.ColumnFirst == 1 || block.RowFirst == 512)
                return 1d;
            var modes = new object[block.RowLast - block.RowFirst + 1, 1];
            for (int row = 0; row < modes.Length; row++)
                modes[row, 0] = -1d;
            return modes;
        }
    }

    [Fact]
    public void SingleCellModeReferenceKeepsScalarFallback()
    {
        var search = new ExcelReference(0, 0, 0, 0, new IntPtr(42));
        var fallback = new ExcelReference(0, 0, 1, 1, search.SheetId);

        Assert.Equal(fallback, LookupFunctions.XLookup(2d, 1d, 10d, fallback, 0d, search, block =>
        {
            Assert.Equal(search, block);
            return 1d;
        }, is64BitProcess: true));
    }

    [Fact]
    public void AscendingBinarySearchHandlesTrailingBlanks()
    {
        var keys = new object[,] { { 4d }, { 9d }, { ExcelEmpty.Value }, { ExcelEmpty.Value }, { ExcelEmpty.Value } };
        var values = new object[,] { { 4d }, { 9d }, { "blank" }, { "blank" }, { "blank" } };

        Assert.Equal(4d, Lookup(4d, keys, values, "missing", search: 2));
        Assert.Equal(9d, Lookup(8d, keys, values, "missing", match: 1, search: 2));
        Assert.Equal("blank", Lookup(ExcelMissing.Value, keys, values, "missing", search: 2));
    }

    [Fact]
    public void DescendingBinarySearchUsesExcelBlankProbeOrder()
    {
        var values = new object[,] { { 10d }, { 20d }, { 30d } };
        var trailingBlank = new object[,] { { 9d }, { 4d }, { ExcelEmpty.Value } };
        var leadingBlank = new object[,] { { ExcelEmpty.Value }, { 9d }, { 4d } };

        Assert.Equal("missing", Lookup(ExcelMissing.Value, trailingBlank, values, "missing", search: -2));
        Assert.Equal(30d, Lookup(0d, trailingBlank, values, "missing", match: 1, search: -2));
        Assert.Equal(10d, Lookup(ExcelMissing.Value, leadingBlank, values, "missing", search: -2));
        Assert.Equal(new object[,] { { "missing", 20d } },
            Lookup(new object[,] { { ExcelMissing.Value, 4d } }, trailingBlank, values, "missing", search: -2));
    }

    [Fact]
    public void AscendingBinarySearchFindsBlankAfterErrors()
    {
        var keys = new object[,] { { 1d, ExcelError.ExcelErrorNA, ExcelError.ExcelErrorRef, ExcelEmpty.Value } };
        var values = new object[,] { { "number", "error 1", "error 2", "blank" } };

        Assert.Equal("blank", Lookup(ExcelMissing.Value, keys, values, search: 2));
    }

    [Theory]
    [InlineData(1, 1d, 2d, 2d, 20d)]
    [InlineData(-1, 1d, 2d, 2d, 30d)]
    [InlineData(2, 1d, 2d, 2d, 30d)]
    [InlineData(-2, 2d, 2d, 1d, 10d)]
    public void NextSmallerBlankKeyReturnsTheExcelCandidate(double search, double first,
        double second, double third, double expected)
    {
        var lookup = new object[,] { { first, second, third } };
        var returns = new object[,] { { 10d, 20d, 30d } };
        var keys = new object[,] { { ExcelEmpty.Value, ExcelMissing.Value } };

        Assert.Equal(expected, Lookup(ExcelMissing.Value, lookup, returns, "missing", match: -1, search: search));
        Assert.Equal(new object[,] { { expected, expected } },
            Lookup(keys, lookup, returns, "missing", match: -1, search: search));
        Assert.Equal("missing", Lookup(ExcelMissing.Value, lookup, returns, "missing", match: 1, search: search));

        var lookupWithBlank = new object[,] { { first, second, third, ExcelEmpty.Value } };
        var returnsWithBlank = new object[,] { { 10d, 20d, 30d, "blank" } };
        object withBlank = search == -2 ? expected : "blank";
        Assert.Equal(withBlank, Lookup(ExcelMissing.Value, lookupWithBlank, returnsWithBlank, "missing", match: -1, search: search));
        Assert.Equal(new object[,] { { withBlank, withBlank } },
            Lookup(keys, lookupWithBlank, returnsWithBlank, "missing", match: -1, search: search));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(-2)]
    public void BinarySearchFollowsExcelTextSortOrder(double search)
    {
        var punctuationKeys = search == 2 ? new object[,] { { "A!", "A1" } } : new object[,] { { "A1", "A!" } };
        var keys = search == 2
            ? new object[,] { { "A!", "A1", "A2", "AA", "A-A" } }
            : new object[,] { { "A-A", "AA", "A2", "A1", "A!" } };

        Assert.Equal("A!", Lookup("a!", punctuationKeys, punctuationKeys, "missing", search: search));
        Assert.Equal("A1", Lookup("a1", punctuationKeys, punctuationKeys, "missing", search: search));
        Assert.Equal(new object[,] { { "A1", "A!" } },
            Lookup(new object[,] { { "a1", "a!" } }, punctuationKeys, punctuationKeys, "missing", search: search));
        Assert.Equal("A2", Lookup("a3", keys, keys, "missing", match: -1, search: search));
        Assert.Equal("AA", Lookup("a3", keys, keys, "missing", match: 1, search: search));
        Assert.Equal("A-A", Lookup("a-a", keys, keys, "missing", search: search));
    }

    [Fact]
    public void BinarySearchDistinguishesApostrophePositions()
    {
        var keys = new object[,] { { "BS", "B's", "b's", "B-T" } };

        Assert.Equal("BS", Lookup("bs", keys, keys, "missing", search: 2));
        Assert.Equal("BS", Lookup("bs'", keys, keys, "missing", match: -1, search: 2));
        Assert.Equal("B's", Lookup("bs'", keys, keys, "missing", match: 1, search: 2));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ApproximateTextMatchesUseTheSameOrderAcrossSearchModes(double search)
    {
        var keys = search == -2 ? new object[,] { { "A2", "A1", "A!" } } : new object[,] { { "A!", "A1", "A2" } };

        Assert.Equal("missing", Lookup("A3", keys, keys, "missing", match: 1, search: search));
        Assert.Equal("A2", Lookup("A3", keys, keys, "missing", match: -1, search: search));
        Assert.Equal("A1", Lookup("A0", keys, keys, "missing", match: 1, search: search));
        Assert.Equal("A!", Lookup("A0", keys, keys, "missing", match: -1, search: search));
        Assert.Equal(new object[,] { { "A!", "A2" } },
            Lookup(new object[,] { { "A0", "A3" } }, keys, keys, "missing", match: -1, search: search));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void TextSearchIgnoresSoftHyphensInKeysAndLookupValues(double search)
    {
        const string withSoftHyphen = "1\u00AD1";
        var keys = search == -2 ? new object[,] { { "13", "12", "11" } } : new object[,] { { "11", "12", "13" } };

        Assert.Equal("11", Lookup(withSoftHyphen, keys, keys, "missing", search: search));
        keys[0, search == -2 ? 2 : 0] = withSoftHyphen;
        Assert.Equal(withSoftHyphen, Lookup("11", keys, keys, "missing", search: search));
    }

    [Theory]
    [InlineData("\u00AD")]
    [InlineData("\uFEFF")]
    [InlineData("\u200D")]
    public void ExactLookupDistinguishesEmptyTextFromIgnoredCharacters(string ignored)
    {
        var lookup = new object[,] { { ignored, "" } };
        var values = new object[,] { { 10d, 20d } };

        Assert.Equal(20d, Lookup("", lookup, values));
        Assert.Equal(10d, Lookup(ignored, lookup, values, search: -1));
        Assert.Equal(new object[,] { { 20d, 10d } }, Lookup(new object[,] { { "", ignored } }, lookup, values));
        Assert.Equal(10d, Lookup("", ignored, 10d, "missing", match: 1));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void GreekQuestionMarkMatchesSemicolon(double match)
    {
        var lookup = new object[,] { { "\u037E", ";" } };
        var values = new object[,] { { 10d, 20d } };

        Assert.Equal(10d, Lookup(";", lookup, values, match: match));
        Assert.Equal(20d, Lookup("\u037E", lookup, values, match: match, search: -1));
        Assert.Equal(new object[,] { { 10d, 10d } },
            Lookup(new object[,] { { ";", "\u037E" } }, lookup, values, match: match));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ExactTextMatchesDistinguishZeroWidthSpaceAndNormalizeWidthAndBom(double search)
    {
        var identifiers = search == -2 ? new object[,] { { "13", "12", "11" } } : new object[,] { { "11", "12", "13" } };

        Assert.Equal("missing", Lookup("1\u200B1", identifiers, identifiers, "missing", search: search));
        Assert.Equal("11", Lookup("\uFEFF11", identifiers, identifiers, "missing", search: search));
        Assert.Equal("11", Lookup("１１", identifiers, identifiers, "missing", search: search));

        var distinct = new object[,] { { "ACME", "ACME\u200B" } };
        var reversed = new object[,] { { "ACME\u200B", "ACME" } };
        var cells = search == -2 ? reversed : distinct;
        Assert.Equal("ACME\u200B", Lookup("ACME\u200B", cells, cells, "missing", search: search));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(-2)]
    public void BinarySearchFindsKeysInExcelSortedPunctuationOrder(double search)
    {
        var keys = search == 2
            ? new object[,] { { "b", "b'", "b-", "B--", "'b", "-b" } }
            : new object[,] { { "-b", "'b", "B--", "b-", "b'", "b" } };

        Assert.Equal("'b", Lookup("'B", keys, keys, "missing", search: search));
        Assert.Equal(new object[,] { { "'b", "b-", "-b" } },
            Lookup(new object[,] { { "'B", "B-", "-B" } }, keys, keys, "missing", search: search));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ApproximateLookupsUsePunctuationPosition(double search)
    {
        var keys = search == -2 ? new object[,] { { "c", "b-" } } : new object[,] { { "b-", "c" } };
        Assert.Equal("b-", Lookup("-b", keys, keys, "missing", match: -1, search: search));
        Assert.Equal(new object[,] { { "b-", "c" } },
            Lookup(new object[,] { { "-b", "c" } }, keys, keys, "missing", match: -1, search: search));

        keys = search == -2 ? new object[,] { { "c", "-b" } } : new object[,] { { "-b", "c" } };
        Assert.Equal("-b", Lookup("b-", keys, keys, "missing", match: 1, search: search));
    }

    [Theory]
    [InlineData(2, 0, "missing")]
    [InlineData(2, -1, "ACME")]
    [InlineData(-2, 1, "AC'ME")]
    public void BinaryBatchProbesLargePunctuationGroupsWithoutScanning(double search, double match, string missing)
    {
        const int length = 1_033;
        var lookup = new ExcelReference(0, length - 1, 0, 0, new IntPtr(42));
        var worksheet = Enumerable.Repeat("ACME", length).ToArray();
        worksheet[0] = "AB";
        worksheet[^3] = "AC'ME";
        worksheet[^2] = "A'CME";
        worksheet[^1] = "AD";
        if (search == -2)
            Array.Reverse(worksheet);
        var values = new object[length, 1];
        for (int i = 0; i < length; i++)
            values[i, 0] = worksheet[i];

        var keys = new object[,] { { "a'cme", "ac'me", "acm'e", "ACME" } };
        int reads = 0;
        Assert.Equal(new object[,] { { "A'CME", "AC'ME", missing, "ACME" } },
            LookupFunctions.XLookup(keys, lookup, values, "missing", match, search, ReadBlock, true));
        Assert.InRange(reads, 1, 44);

        object ReadBlock(ExcelReference block)
        {
            reads++;
            int count = block.RowLast - block.RowFirst + 1;
            Assert.InRange(count, 1, 64);
            if (count == 1)
                return worksheet[block.RowFirst];

            var values = new object[count, 1];
            for (int i = 0; i < count; i++)
                values[i, 0] = worksheet[block.RowFirst + i];
            return values;
        }
    }

    [Fact]
    public void BinaryBatchReusesRepeatedTextKeysInLargeRanges()
    {
        var keys = new ExcelReference(0, 2_041, 0, 0, new IntPtr(42));
        var table = new object[65_536, 1];
        for (int row = 0; row < table.Length; row++)
            table[row, 0] = "ab";
        string longKey = new('z', 32_767);

        var result = Assert.IsType<object[,]>(
            LookupFunctions.XLookup(keys, table, table, "missing", 0d, 2d, ReadKeys, is64BitProcess: false));

        Assert.Equal(keys.RowLast + 1, result.Length);
        Assert.Equal("ab", result[1, 0]);
        for (int row = 0; row < result.Length; row++)
            if (row != 1)
                Assert.Equal("missing", result[row, 0]);

        object ReadKeys(ExcelReference block)
        {
            var values = new object[block.RowLast - block.RowFirst + 1, 1];
            for (int row = 0; row < values.Length; row++)
                values[row, 0] = (block.RowFirst + row) switch { 0 => "a'b", 1 => "ab", _ => longKey };
            return values;
        }
    }

    [Fact]
    public void BinaryBatchReusesAResultForPunctuatedCaseVariants()
    {
        const int length = 4_096;
        var lookup = new ExcelReference(0, length - 1, 0, 0, new IntPtr(42));
        var values = new object[length, 1];
        values[length - 1, 0] = "found";
        var keys = new object[,] { { "A'BC", "a'bc", "A'bC", "a'Bc" } };
        int reads = 0;

        Assert.Equal(new object[,] { { "found", "found", "found", "found" } },
            LookupFunctions.XLookup(keys, lookup, values, "missing", 0d, 2d, ReadBlock, is64BitProcess: true));
        Assert.InRange(reads, 1, 13);

        object ReadBlock(ExcelReference block)
        {
            Assert.Equal(block.RowFirst, block.RowLast);
            reads++;
            return block.RowFirst == length - 1 ? "A'BC" : "ABC";
        }
    }

    [Theory]
    [InlineData("part-~*", 1, "asterisk")]
    [InlineData("part-~?", 1, "question mark")]
    [InlineData("*~~", 1, "tilde")]
    [InlineData("part-?", -1, "last")]
    [InlineData("*a*?x", 1, "backtrack")]
    [InlineData("part", 1, "missing")]
    public void WildcardsMatchEntireTextAndHonorEscapes(string pattern, double search, string expected)
    {
        var keys = new object[,] { { "Part-*", "part-?", "part-~", "part-Z", "abacyx" } };
        var values = new object[,] { { "asterisk", "question mark", "tilde", "last", "backtrack" } };

        Assert.Equal(expected, Lookup(pattern, keys, values, "missing", 2, search));
    }

    [Theory]
    [InlineData("~a", "a", "~a")]
    [InlineData("a~b", "ab", "a~b")]
    [InlineData("ab~", "ab", "ab~")]
    [InlineData("ab*~", "abcd", "ab")]
    [InlineData("~", "", "~")]
    [InlineData("~~", "~", "")]
    public void WildcardEscapesWorkInScalarAndBatchedLookups(string pattern, string first, string second)
    {
        var lookup = new object[,] { { first, second } };
        var returns = new object[,] { { "first", "second" } };

        Assert.Equal("first", Lookup(pattern, lookup, returns, "missing", match: 2));
        Assert.Equal(new object[,] { { "first", "missing" } },
            Lookup(new object[,] { { pattern, "absent" } }, lookup, returns, "missing", match: 2));
    }

    [Theory]
    [InlineData(1, "empty text")]
    [InlineData(-1, "letter")]
    public void StandaloneControlCharacterMatchesAnyTextInWildcardMode(double search, string expected)
    {
        const string pattern = "\u0001";
        var lookup = new object[,] { { ExcelEmpty.Value, "", "a", 1d } };
        var returns = new object[,] { { "blank cell", "empty text", "letter", "number" } };

        Assert.Equal(expected, Lookup(pattern, lookup, returns, "missing", match: 2, search: search));
        Assert.Equal(new object[,] { { expected, "letter" } },
            Lookup(new object[,] { { pattern, "a" } }, lookup, returns, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData("µ", "Μ", "µ", "second")]
    [InlineData("μ", "µ", "Μ", "second")]
    [InlineData("ß", "ẞ", "ß", "first")]
    [InlineData("ẞ", "ß", "ẞ", "first")]
    [InlineData("-*", "﹣SKU", "－SKU", "first")]
    [InlineData("﹦*", "＝SKU", "=SKU", "second")]
    [InlineData("⁰*", "0 SKU", "⓪ SKU", "second")]
    [InlineData("¹*", "1 SKU", "➊ SKU", "second")]
    [InlineData("❶*", "1 SKU", "₁ SKU", "second")]
    [InlineData("6*", "➅ SKU", "➏ SKU", "first")]
    [InlineData("⓿*", "⓪ SKU", "0 SKU", "second")]
    [InlineData("K*", "kg", "Kg", "first")]
    [InlineData("Ω*", "Ω1", "Ω1", "first")]
    [InlineData("Å*", "å1", "Å1", "first")]
    [InlineData("a.b", "a\u2024b", "a.b", "first")]
    [InlineData("a\u2024*", "a.b", "a\u2024b", "first")]
    [InlineData("a.*", "a．b", "a.b", "second")]
    [InlineData("\u200C", "\u200D", "\u200C", "first")]
    [InlineData("a\u200D*", "a\u200Cb", "a\u200Db", "first")]
    [InlineData("\u03B9*", "\u0345-item", "\u03B9-item", "second")]
    [InlineData("\u0345*", "\u03B9-item", "\u0345-item", "second")]
    [InlineData("\U00010400*", "\U00010428-item", "\U00010400-item", "second")]
    [InlineData("*\U00010428", "item-\U00010400", "item-\U00010428", "second")]
    public void WildcardLiteralEquivalencesFollowExcel(string pattern, string first, string second, string expected)
    {
        var lookup = new object[,] { { first, second } };
        var returns = new object[,] { { "first", "second" } };

        Assert.Equal(expected, Lookup(pattern, lookup, returns, "missing", match: 2));
        Assert.Equal(new object[,] { { expected, "missing" } },
            Lookup(new object[,] { { pattern, "absent" } }, lookup, returns, "missing", match: 2));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void ExactLookupsTreatOneDotLeaderLikeAPeriod(double search)
    {
        var lookup = new object[,] { { "a\u2024b", "a.b", "a．b" } };
        var returns = new object[,] { { "first", "second", "third" } };

        Assert.Equal("first", Lookup("a.b", lookup, returns, "missing", search: search));
        Assert.Equal(new object[,] { { "first", "first", "first" } },
            Lookup(lookup, lookup, returns, "missing", search: search));
    }

    [Theory]
    [InlineData("⁰¹²³⁴⁵⁶⁷⁸⁹")]
    [InlineData("₀₁₂₃₄₅₆₇₈₉")]
    [InlineData("⓪①②③④⑤⑥⑦⑧⑨")]
    [InlineData("⓿❶❷❸❹❺❻❼❽❾")]
    [InlineData("⓿⓵⓶⓷⓸⓹⓺⓻⓼⓽")]
    [InlineData("⓿➀➁➂➃➄➅➆➇➈")]
    [InlineData("⓿➊➋➌➍➎➏➐➑➒")]
    public void ExactLookupsMatchEquivalentDigitForms(string variants)
    {
        for (int digit = 0; digit < variants.Length; digit++)
        {
            string key = $"Code-{digit}";
            var lookup = new object[,] { { $"Code-{variants[digit]}", key } };
            var returns = new object[,] { { "first", "second" } };

            Assert.Equal("first", Lookup(key, lookup, returns, "missing"));
            Assert.Equal(new object[,] { { "first", "missing" } },
                Lookup(new object[,] { { key, "absent" } }, lookup, returns, "missing", search: 2));
        }
    }

    [Theory]
    [InlineData("㎡", "m²", "first")]
    [InlineData("㎥", "m3", "first")]
    [InlineData("㎠", "cm²", "first")]
    [InlineData("㎤", "cm³", "first")]
    [InlineData("㎟", "mm2", "second")]
    [InlineData("㎦", "km3", "second")]
    public void ExactLookupsMatchEquivalentSquaredAndCubedUnits(string unit, string spelledOut, string expected)
    {
        var returns = new object[,] { { "first", "second" } };
        var lookup = new object[,] { { spelledOut, unit } };

        Assert.Equal(expected, Lookup(unit, lookup, returns, "missing"));
        Assert.Equal(new object[,] { { expected, "missing" } },
            Lookup(new object[,] { { unit, "absent" } }, lookup, returns, "missing"));
        Assert.Equal(expected == "first" ? "first" : "missing",
            Lookup(spelledOut, unit, "first", "missing", search: 2));
    }

    [Fact]
    public void WildcardMatchesPreserveExcelTextWidth()
    {
        Assert.Equal("missing", Lookup("１１", "11", "found", "missing", match: 2));
        Assert.Equal("missing", Lookup("A\u3000B", "A B", "found", "missing", match: 2));
        Assert.Equal("missing", Lookup("Ａ*", "ABC", "found", "missing", match: 2));
        Assert.Equal("missing", Lookup("Ａ＊", "ＡＢＣ", "found", "missing", match: 2));
        Assert.Equal("found", Lookup("Ａ＊", "Ａ＊", "found", "missing", match: 2));
    }

    [Theory]
    [InlineData("\uFFE9", "\u2190", false)]
    [InlineData("\uFFEA", "\u2191", true)]
    [InlineData("\uFFEB", "\u2192", false)]
    [InlineData("\uFFEC", "\u2193", true)]
    public void ExactLookupsFollowExcelArrowWidthRules(string narrow, string wide, bool equivalent)
    {
        var lookup = new object[,] { { wide, narrow } };
        var returns = new object[,] { { "wide", "narrow" } };
        string expected = equivalent ? "wide" : "narrow";

        Assert.Equal(expected, Lookup(narrow, lookup, returns));
        Assert.Equal(new object[,] { { "wide", expected } }, Lookup(lookup, lookup, returns, search: 2));
        Assert.Equal(new object[,] { { "wide", "narrow" } }, Lookup(lookup, lookup, returns, match: 2));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ExactLookupsNormalizeVoicedHalfWidthKana(double search)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ja-JP");
            var lookup = search == -2 ? new object[,] { { "ギ", "ｶﾞ" } } : new object[,] { { "ｶﾞ", "ギ" } };

            Assert.Equal("ｶﾞ", Lookup("ガ", lookup, lookup, search: search));
            Assert.Equal(new object[,] { { "ｶﾞ", "ギ", "missing" } },
                Lookup(new object[,] { { "ガ", "ｷﾞ", "カ" } }, lookup, lookup, "missing", search: search));
            Assert.Equal("missing", Lookup("ガ", lookup, lookup, "missing", match: 2));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ExactLookupsIgnoreSpaceWidth(double search)
    {
        const string key = "A\u3000B";
        var lookup = search == -2 ? new object[,] { { "A C", "A B" } } : new object[,] { { "A B", "A C" } };

        Assert.Equal("A B", Lookup(key, lookup, lookup, search: search));
        Assert.Equal(new object[,] { { "A B", "A C" } },
            Lookup(new object[,] { { key, "A C" } }, lookup, lookup, search: search));
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(0, true)]
    [InlineData(2, false)]
    [InlineData(2, true)]
    public void TextLookupTreatsMediumMathematicalSpaceAsSpace(double match, bool batched)
    {
        var lookup = new object[,] { { "A\u205FB", "A B" } };
        var returns = new object[,] { { 10d, 20d } };

        if (batched)
            Assert.Equal(new object[,] { { 10d, 10d } },
                Lookup(new object[,] { { "A B", "A\u205FB" } }, lookup, returns, match: match));
        else
            Assert.Equal(10d, Lookup("A B", lookup, returns, match: match));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void ScalarAndBatchExactLookupsUseWindowsCollation(double search)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            const string key = "\u0153uf";
            var lookup = new object[,] { { "oeuf" } };
            var values = new object[,] { { "found" } };

            Assert.Equal("found", Lookup(key, lookup, values, "missing", search: search));
            Assert.Equal(new object[,] { { "found", "missing" } },
                Lookup(new object[,] { { key, "absent" } }, lookup, values, "missing", search: search));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ExactLookupsNormalizeSmallPunctuation(double search)
    {
        const string standard = "+-<>=++";
        const string small = "\uFE62\uFE63\uFE64\uFE65\uFE66⁺₊";
        for (int i = 0; i < standard.Length; i++)
        {
            string key = $"SKU{small[i]}123";
            var lookup = new object[,] { { $"SKU{standard[i]}123" } };

            Assert.Equal("found", Lookup(key, lookup, "found", "missing", search: search));
            Assert.Equal(new object[,] { { "found", "missing" } },
                Lookup(new object[,] { { key, "absent" } }, lookup, "found", "missing", search: search));
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void BatchedTextLookupsAgreeWithScalarWhenWindowsSortKeysDiffer(double search)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            const string key = "A\u0614B";
            var lookup = search == -2
                ? new object[,] { { "Z", "A'\uFC5EB", "A\uFC5EB" } }
                : new object[,] { { "A\uFC5EB", "A'\uFC5EB", "Z" } };

            Assert.Equal("A\uFC5EB", Lookup(key, lookup, lookup, "missing", search: search));
            Assert.Equal(new object[,] { { "A\uFC5EB", "Z" } },
                Lookup(new object[,] { { key, "z" } }, lookup, lookup, "missing", search: search));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    [InlineData(2)]
    [InlineData(-2)]
    public void ExactLookupsIgnoreCollationMarksWhereWindowsDoes(double search)
    {
        CultureInfo original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            var lookup = search == -2 ? new object[,] { { "24", "23\u200D", "22" } } : new object[,] { { "22", "23\u200D", "24" } };

            Assert.Equal("23\u200D", Lookup("23", lookup, lookup, "missing", search: search));
            Assert.Equal("23\u200D", Lookup("23\u200C", lookup, lookup, "missing", search: search));
            Assert.Equal("23\u200D", Lookup("2\u06403", lookup, lookup, "missing", search: search));
            Assert.Equal("23\u200D", Lookup("\u202A23\u202C", lookup, lookup, "missing", search: search));
            Assert.Equal("23\u200D", Lookup("23\u061C", lookup, lookup, "missing", search: search));
            var keys = new object[,] { { "2\u200D3", "23", "23\u200C", "2\u200E3\uFE0F", "2\u06403", "2\u200B3", "\u202A23\u202C", "\u206623\u2069", "2\u061C3" } };
            Assert.Equal(new object[,] { { "23\u200D", "23\u200D", "23\u200D", "23\u200D", "23\u200D", "missing", "23\u200D", "23\u200D", "23\u200D" } },
                Lookup(keys, lookup, lookup, "missing", search: search));
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Theory]
    [InlineData("\u2061", 1d)]
    [InlineData("\u2061", 2d)]
    [InlineData("\u2061", -2d)]
    [InlineData("\U000E0100", 1d)]
    [InlineData("\U000E0100", 2d)]
    [InlineData("\U000E0100", -2d)]
    public void ExactLookupsIgnoreTrailingCollationMarks(string mark, double search)
    {
        string key = "A1" + mark;
        var lookup = search == -2 ? new object[,] { { "A2", "A1!", "A1" } } : new object[,] { { "A1", "A1!", "A2" } };
        var values = search == -2 ? new object[,] { { 20d, 11d, 10d } } : new object[,] { { 10d, 11d, 20d } };

        Assert.Equal(10d, Lookup(key, lookup, values, "missing", search: search));
        Assert.Equal(10d, Lookup("A1", key, 10d, "missing", search: search));
        Assert.Equal(new object[,] { { 10d, 20d } },
            Lookup(new object[,] { { key, "A2" } }, lookup, values, "missing", search: search));
    }

    [Fact]
    public void WildcardScanChecksLongLiteralSuffixAndFindsLaterMatch()
    {
        string suffix = new('a', 8_192);
        var keys = new object[,] { { new string('a', 16_384), "prefix*" + suffix + "b" } };
        var values = new object[,] { { "wrong", "match" } };

        Assert.Equal("match", Lookup("*" + suffix + "b", keys, values, "missing", match: 2));
        Assert.Equal("match", Lookup("*~*" + suffix + "?", keys, values, "missing", match: 2));
        Assert.Equal("match", Lookup("*" + suffix + "b*", keys, values, "missing", match: 2));
    }

    [Fact]
    public void WildcardScanSkipsPositionsWhereRequiredLiteralsCannotAlign()
    {
        const int repeats = 4096;
        string pattern = "*" + string.Concat(Enumerable.Repeat("a?", repeats)) + "b?c*";
        string noFinalLiteral = new('a', repeats * 4);
        var keys = new object[,] { { noFinalLiteral, "b" + noFinalLiteral, noFinalLiteral + "byyc",
            "xx" + new string('a', repeats * 2) + "bycy" } };
        var values = new object[,] { { "missing b", "b out of order", "c misaligned", "match" } };

        Assert.Equal("match", Lookup(pattern, keys, values, "missing", match: 2));
    }

    [Fact]
    public void WildcardScanFindsLateAlignmentAfterRepeatedPrefixMatches()
    {
        const int repeats = 4096;
        string pattern = "*" + string.Concat(Enumerable.Repeat("a?", repeats)) + "b?*";
        string alternating = string.Concat(Enumerable.Repeat("ab", repeats * 2));
        var keys = new object[,] { { alternating, alternating + "bq" } };
        var values = new object[,] { { "different offsets", "match" } };

        Assert.Equal("match", Lookup(pattern, keys, values, "missing", match: 2));
    }

    [Fact]
    public void OmittedLookupValueFindsBlankAndASelectedBlankReturnsZero()
    {
        var keys = new object[,] { { 0d, "", ExcelEmpty.Value } };
        var values = new object[,] { { "zero", "empty text", ExcelEmpty.Value } };

        Assert.Equal(0d, Lookup(ExcelMissing.Value, keys, values));
        Assert.Equal("zero", Lookup(0d, keys, values));
        Assert.Equal("empty text", Lookup("", keys, values));
    }

    [Fact]
    public void CalculatedLookupResultsPreserveBlanksInArrays()
    {
        var column = new object[,] { { ExcelEmpty.Value }, { "hit" } };
        var row = new object[,] { { ExcelEmpty.Value, "hit" } };
        var keys = new object[,] { { 1d, 2d } };
        var expectedRow = (object[,])row.Clone();
        var expectedColumn = (object[,])column.Clone();

        Assert.Equal(expectedRow, Lookup(keys, keys, row));
        Assert.Equal(expectedColumn, Lookup(1d, 1d, column));
        Assert.Equal(expectedRow, LookupFunctions.XLookup(keys, keys, row, "missing", new object[,] { { 0d, 1d } }, 1d));
    }

    [Fact]
    public void FallbackIsUsedOnlyForAMissingMatch()
    {
        var keys = new object[,] { { "found", ExcelError.ExcelErrorRef } };
        var values = new object[,] { { ExcelError.ExcelErrorDiv0, "ignored" } };

        Assert.Equal("missing", Lookup("other", keys, values, "missing"));
        Assert.Equal(ExcelError.ExcelErrorNA, Lookup("other", keys, values));
        Assert.Equal(ExcelError.ExcelErrorDiv0, Lookup("found", keys, values, "missing"));
        Assert.Equal(ExcelError.ExcelErrorRef, Lookup(ExcelError.ExcelErrorRef, keys, values, "missing"));
    }

    [Fact]
    public void CalculatedArraysReturnTheMatchedRowOrColumn()
    {
        var table = new object[,] { { "ID", "Alice" }, { 8d, 14d } };
        var row = Assert.IsType<object[,]>(Lookup(2d, new object[,] { { 1d }, { 2d } }, table));
        var column = Assert.IsType<object[,]>(Lookup(2d, new object[,] { { 1d, 2d } }, table));

        Assert.Equal(new object[,] { { 8d, 14d } }, row);
        Assert.Equal(new object[,] { { "Alice" }, { 14d } }, column);
        Assert.Equal(new object[,] { { "ID", "Alice" } }, Lookup(1d, 1d, new object[,] { { "ID", "Alice" } }));
        Assert.Equal(new object[,] { { "Alice" }, { 14d } }, Lookup(1d, 1d, new object[,] { { "Alice" }, { 14d } }));
    }

    [Fact]
    public void WorksheetResultsPreserveTheMatchedCoordinatesAndSheet()
    {
        IntPtr sheet = new(42);
        var table = new ExcelReference(10, 11, 20, 21, sheet);

        Assert.Equal(new ExcelReference(11, 11, 20, 21, sheet),
            Lookup(2d, new object[,] { { 1d }, { 2d } }, table));
        Assert.Equal(new ExcelReference(10, 11, 21, 21, sheet),
            Lookup(2d, new object[,] { { 1d, 2d } }, table));
        Assert.Equal(new ExcelReference(11, 11, 20, 20, sheet),
            Lookup(2d, new object[,] { { 1d }, { 2d } }, new ExcelReference(10, 11, 20, 20, sheet)));
    }

    [Fact]
    public void WorksheetKeyCellKeepsScalarLookupAndReturnShape()
    {
        var key = new ExcelReference(8, 8, 9, 9, new IntPtr(42));
        var lookup = new object[,] { { 1d }, { 2d } };
        var values = new ExcelReference(10, 11, 20, 21, new IntPtr(43));

        Assert.Equal(new ExcelReference(11, 11, 20, 21, values.SheetId),
            LookupFunctions.XLookup(key, lookup, values, "missing", 0d, 1d, ReadCell, is64BitProcess: true));

        object ReadCell(ExcelReference cell)
        {
            Assert.Equal(key, cell);
            return 2d;
        }
    }

    [Theory]
    [InlineData(1, 511, true)]
    [InlineData(-1, 511, true)]
    [InlineData(1, 1_025, false)]
    [InlineData(-1, 1_025, false)]
    public void ScalarWorksheetScanProbesTheFirstCellAndContinuesInBlocks(double search, int length, bool earlyMatch)
    {
        IntPtr sheet = new(42);
        var keys = new ExcelReference(10, 10 + length - 1, 2, 2, sheet);
        var values = new ExcelReference(10, keys.RowLast, 3, 3, sheet);
        int first = search == 1 ? keys.RowFirst : keys.RowLast;
        int last = search == 1 ? keys.RowLast : keys.RowFirst;
        int target = earlyMatch ? first : last;
        var reads = new List<ExcelReference>();

        Assert.Equal(new ExcelReference(target, target, 3, 3, sheet),
            LookupFunctions.XLookup((double)target, keys, values, "missing", 0d, search, ReadBlock, is64BitProcess: true));
        Assert.Equal(earlyMatch ? 1 : 3, reads.Count);
        Assert.Equal(new ExcelReference(first, first, 2, 2, sheet), reads[0]);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            int rows = block.RowLast - block.RowFirst + 1;
            if (rows == 1)
                return (double)block.RowFirst;

            var cells = new object[rows, 1];
            for (int row = 0; row < rows; row++)
                cells[row, 0] = (double)(block.RowFirst + row);
            return cells;
        }
    }

    [Theory]
    [InlineData(0, 1)]
    [InlineData(0, -1)]
    [InlineData(2, 1)]
    [InlineData(2, -1)]
    public void BatchWithOneDistinctKeyReadsOnlyTheFirstCellOnAnEarlyMatch(double match, double search)
    {
        IntPtr sheet = new(42);
        var lookup = new ExcelReference(10, 522, 2, 2, sheet);
        int first = search == 1 ? lookup.RowFirst : lookup.RowLast;
        var returns = new object[513, 1];
        returns[first - lookup.RowFirst, 0] = "found";
        string key = match == 2 ? "hit*" : "hit";
        var keys = new object[,] { { key, key } };
        var reads = new List<ExcelReference>();

        Assert.Equal(new object[,] { { "found", "found" } },
            LookupFunctions.XLookup(keys, lookup, returns, "missing", match, search, block =>
            {
                reads.Add(block);
                return "hit";
            }, is64BitProcess: true));
        Assert.Equal(new ExcelReference(first, first, 2, 2, sheet), Assert.Single(reads));
    }

    [Theory]
    [InlineData(4_097, false)]
    [InlineData(16_385, true)]
    public void OversizedWorksheetKeyRangesReturnNumWithoutReading(int rows, bool is64BitProcess)
    {
        var keyRange = new ExcelReference(0, rows - 1, 0, 1_023, new IntPtr(42));

        Assert.Equal(ExcelError.ExcelErrorNum,
            LookupFunctions.XLookup(keyRange, 1d, 10d, "missing", 0d, 1d, UnexpectedRead, is64BitProcess));
    }

    [Theory]
    [InlineData(4_096, 2_048, false)]
    [InlineData(16_384, 8_192, true)]
    public void LongWorksheetKeysStopReadingBeforeTheWholeRangeIsCopied(int rows, int expectedCellsRead, bool is64BitProcess)
    {
        var keyRange = new ExcelReference(0, rows - 1, 0, 0, new IntPtr(42));
        string text = new('x', 32_767);
        int cellsRead = 0;

        object result = LookupFunctions.XLookup(keyRange, 1d, 42d, 0d, 0d, 1d, ReadBlock, is64BitProcess);

        Assert.Equal(ExcelError.ExcelErrorNum, result);
        Assert.Equal(expectedCellsRead, cellsRead);

        object ReadBlock(ExcelReference block)
        {
            int length = block.RowLast - block.RowFirst + 1;
            cellsRead += length;
            var keys = new object[length, 1];
            for (int row = 0; row < length; row++)
                keys[row, 0] = text;
            return keys;
        }
    }

    [Theory]
    [InlineData(0, 41, false)]
    [InlineData(0, 42, true)]
    [InlineData(-1, 20, false)]
    [InlineData(-1, 21, true)]
    public void BatchWorkingBudgetIncludesScanIndexes(double match, int numericKeys, bool overBudget)
    {
        const int textKeys = 2_047;
        string text = new('x', 32_767);
        var keys = new ExcelReference(0, textKeys + numericKeys - 1, 0, 0, new IntPtr(42));
        var lookup = new ExcelReference(0, 0, 0, 0, new IntPtr(43));
        int lookupReads = 0;

        object result = LookupFunctions.XLookup(keys, lookup, "found", "missing", match, 1d, ReadBlock,
            is64BitProcess: false);

        if (overBudget)
        {
            Assert.Equal(ExcelError.ExcelErrorNum, result);
            Assert.Equal(0, lookupReads);
        }
        else
        {
            var matches = Assert.IsType<object[,]>(result);
            Assert.Equal(textKeys + numericKeys, matches.Length);
            Assert.Equal("found", matches[0, 0]);
            Assert.Equal("missing", matches[textKeys, 0]);
            Assert.Equal(1, lookupReads);
        }

        object ReadBlock(ExcelReference block)
        {
            if (block.SheetId == lookup.SheetId)
            {
                lookupReads++;
                return text;
            }
            int length = block.RowLast - block.RowFirst + 1;
            var cells = new object[length, 1];
            for (int row = 0; row < length; row++)
            {
                int offset = block.RowFirst + row;
                cells[row, 0] = offset < textKeys ? text : (double)offset;
            }
            return cells;
        }
    }

    [Fact]
    public void WorksheetKeysKeepShapeAndErrorsAcrossBoundedReads()
    {
        var keyRange = new ExcelReference(10, 267, 20, 21, new IntPtr(42));
        var reads = new List<ExcelReference>();

        var result = Assert.IsType<object[,]>(LookupFunctions.XLookup(keyRange, 1d, "hit", "missing",
            0d, 1d, ReadBlock, is64BitProcess: true));

        Assert.Equal(new[] { new ExcelReference(10, 265, 20, 21, keyRange.SheetId),
            new ExcelReference(266, 267, 20, 21, keyRange.SheetId) }, reads);
        Assert.Equal(258, result.GetLength(0));
        Assert.Equal(2, result.GetLength(1));
        for (int row = 0; row < result.GetLength(0); row++)
        {
            Assert.Equal("hit", result[row, 0]);
            Assert.Equal(row == 257 ? ExcelError.ExcelErrorRef : "missing", result[row, 1]);
        }

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            int rows = block.RowLast - block.RowFirst + 1;
            var cells = new object[rows, 2];
            for (int i = 0; i < rows; i++)
            {
                cells[i, 0] = 1d;
                cells[i, 1] = block.RowFirst + i == 267 ? ExcelError.ExcelErrorRef : ExcelEmpty.Value;
            }
            return cells;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void BatchedFallbackUsesTheCellAtEachMissingKey(bool useReference)
    {
        var keys = new object[,] { { 2d, 9d }, { 8d, 1d } };
        var lookup = new object[,] { { 1d, 2d } };
        var values = new object[,] { { 10d, 20d } };
        var fallback = new object[,] { { ExcelError.ExcelErrorNA, "missing 9" }, { ExcelEmpty.Value, ExcelError.ExcelErrorNA } };
        IntPtr sheet = new(42);
        var reads = new List<ExcelReference>();

        object result = LookupFunctions.XLookup(keys, lookup, values,
            useReference ? new ExcelReference(10, 11, 20, 21, sheet) : fallback, 0d, 1d, ReadCell, is64BitProcess: true);

        Assert.Equal(new object[,] { { 20d, "missing 9" }, { 0d, 10d } }, result);
        ExcelReference[] expectedReads = useReference ? [new(10, 10, 21, 21, sheet), new(11, 11, 20, 20, sheet)] : [];
        Assert.Equal(expectedReads, reads);

        object ReadCell(ExcelReference cell)
        {
            reads.Add(cell);
            return fallback[cell.RowFirst - 10, cell.ColumnFirst - 20];
        }
    }

    [Fact]
    public void BatchedFallbackBroadcastsASingleCellOrTheFirstValueOfAnUnshapedArray()
    {
        var keys = new object[,] { { 9d, 8d, 1d } };
        var fallback = new ExcelReference(10, 10, 20, 20, new IntPtr(42));
        int reads = 0;

        Assert.Equal(new object[,] { { "missing", "missing", 10d } },
            LookupFunctions.XLookup(keys, 1d, 10d, fallback, 0d, 1d, ReadCell, is64BitProcess: true));
        Assert.Equal(1, reads);
        Assert.Equal(new object[,] { { "one", "one", 10d } },
            Lookup(keys, 1d, 10d, new object[,] { { "one", "two" } }));

        object ReadCell(ExcelReference cell)
        {
            Assert.Equal(fallback, cell);
            reads++;
            return "missing";
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AdjacentWorksheetFallbacksUseBoundedReadsAndSkipHits(bool vertical)
    {
        const int length = 519;
        int rows = vertical ? length : 2;
        int columns = vertical ? 1 : length;
        var keys = new object[rows, columns];
        var fallback = new ExcelReference(10, 10 + rows - 1, 20, 20 + columns - 1, new IntPtr(42));
        var reads = new List<ExcelReference>();
        var expected = new object[rows, columns];
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++)
            {
                int index = vertical ? row : column;
                keys[row, column] = index switch { 514 => 1d, 515 => ExcelError.ExcelErrorRef, _ => 9d };
                expected[row, column] = index switch
                {
                    514 => "hit",
                    515 => ExcelError.ExcelErrorRef,
                    0 => 0d,
                    _ => (double)(row * 1_000 + column)
                };
            }

        Assert.Equal(expected,
            LookupFunctions.XLookup(keys, 1d, "hit", fallback, 0d, 1d, ReadBlock, is64BitProcess: true));
        ExcelReference[] expectedReads = vertical
            ? [new(10, 521, 20, 20, fallback.SheetId), new(522, 523, 20, 20, fallback.SheetId),
                new(526, 528, 20, 20, fallback.SheetId)]
            : Enumerable.Range(10, 2).SelectMany(row => new ExcelReference[]
            {
                new(row, row, 20, 531, fallback.SheetId), new(row, row, 532, 533, fallback.SheetId),
                new(row, row, 536, 538, fallback.SheetId)
            }).ToArray();
        Assert.Equal(expectedReads, reads);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            var cells = new object[block.RowLast - block.RowFirst + 1, block.ColumnLast - block.ColumnFirst + 1];
            for (int row = 0; row < cells.GetLength(0); row++)
                for (int column = 0; column < cells.GetLength(1); column++)
                {
                    int sourceRow = block.RowFirst - 10 + row;
                    int sourceColumn = block.ColumnFirst - 20 + column;
                    cells[row, column] = (vertical ? sourceRow : sourceColumn) == 0
                        ? ExcelEmpty.Value
                        : (double)(sourceRow * 1_000 + sourceColumn);
                }
            return cells;
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void AlternatingFallbackMissesShareBoundedReads(bool vertical)
    {
        const int length = 515;
        var keys = new object[vertical ? length : 1, vertical ? 1 : length];
        var expected = new object[vertical ? length : 1, vertical ? 1 : length];
        for (int i = 0; i < length; i++)
        {
            bool missing = i % 2 == 0;
            keys[vertical ? i : 0, vertical ? 0 : i] = missing ? 0d : 1d;
            expected[vertical ? i : 0, vertical ? 0 : i] = missing ? (object)(double)i : "hit";
        }
        var reads = new List<ExcelReference>();
        Assert.Equal(expected, LookupFunctions.XLookup(keys, 1d, "hit", Range(0, length - 1), 0d, 1d, ReadBlock, true));
        Assert.Equal(new[] { Range(0, 510), Range(512, 514) }, reads);

        ExcelReference Range(int first, int last) => vertical
            ? new(first, last, 20, 20, new IntPtr(42))
            : new(10, 10, first, last, new IntPtr(42));

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            int first = vertical ? block.RowFirst : block.ColumnFirst;
            int last = vertical ? block.RowLast : block.ColumnLast;
            if (first == last)
                return (double)first;
            int count = last - first + 1;
            var cells = new object[vertical ? count : 1, vertical ? 1 : count];
            for (int i = 0; i < count; i++)
                cells[vertical ? i : 0, vertical ? 0 : i] = (double)(first + i);
            return cells;
        }
    }

    [Fact]
    public void TallWorksheetFallbacksBatchAcrossRows()
    {
        const int rows = 1_025;
        var keys = new object[rows, 2];
        var expected = new object[rows, 2];
        var fallback = new ExcelReference(10, 10 + rows - 1, 20, 21, new IntPtr(42));
        int reads = 0;
        for (int row = 0; row < rows; row++)
        {
            keys[row, 0] = row % 2 == 0 ? 9d : 1d;
            keys[row, 1] = row == 512 ? ExcelError.ExcelErrorRef : 9d;
            expected[row, 0] = row % 2 == 0 ? (double)(2 * row) : "hit";
            expected[row, 1] = row == 512 ? ExcelError.ExcelErrorRef : (double)(2 * row + 1);
        }

        Assert.Equal(expected, LookupFunctions.XLookup(keys, 1d, "hit", fallback, 0d, 1d, ReadBlock, true));
        Assert.Equal(5, reads);

        object ReadBlock(ExcelReference block)
        {
            reads++;
            Assert.Equal(fallback.SheetId, block.SheetId);
            Assert.Equal(block.ColumnFirst, block.ColumnLast);
            Assert.InRange(block.RowLast - block.RowFirst + 1, 1, 512);
            int length = block.RowLast - block.RowFirst + 1;
            int first = 2 * (block.RowFirst - fallback.RowFirst) + block.ColumnFirst - fallback.ColumnFirst;
            if (length == 1)
                return (double)first;
            var cells = new object[length, 1];
            for (int row = 0; row < length; row++)
                cells[row, 0] = (double)(first + 2 * row);
            return cells;
        }
    }

    [Theory]
    [InlineData(true, 2)]
    [InlineData(false, -2)]
    public void BinaryWorksheetBatchesReuseBoundedProbeReads(bool vertical, double search)
    {
        const int length = 4_099;
        var lookup = vertical ? new ExcelReference(10, 10 + length - 1, 20, 20, new IntPtr(42))
            : new ExcelReference(10, 10, 20, 20 + length - 1, new IntPtr(42));
        var returns = new object[vertical ? length : 1, vertical ? 1 : length];
        var keys = new object[length, 1];
        int reads = 0;
        for (int i = 0; i < length; i++)
        {
            keys[i, 0] = i + 0.5;
            returns[vertical ? i : 0, vertical ? 0 : i] = (double)(vertical ? i : length - i - 1);
        }

        var result = Assert.IsType<object[,]>(LookupFunctions.XLookup(keys, lookup, returns, "missing", -1d, search,
            ReadBlock, is64BitProcess: true));

        Assert.Equal((length, 1), (result.GetLength(0), result.GetLength(1)));
        for (int i = 0; i < length; i++)
            Assert.Equal((double)i, result[i, 0]);
        Assert.InRange(reads, 1, 200);

        object ReadBlock(ExcelReference block)
        {
            reads++;
            Assert.Equal(lookup.SheetId, block.SheetId);
            int count = vertical ? block.RowLast - block.RowFirst + 1 : block.ColumnLast - block.ColumnFirst + 1;
            Assert.InRange(count, 1, 64);
            var values = new object[vertical ? count : 1, vertical ? 1 : count];
            int start = vertical ? block.RowFirst - lookup.RowFirst : block.ColumnFirst - lookup.ColumnFirst;
            for (int i = 0; i < count; i++)
                values[vertical ? i : 0, vertical ? 0 : i] = returns[vertical ? start + i : 0, vertical ? 0 : start + i];
            return values;
        }
    }

    [Fact]
    public void BinaryWorksheetBatchesLimitCopiesForScatteredKeys()
    {
        const int length = 1_048_576;
        var sheet = new IntPtr(42);
        var lookup = new ExcelReference(0, length - 1, 0, 0, sheet);
        var returns = new ExcelReference(0, length - 1, 1, 1, sheet);
        var keys = new object[1_024, 1];
        var expected = new object[keys.GetLength(0), 1];
        var random = new Random(3411);
        for (int i = 0; i < keys.Length; i++)
        {
            int row = random.Next(length);
            keys[i, 0] = (double)row;
            expected[i, 0] = -(row + 1d);
        }

        int cellsRead = 0;
        int reads = 0;
        Assert.Equal(expected, LookupFunctions.XLookup(keys, lookup, returns, "missing", 0d, 2d, Read, is64BitProcess: true));
        Assert.InRange(cellsRead, 1, 150_000);
        Assert.InRange(reads, 1, 20_000);

        object Read(ExcelReference block)
        {
            reads++;
            int count = block.RowLast - block.RowFirst + 1;
            cellsRead += count;
            bool readingKeys = block.ColumnFirst == lookup.ColumnFirst;
            if (count == 1)
                return readingKeys ? (double)block.RowFirst : -(block.RowFirst + 1d);

            var values = new object[count, 1];
            for (int i = 0; i < count; i++)
            {
                int row = block.RowFirst + i;
                values[i, 0] = readingKeys ? (double)row : -(row + 1d);
            }
            return values;
        }
    }

    [Fact]
    public void BinaryWorksheetProbesRespectTheWorkingBudgetAfterEviction()
    {
        const int tableLength = 4096;
        const int distinctKeys = 16;
        var table = new ExcelReference(0, tableLength - 1, 0, 0, new IntPtr(42));
        string suffix = new('x', 32_762);
        var keyText = Enumerable.Range(0, distinctKeys).Select(i => Text(ResultAt(i))).ToArray();
        var returns = new object[tableLength, 1];
        for (int row = 0; row < tableLength; row++)
            returns[row, 0] = (double)row;

        Check(1970, expectBlockReads: true);
        Check(2000, expectBlockReads: false);

        void Check(int keyCount, bool expectBlockReads)
        {
            var keys = new ExcelReference(0, keyCount - 1, 0, 0, new IntPtr(43));
            int blockReads = 0;
            var result = Assert.IsType<object[,]>(
                LookupFunctions.XLookup(keys, table, returns, "missing", 0d, 2d, Read, is64BitProcess: false));
            Assert.Equal(keyCount, result.Length);
            for (int row = 0; row < keyCount; row++)
                Assert.Equal((double)ResultAt(row % distinctKeys), result[row, 0]);

            if (expectBlockReads)
            {
                // Only one full block fits. Further promotions reuse the evicted block's budget.
                Assert.True(blockReads > 1, $"Expected multiple block reads after eviction, got {blockReads}.");
            }
            else
                Assert.Equal(0, blockReads);

            object Read(ExcelReference block)
            {
                int count = block.RowLast - block.RowFirst + 1;
                if (block.SheetId == table.SheetId)
                {
                    if (count == 1)
                        return Text(block.RowFirst);
                    blockReads++;
                }

                var values = new object[count, 1];
                for (int i = 0; i < count; i++)
                {
                    int row = block.RowFirst + i;
                    values[i, 0] = block.SheetId == keys.SheetId ? keyText[row % distinctKeys] : Text(row);
                }
                return values;
            }
        }

        string Text(int row) => row.ToString("D5", CultureInfo.InvariantCulture) + suffix;
        static int ResultAt(int index) => 56 + 251 * index;
    }

    [Fact]
    public void BinaryBatchesWithOneDistinctKeyKeepSingleCellProbes()
    {
        const int length = 256;
        var lookup = new ExcelReference(0, length - 1, 0, 0, new IntPtr(42));
        var keys = new object[length, 1];
        var values = new object[length, 1];
        for (int i = 0; i < length; i++)
            keys[i, 0] = 1d;
        values[1, 0] = "found";

        int reads = 0;
        object result = LookupFunctions.XLookup(keys, lookup, values, "missing", 0d, 2d,
            block =>
            {
                Assert.Equal(block.RowFirst, block.RowLast);
                reads++;
                return (double)block.RowFirst;
            }, is64BitProcess: true);
        AssertTextColumn(result, length, "found");
        Assert.InRange(reads, 1, 9);
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void DistinctWildcardPatternsShareEachWorksheetBlock(double search, string match)
    {
        IntPtr sheet = new(42);
        var lookup = new ExcelReference(10, 10, 20, 1_044, sheet);
        var values = new object[1, 1_025];
        values[0, 0] = "first";
        values[0, 512] = "number";
        values[0, 513] = "blank";
        values[0, 1_024] = "last";
        var keys = new object[,] { { "a*", "z*", "a*", 7d, ExcelMissing.Value, ExcelError.ExcelErrorRef } };
        var reads = new List<ExcelReference>();

        object result = LookupFunctions.XLookup(keys, lookup, values, "missing", 2d, search,
            ReadBlock, is64BitProcess: true);

        Assert.Equal(new object[,] { { match, "missing", match, "number", "blank", ExcelError.ExcelErrorRef } }, result);
        ExcelReference[] expected = search == 1
            ? [new(10, 10, 20, 531, sheet), new(10, 10, 532, 1_043, sheet), new(10, 10, 1_044, 1_044, sheet)]
            : [new(10, 10, 533, 1_044, sheet), new(10, 10, 21, 532, sheet), new(10, 10, 20, 20, sheet)];
        Assert.Equal(expected, reads);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            if (block.ColumnLast == block.ColumnFirst)
                return "apple";
            var cells = new object[1, block.ColumnLast - block.ColumnFirst + 1];
            for (int column = block.ColumnFirst; column <= block.ColumnLast; column++)
                cells[0, column - block.ColumnFirst] = column switch
                {
                    20 or 1_044 => "apple",
                    532 => 7d,
                    533 => ExcelEmpty.Value,
                    _ => "other"
                };
            return cells;
        }
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void WildcardPrefixBucketsPreservePatternSyntaxAndSearchDirection(double search, string duplicate)
    {
        var keys = new object[,] { { "prefix~?A*", "?oo", "*foo", "ABCDEFGHIJK\U00010428*", "Ａ*", "PREFIX~?B*" } };
        var lookup = new object[,] { { "PREFIX?aa", "prefixZaa", "foo", "ABCDEFGHIJK\U00010428-rest", "prefix?ab", "ACME", "prefix?b" } };
        var values = new object[,] { { "first", "wrong", "foo", "unicode", "last", "wrong width", "B" } };

        Assert.Equal(new object[,] { { duplicate, "foo", "foo", "unicode", "missing", "B" } },
            Lookup(keys, lookup, values, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void WildcardBatchesUseFullAnchorsOfDifferentLengths(double search, string duplicate)
    {
        var keys = new object[,] { { "sku-variant-*", "sku-variant-01*", "sku-variant-01b*", "*east-warehouse-09", "*west-warehouse-09", "part-*", "*!" } };
        var lookup = new object[,] { { "", "sku-variant-01a", "east-warehouse-09", "sku-variant-01b", "sku-variant-01c", "part-a", "order!", "" } };
        var values = new object[,] { { "empty", "first", "east", "only B", "last", "part", "punctuation", "empty" } };

        Assert.Equal(new object[,] { { duplicate, duplicate, "only B", "east", "missing", "part", "punctuation" } },
            Lookup(keys, lookup, values, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void WildcardSuffixBucketsCheckTheFullPatternInSearchOrder(double search, string duplicate)
    {
        var keys = new object[,] { { "*sku", "?x*~*", "*\U00010428ABCDEFGHIJK", "*missing", "*?sku" } };
        var lookup = new object[,] { { "firstSku", "wrong*", "ax-match*", "end\U00010428ABCDEFGHIJK", "lastSKU", "sku " } };
        var values = new object[,] { { "first", "wrong", "escaped", "unicode", "last", "wrong end" } };

        Assert.Equal(new object[,] { { duplicate, "escaped", "unicode", "missing", duplicate } },
            Lookup(keys, lookup, values, "missing", match: 2, search: search));
        Assert.Equal("unicode", Lookup(keys[0, 2], lookup, values, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData(1, "first", "first")]
    [InlineData(-1, "last", "second code")]
    public void RepeatedWildcardCandidatesKeepTheCorrectMatchOrder(double search, string duplicate, string anchored)
    {
        var keys = new object[,] { { "*code-1*", "*code-2*", "*missing*", "ready*" } };
        var lookup = new object[,] { { "ready code-1", "closed", "ready code-1", "READY CODE-2", "closed" } };
        var values = new object[,] { { "first", "ignored", "last", "second code", "ignored" } };

        Assert.Equal(new object[,] { { duplicate, "second code", "missing", anchored } },
            Lookup(keys, lookup, values, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void WildcardSubstringBucketsCheckOverlappingFragmentsAndFullPatterns(double search, string duplicate)
    {
        var keys = new object[,] { { "*ab*", "*ba*", "*~?ab*", "*ABCDEFGH-123456789012*", "*ZYXWVUTS-123456789012*", "*?*" } };
        var lookup = new object[,] { { "aBa", "s?AB", "xZYXWVUTS-123456789012y", "xABCDEFGH-123456789012y", "aba" } };
        var values = new object[,] { { "first", "escaped", "reverse letters", "ordered letters", "last" } };

        Assert.Equal(new object[,] { { duplicate, duplicate, "escaped", "ordered letters", "reverse letters", duplicate } },
            Lookup(keys, lookup, values, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData(1, "first", true)]
    [InlineData(-1, "last", true)]
    [InlineData(1, "first", false)]
    [InlineData(-1, "last", false)]
    public void WildcardBatchesFindMatchesWithSharedLiterals(double search, string duplicate, bool anchored)
    {
        var keys = anchored
            ? new object[,] { { "PRODUCTGROUP001*ID0001*", "PRODUCTGROUP001*ID0002*", "PRODUCTGROUP001*ID0003*" } }
            : new object[,] { { "*ID0001?PRODUCTGROUP001*", "*ID0002?PRODUCTGROUP001*", "*ID0003?PRODUCTGROUP001*" } };
        var lookup = anchored
            ? new object[,] { { "PRODUCTGROUP001-ID0001", "PRODUCTGROUP001-ID0002",
                "X-PRODUCTGROUP001-ID0003", "PRODUCTGROUP001-ID0001-2" } }
            : new object[,] { { "ID0001-PRODUCTGROUP001", "ID0002-PRODUCTGROUP001",
                "ID0003-PRODUCTGROUP002", "ID0001-PRODUCTGROUP001-2" } };
        var values = new object[,] { { "first", "second", "wrong group", "last" } };

        Assert.Equal(new object[,] { { duplicate, "second", "missing" } },
            Lookup(keys, lookup, values, "missing", match: 2, search: search));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WildcardBatchesStopProbingRetiredAnchors(bool suffix)
    {
        const int patternCount = 300;
        const int rows = 13_000;
        var keys = new object[1, patternCount + 1];
        for (int i = 0; i < patternCount; i++)
        {
            string literal = new('A', i + 1);
            keys[0, i] = suffix ? "*" + literal : literal + "*";
        }
        keys[0, patternCount] = "Z*";

        var lookup = new object[rows, 1];
        string anchor = new('A', patternCount);
        for (int i = 0; i < rows - 1; i++)
            lookup[i, 0] = suffix ? i + anchor : anchor + i;
        lookup[rows - 1, 0] = "Zebra";

        var result = Assert.IsType<object[,]>(Lookup(keys, lookup, lookup, match: 2));
        for (int i = 0; i < patternCount; i++)
            Assert.Equal(lookup[0, 0], result[0, i]);
        Assert.Equal("Zebra", result[0, patternCount]);
    }

    [Fact]
    public void WildcardBatchesRefreshSubstringIndexAfterRepeatedRetiredHits()
    {
        var keys = new object[,] { { "*A*", "*B*", "*C*" } };
        var lookup = new object[425, 1];
        string repeated = new('A', 25_000);
        for (int i = 0; i < lookup.Length - 1; i++)
            lookup[i, 0] = repeated + i;
        lookup[lookup.Length - 1, 0] = "BC";

        Assert.Equal(new object[,] { { lookup[0, 0], "BC", "BC" } }, Lookup(keys, lookup, lookup, match: 2));
    }

    [Theory]
    [InlineData(false, 1, false)]
    [InlineData(false, 300, true)]
    [InlineData(true, 1, false)]
    [InlineData(true, 300, true)]
    public void WildcardLookupsLimitWorkInsideTheMatcher(bool batched, int rows, bool overBudget)
    {
        string pattern = "*" + string.Concat(Enumerable.Repeat("A?", 8_000)) + "B*";
        string text = new string('A', 15_998) + string.Concat(Enumerable.Repeat("BA", 7_990));
        object keys = batched ? new object[,] { { pattern, "Z*" } } : pattern;
        var lookup = new object[rows, 1];
        for (int i = 0; i < rows; i++)
            lookup[i, 0] = text + i;

        object result = Lookup(keys, lookup, lookup, "missing", match: 2);

        if (overBudget)
            Assert.Equal(ExcelError.ExcelErrorNum, result);
        else if (batched)
            Assert.Equal(new object[,] { { "missing", "missing" } }, result);
        else
            Assert.Equal("missing", result);
    }

    [Fact]
    public void WildcardBatchesSkipPatternsThatCannotFitTheText()
    {
        var keys = new object[4_098, 1];
        keys[0, 0] = "0000";
        keys[1, 0] = "?????";
        for (int i = 0; i < 4_096; i++)
            keys[i + 2, 0] = string.Concat(Enumerable.Range(0, 12).Select(bit => (i & (1 << bit)) == 0 ? "?*" : "*?"));
        var lookup = new object[2_442, 1];
        for (int i = 0; i < lookup.Length - 1; i++)
            lookup[i, 0] = i.ToString("D4", CultureInfo.InvariantCulture);
        lookup[lookup.Length - 1, 0] = "later";

        var result = Assert.IsType<object[,]>(Lookup(keys, lookup, lookup, "miss", match: 2));

        Assert.Equal(keys.Length, result.Length);
        Assert.Equal("0000", result[0, 0]);
        Assert.Equal("later", result[1, 0]);
        Assert.All(result.Cast<object>().Skip(2), value => Assert.Equal("miss", value));
    }

    [Fact]
    public void WildcardBatchBudgetIncludesPrefixProbesThatMiss()
    {
        const int patterns = 600;
        var keys = new object[patterns, 1];
        for (int index = 0; index < patterns; index++)
            keys[index, 0] = "A" + new string('B', index + 1) + "*";

        var lookup = new object[4_000, 1];
        string prefix = "A" + new string('C', patterns);
        for (int index = 0; index < lookup.Length; index++)
            lookup[index, 0] = prefix + index.ToString(CultureInfo.InvariantCulture);

        Assert.Equal(ExcelError.ExcelErrorNum, Lookup(keys, lookup, lookup, "missing", match: 2));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(-1)]
    public void WildcardLookupsReachAMatchAfterRepeatedLongText(double search)
    {
        const int distinctRows = 300;
        const int repeatedRows = 33_000;
        string text = new('x', 20_000);
        var lookup = new object[distinctRows + repeatedRows + 1, 1];
        for (int row = 0; row < lookup.Length - 1; row++)
            lookup[search == 1 ? row : lookup.Length - 1 - row, 0] =
                row < distinctRows ? row.ToString(CultureInfo.InvariantCulture) : text;
        lookup[search == 1 ? lookup.Length - 1 : 0, 0] = "needle";

        Assert.Equal("needle", Lookup("needle*", lookup, lookup, "miss", match: 2, search: search));
        Assert.Equal(new object[,] { { "needle", "miss" } },
            Lookup(new object[,] { { "needle*", "other*" } }, lookup, lookup, "miss", match: 2, search: search));
    }

    [Theory]
    [InlineData(19_531, false)]
    [InlineData(19_532, true)]
    public void WildcardBatchBudgetCountsUncachedTextEvenWhenNoPrefixCanMatch(int rows, bool overBudget)
    {
        string prefix = new('A', 32_763);
        string[] texts = Enumerable.Range(0, 257)
            .Select(i => prefix + i.ToString("D4", CultureInfo.InvariantCulture)).ToArray();
        var lookup = new object[rows, 1];
        for (int row = 0; row < rows; row++)
            lookup[row, 0] = texts[row % texts.Length];

        object result = Lookup(new object[,] { { "B*", "C*" } }, lookup, lookup, "missing", match: 2);

        if (overBudget)
            Assert.Equal(ExcelError.ExcelErrorNum, result);
        else
            Assert.Equal(new object[,] { { "missing", "missing" } }, result);
    }

    [Fact]
    public void CompiledWildcardPatternsRespectTheWorkingBudget()
    {
        const int count = 384;
        var keys = new object[count, 1];
        string common = string.Concat(Enumerable.Repeat("*?", 16_000));
        for (int i = 0; i < count; i++)
            keys[i, 0] = common + i + "*";
        var lookup = new ExcelReference(0, 0, 0, 0, new IntPtr(42));

        Assert.Equal(ExcelError.ExcelErrorNum,
            LookupFunctions.XLookup(keys, lookup, "hit", "missing", 2d, 1d, UnexpectedRead, is64BitProcess: false));
    }

    [Fact]
    public void RepeatedWorksheetMatchesReadEachReturnCellOncePerCall()
    {
        IntPtr sheet = new(42);
        var keys = new object[,] { { "b", "B", "a", "b", "a", "missing" } };
        var lookup = new object[,] { { "a" }, { "b" } };
        var values = new ExcelReference(10, 11, 20, 21, sheet);
        var reads = new List<ExcelReference>();
        double secondValue = 20d;

        Assert.Equal(new object[,] { { 20d, 20d, ExcelEmpty.Value, 20d, ExcelEmpty.Value, "missing" } }, Evaluate());
        Assert.Equal(new ExcelReference(10, 11, 20, 20, sheet), Assert.Single(reads));

        reads.Clear();
        secondValue = 25d;
        Assert.Equal(new object[,] { { 25d, 25d, ExcelEmpty.Value, 25d, ExcelEmpty.Value, "missing" } }, Evaluate());
        Assert.Single(reads);

        object Evaluate() => LookupFunctions.XLookup(keys, lookup, values, "missing", 0d, 1d,
            block =>
            {
                reads.Add(block);
                return new object[,] { { ExcelEmpty.Value }, { secondValue } };
            }, is64BitProcess: true);
    }

    [Fact]
    public void RepeatedTextKeysReuseNormalizationAcrossTheBatch()
    {
        const int count = 10_000;
        string key = string.Concat(Enumerable.Repeat("ｶﾞ", 8));
        var keys = new object[count, 1];
        for (int row = 0; row < count; row++)
            keys[row, 0] = key;
        var lookup = new object[,] { { "ガガガガガガガガ" } };
        var values = new ExcelReference(0, 0, 0, 0, new IntPtr(42));

        _ = Evaluate();
        long before = GC.GetAllocatedBytesForCurrentThread();
        object result = Evaluate();
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        AssertTextColumn(result, count, "hit");
        Assert.InRange(allocated, 0, 4L * 1024 * 1024);

        object Evaluate() => LookupFunctions.XLookup(keys, lookup, values, "missing", 0d, 1d, _ => "hit",
            is64BitProcess: true);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AdjacentWorksheetReturnsUseBoundedReadsAcrossSmallGaps(bool vertical)
    {
        IntPtr sheet = new(42);
        var lookup = new object[vertical ? 518 : 1, vertical ? 1 : 518];
        for (int index = 0; index < 518; index++)
            lookup[vertical ? index : 0, vertical ? 0 : index] = (double)index;
        var returns = Range(0, 517);
        var keys = new object[1, 517];
        var expected = new object[1, 517];
        keys[0, 0] = keys[0, 516] = 517d;
        expected[0, 0] = expected[0, 516] = -518d;
        for (int index = 1; index <= 515; index++)
        {
            double key = 515 - index;
            keys[0, index] = key;
            expected[0, index] = -(key + 1d);
        }
        var reads = new List<ExcelReference>();

        object result = LookupFunctions.XLookup(keys, lookup, returns, "missing", 0d, 1d, ReadBlock, is64BitProcess: true);

        Assert.Equal(expected, result);
        Assert.Equal(new[] { Range(0, 511), Range(512, 517) }, reads);

        ExcelReference Range(int first, int last) => vertical
            ? new(10 + first, 10 + last, 20, 20, sheet)
            : new(10, 10, 20 + first, 20 + last, sheet);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            int first = vertical ? block.RowFirst - 10 : block.ColumnFirst - 20;
            int last = vertical ? block.RowLast - 10 : block.ColumnLast - 20;
            if (first == last)
                return -(first + 1d);
            int length = last - first + 1;
            var cells = new object[vertical ? length : 1, vertical ? 1 : length];
            for (int index = 0; index < length; index++)
                cells[vertical ? index : 0, vertical ? 0 : index] = -(first + index + 1d);
            return cells;
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorksheetMatchesSeparatedByTwoCellsShareBoundedReads(bool vertical)
    {
        const int length = 520;
        const int stride = 3;
        var lookup = new object[vertical ? length : 1, vertical ? 1 : length];
        var keys = new object[1, (length - 1) / stride + 1];
        var expected = new object[1, keys.Length];
        for (int index = 0; index < length; index++)
        {
            lookup[vertical ? index : 0, vertical ? 0 : index] = (double)index;
            if (index % stride == 0)
            {
                keys[0, index / stride] = (double)index;
                expected[0, index / stride] = -(index + 1d);
            }
        }

        var reads = new List<ExcelReference>();
        Assert.Equal(expected, LookupFunctions.XLookup(keys, lookup, Range(0, length - 1), "missing", 0d, 1d, ReadBlock, true));
        Assert.Equal(new[] { Range(0, 510), Range(513, 519) }, reads);

        ExcelReference Range(int first, int last) => vertical
            ? new(first, last, 0, 0, new IntPtr(42))
            : new(0, 0, first, last, new IntPtr(42));

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            int first = vertical ? block.RowFirst : block.ColumnFirst;
            int last = vertical ? block.RowLast : block.ColumnLast;
            if (first == last)
                return -(first + 1d);

            int count = last - first + 1;
            var values = new object[vertical ? count : 1, vertical ? 1 : count];
            for (int index = 0; index < count; index++)
                values[vertical ? index : 0, vertical ? 0 : index] = (first + index) % stride == 0 ? -(first + index + 1d) : "skipped";
            return values;
        }
    }

    [Fact]
    public void RepeatedWorksheetTextConsumesBudgetBeforeReadingMoreBlocks()
    {
        var keys = new object[8_193, 1];
        for (int row = 0; row < 8_192; row++)
            keys[row, 0] = 1d;
        keys[8_192, 0] = 5d;
        var lookup = new object[,] { { 1d }, { 2d }, { 3d }, { 4d }, { 5d } };
        IntPtr sheet = new(42);
        var returns = new ExcelReference(10, 14, 20, 20, sheet);
        var reads = new List<ExcelReference>();

        object result = LookupFunctions.XLookup(keys, lookup, returns, "missing", 0d, 1d,
            block =>
            {
                reads.Add(block);
                return new string('x', 32_767);
            }, is64BitProcess: false);

        Assert.Equal(ExcelError.ExcelErrorNum, result);
        Assert.Equal(new ExcelReference(10, 10, 20, 20, sheet), Assert.Single(reads));
    }

    [Theory]
    [InlineData(4_094, true, false)]
    [InlineData(4_095, true, false)]
    [InlineData(4_094, false, true)]
    [InlineData(4_095, false, true)]
    [InlineData(4_094, true, true)]
    [InlineData(4_095, true, true)]
    public void BatchedWorksheetResultsUseTheLower32BitLimit(int rows, bool worksheetMatches, bool worksheetFallbacks)
    {
        int matchedRows = worksheetMatches ? (worksheetFallbacks ? rows / 2 : rows) : 0;
        var keys = new object[rows, 1];
        var matchedKeys = new object[matchedRows, 1];
        for (int i = 0; i < rows; i++)
            keys[i, 0] = (double)i;
        Array.Copy(keys, matchedKeys, matchedRows);
        var sheet = new IntPtr(42);
        string text = new('x', 32_767);
        object lookup = worksheetMatches ? matchedKeys : -1d;
        object returns = worksheetMatches ? new ExcelReference(0, matchedRows - 1, 1, 1, sheet) : "hit";
        object fallback = worksheetFallbacks ? new ExcelReference(0, rows - 1, 2, 2, sheet) : "miss";

        object result = LookupFunctions.XLookup(keys, lookup, returns, fallback, 0d, 1d, ReadBlock, is64BitProcess: false);

        if (rows == 4_095)
            Assert.Equal(ExcelError.ExcelErrorNum, result);
        else
            AssertTextColumn(result, rows, text);

        object ReadBlock(ExcelReference block)
        {
            int length = block.RowLast - block.RowFirst + 1;
            if (length == 1)
                return text;
            var values = new object[length, 1];
            for (int i = 0; i < length; i++)
                values[i, 0] = text;
            return values;
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void UnusedWorksheetRangeDoesNotLowerThe32BitLimit(bool fallbackUnused, bool useModeArray)
    {
        var keys = new object[4_095, 1];
        for (int i = 0; i < keys.Length; i++)
            keys[i, 0] = 1d;
        string text = new('x', 32_767);
        var sheet = new IntPtr(42);
        object result = LookupFunctions.XLookup(keys, fallbackUnused ? 1d : 2d,
            fallbackUnused ? text : new ExcelReference(0, 0, 0, 0, sheet),
            fallbackUnused ? new ExcelReference(0, keys.Length - 1, 0, 0, sheet) : text,
            0d, useModeArray ? new object[,] { { 1d } } : 1d, UnexpectedRead, is64BitProcess: false);

        AssertTextColumn(result, keys.Length, text);
    }

    [Theory]
    [InlineData(4_094, false)]
    [InlineData(4_095, true)]
    public void ModeArraysApplyTheLower32BitLimitAcrossGroupsWhenWorksheetResultsAreCopied(int rows, bool tooLarge)
    {
        var keys = new object[rows, 1];
        var searches = new object[rows, 1];
        for (int row = 0; row < rows; row++)
        {
            keys[row, 0] = 1d;
            searches[row, 0] = row == rows - 1 ? -1d : 1d;
        }
        var returns = new ExcelReference(0, 0, 0, 0, new IntPtr(42));
        string text = new('x', 32_767);

        object result = LookupFunctions.XLookup(keys, 1d, returns, "missing", 0d, searches, _ => text,
            is64BitProcess: false);

        if (tooLarge)
            Assert.Equal(ExcelError.ExcelErrorNum, result);
        else
            AssertTextColumn(result, rows, text);
    }

    [Theory]
    [InlineData(8_192, false)]
    [InlineData(32_768, true)]
    public void SingletonModeGroupBudgetsOnlyTheLeadingCell(int rows, bool is64BitProcess)
    {
        var lookup = new object[,] { { 1d, 2d } };
        var returns = new object[rows, 2];
        string text = new('x', 32_767);
        for (int row = 1; row < rows; row++)
            returns[row, 0] = text;
        returns[0, 0] = 42d;

        Assert.Equal(ExcelError.ExcelErrorNum,
            LookupFunctions.XLookup(1d, lookup, returns, "missing", 0d, 1d, UnexpectedRead, is64BitProcess));
        Assert.Equal(new object[,] { { 42d } },
            LookupFunctions.XLookup(1d, lookup, returns, "missing", new object[,] { { 0d } }, 1d,
                UnexpectedRead, is64BitProcess));
    }

    [Fact]
    public void ModeArraysStopReadingWorksheetResultsAtTheSharedTextLimit()
    {
        const int groupSize = 6_500;
        var keys = new object[2 * groupSize, 1];
        var modes = new object[keys.Length, 1];
        for (int row = 0; row < keys.Length; row++)
        {
            keys[row, 0] = (double)row;
            modes[row, 0] = row < groupSize ? 0d : 1d;
        }
        var returns = new ExcelReference(0, keys.Length - 1, 0, 0, new IntPtr(42));
        string text = new('x', 16_384);
        int cellsRead = 0;

        object result = LookupFunctions.XLookup(keys, keys, returns, "missing", modes, 1d, ReadBlock,
            is64BitProcess: false);

        Assert.Equal(ExcelError.ExcelErrorNum, result);
        // The 8,179th result exceeds the allowance. Its block can contain up to 511 further cells.
        Assert.InRange(cellsRead, 8_179, 8_179 + 511);

        object ReadBlock(ExcelReference block)
        {
            int rows = block.RowLast - block.RowFirst + 1;
            cellsRead += rows;
            if (rows == 1)
                return text;
            var values = new object[rows, 1];
            for (int row = 0; row < rows; row++)
                values[row, 0] = text;
            return values;
        }
    }
    [Theory]
    [InlineData(8_192, false)]
    [InlineData(32_768, true)]
    public void BatchedTextResultsRespectNativeBufferLimits(int rows, bool is64BitProcess)
    {
        var keys = new object[rows, 1];
        string text = new('x', 32_767);
        for (int row = 0; row < rows; row++)
            keys[row, 0] = 1d;

        Assert.Equal(ExcelError.ExcelErrorNum, Evaluate());

        for (int row = 0; row < 32; row++)
            keys[row, 0] = 2d;

        var result = Assert.IsType<object[,]>(Evaluate());
        Assert.Equal(rows, result.Length);
        Assert.Equal("missing", result[31, 0]);
        Assert.Equal(text, result[32, 0]);

        object Evaluate() => LookupFunctions.XLookup(keys, 1d, text, "missing", 0d, 1d,
            UnexpectedRead, is64BitProcess);
    }

    [Theory]
    [InlineData(8_192, false)]
    [InlineData(32_768, true)]
    public void CalculatedMatchesAndFallbacksRespectNativeBufferLimits(int rows, bool is64BitProcess)
    {
        var values = new object[rows, 1];
        string text = new('x', 32_767);
        for (int row = 0; row < rows; row++)
            values[row, 0] = text;

        Assert.Equal(ExcelError.ExcelErrorNum, Evaluate(useFallback: false));
        Assert.Equal(ExcelError.ExcelErrorNum, Evaluate(useFallback: true));

        for (int row = 0; row < 32; row++)
            values[row, 0] = "short";
        var expected = (object[,])values.Clone();

        Assert.Equal(expected, Evaluate(useFallback: false));
        Assert.Equal(expected, Evaluate(useFallback: true));

        object Evaluate(bool useFallback) => LookupFunctions.XLookup(1d, useFallback ? 2d : 1d,
            useFallback ? "unexpected hit" : values, useFallback ? values : "unexpected miss",
            0d, 1d, UnexpectedRead, is64BitProcess);
    }

    private static void AssertTextColumn(object result, int rows, string text)
    {
        var cells = Assert.IsType<object[,]>(result);
        Assert.Equal((rows, 1), (cells.GetLength(0), cells.GetLength(1)));
        Assert.All(cells.Cast<object>(), cell => Assert.Equal(text, cell));
    }

    private static object UnexpectedRead(ExcelReference _) =>
        throw new InvalidOperationException("Calculated arrays do not require worksheet reads.");

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void MultipleKeysTakeLeadingCellsAndPreserveKeyErrors(double search)
    {
        const ExcelError keyError = ExcelError.ExcelErrorRef;
        var keys = new object[,] { { 2d, 9d }, { 1d, keyError } };
        var table = new object[,] { { "Alice", "Sales" }, { "Bob", "Engineering" } };

        Assert.Equal(new object[,] { { "Bob", "missing" }, { "Alice", keyError } },
            Lookup(keys, new object[,] { { 1d }, { 2d } }, table, "missing", search: search));
        Assert.Equal(new object[,] { { "Sales", "missing" }, { "Alice", keyError } },
            Lookup(keys, new object[,] { { 1d, 2d } }, table, "missing", search: search));
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void BatchedExactScanPreservesDuplicateDirectionAndExcelValueTypes(double search, string duplicate)
    {
        var keys = new object[,] { { "key", "KEY", 1d, "1" }, { true, ExcelMissing.Value, "absent", ExcelError.ExcelErrorRef } };
        var lookup = new object[,] { { "KeY", 1d, "1", true, ExcelEmpty.Value, ExcelError.ExcelErrorNA, "keY" } };
        var values = new object[,] { { "first", "number", "text", "logical", "blank", "error", "last" } };

        Assert.Equal(new object[,] { { duplicate, duplicate, "number", "text" },
            { "logical", "blank", "fallback", ExcelError.ExcelErrorRef } },
            Lookup(keys, lookup, values, "fallback", search: search));
    }

    [Theory]
    [InlineData(1, "first")]
    [InlineData(-1, "last")]
    public void BatchedApproximateScanPicksNearestValuesAndKeepsDuplicateDirection(double search, string duplicate)
    {
        var keys = new object[,] { { -1d, 2d, 5d, 6d }, { 12d, ExcelMissing.Value, ExcelError.ExcelErrorRef, 11d } };
        var lookup = new object[,] { { 9d, 1d, 5d, 5d, 7d, ExcelEmpty.Value, ExcelError.ExcelErrorNA } };
        var values = new object[,] { { "nine", "one", "first", "last", "seven", "blank", "error" } };

        Assert.Equal(new object[,] { { "one", duplicate, duplicate, "seven" },
            { "error", "blank", ExcelError.ExcelErrorRef, "error" } },
            Lookup(keys, lookup, values, "missing", match: 1, search: search));
        Assert.Equal(new object[,] { { "missing", "one", duplicate, duplicate },
            { "nine", "blank", ExcelError.ExcelErrorRef, "nine" } },
            Lookup(keys, lookup, values, "missing", match: -1, search: search));
    }

    [Fact]
    public void InvalidShapesAndModesReturnValueErrorEvenWithFallback()
    {
        var keys = new object[,] { { "a", "b" } };
        var values = new object[,] { { 1d, 2d } };
        var ambiguous = new object[,] { { "a", "b" }, { "c", "d" } };
        var multipleAreas = new ExcelReference(new[] { new[] { 0, 0, 0, 0 }, new[] { 0, 0, 2, 2 } }, new IntPtr(42));

        Assert.Equal(ExcelError.ExcelErrorValue, Lookup("a", ambiguous, values, "missing"));
        Assert.Equal(ExcelError.ExcelErrorValue, Lookup("a", keys, 1d, "missing"));
        Assert.Equal(ExcelError.ExcelErrorValue, Lookup("a", keys, values, "missing", match: 3));
        Assert.Equal(ExcelError.ExcelErrorValue, Lookup("a", keys, values, "missing", search: 0));
        Assert.Equal(ExcelError.ExcelErrorValue, Lookup("a*", keys, values, "missing", match: 2, search: 2));
        Assert.Equal(ExcelError.ExcelErrorValue, Lookup(multipleAreas, keys, values, "missing"));
    }

    [Fact]
    public void OneElementKeyArrayReturnsOnlyTheLeadingCell()
    {
        var lookup = new object[,] { { 1d }, { 2d } };
        var returns = new object[,] { { 10d, 20d }, { 30d, 40d } };

        Assert.Equal(new object[,] { { 10d, 20d } }, Lookup(1d, lookup, returns));
        Assert.Equal(new object[,] { { 10d } }, Lookup(new object[,] { { 1d } }, lookup, returns));
    }

    [Fact]
    public void InvalidLookupAndModeArgumentsKeepTheShapeAndErrorsOfTheKeys()
    {
        var keys = new object[,] { { 1d }, { ExcelError.ExcelErrorDiv0 }, { 3d } };
        var expectedKeys = (object[,])keys.Clone();
        var worksheetKeys = (object[,])keys.Clone();
        var lookup = new object[,] { { 1d }, { 2d } };
        var invalidLookup = new object[,] { { 1d, 2d }, { 3d, 4d } };
        var expected = new object[,] { { ExcelError.ExcelErrorValue }, { ExcelError.ExcelErrorDiv0 },
            { ExcelError.ExcelErrorValue } };

        Assert.Equal(expected, Lookup(keys, invalidLookup, lookup));
        Assert.Equal(new object[,] { { ExcelError.ExcelErrorValue } },
            Lookup(new object[,] { { 1d } }, invalidLookup, lookup));
        Assert.Equal(expected, Lookup(keys, lookup, 10d));
        Assert.Equal(expected, Lookup(keys, lookup, lookup, search: 0));
        Assert.Equal(expected, LookupFunctions.XLookup(keys, ExcelMissing.Value, ExcelError.ExcelErrorNA, "missing",
            0d, new object[,] { { 1d }, { -1d } }));
        Assert.Equal(new object[,] { { ExcelError.ExcelErrorDiv0 }, { ExcelError.ExcelErrorDiv0 },
            { ExcelError.ExcelErrorDiv0 }, { ExcelError.ExcelErrorNA } },
            LookupFunctions.XLookup(keys, ExcelError.ExcelErrorDiv0, lookup, "missing",
                0d, new object[,] { { 1d }, { -1d }, { 1d }, { -1d } }));
        Assert.Equal(new object[,] { { ExcelError.ExcelErrorNA }, { ExcelError.ExcelErrorDiv0 }, { ExcelError.ExcelErrorNA } },
            Lookup(keys, lookup, ExcelError.ExcelErrorNA));
        Assert.Equal(new object[,] { { ExcelError.ExcelErrorRef }, { ExcelError.ExcelErrorDiv0 }, { ExcelError.ExcelErrorRef } },
            LookupFunctions.XLookup(new ExcelReference(0, 2, 0, 0, new IntPtr(42)), lookup, lookup, "missing",
                ExcelError.ExcelErrorRef, 1d, _ => worksheetKeys, is64BitProcess: true));
        Assert.Equal(expectedKeys, keys);
        Assert.Equal(expectedKeys, worksheetKeys);
    }

    [Fact]
    public void DisjointLookupAndReturnRangesReturnNAEvenWithFallback()
    {
        var multipleAreas = new ExcelReference(new[] { new[] { 0, 0, 0, 0 }, new[] { 0, 0, 2, 2 } }, new IntPtr(42));
        var keys = new object[,] { { "a", "b" } };
        var values = new object[,] { { 1d, 2d } };

        Assert.Equal(ExcelError.ExcelErrorNA, Lookup("a", multipleAreas, values, "missing"));
        Assert.Equal(ExcelError.ExcelErrorNA, Lookup("a", keys, multipleAreas, "missing"));
        Assert.Equal(ExcelError.ExcelErrorValue, Lookup("a", multipleAreas, values, "missing", match: 5));
    }
}
