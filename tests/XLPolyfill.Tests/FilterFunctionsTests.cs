using ExcelDna.Integration;
using Xunit;

namespace XLPolyfill.Tests;

public class FilterFunctionsTests
{
    [Fact]
    public void FiltersRowsInSourceOrderAndPreservesBlanksAndErrors()
    {
        var source = new object[,]
        {
            { "one", ExcelEmpty.Value },
            { "two", ExcelError.ExcelErrorRef },
            { "three", ExcelError.ExcelErrorDiv0 },
            { "four", 4d }
        };
        var include = new object[,] { { true }, { 0d }, { -2d }, { ExcelEmpty.Value } };

        var result = Assert.IsType<object[,]>(FilterFunctions.Filter(source, include, ExcelError.ExcelErrorNA));

        Assert.Equal(new object[,] { { "one", ExcelEmpty.Value }, { "three", ExcelError.ExcelErrorDiv0 } }, result);
        Assert.Same(ExcelEmpty.Value, source[0, 1]);
    }

    [Fact]
    public void FiltersColumnsInSourceOrder()
    {
        var source = new object[,] { { "a", "b", "c" }, { 1d, 2d, 3d } };
        var include = new object[,] { { false, 1d, true } };

        var result = Assert.IsType<object[,]>(FilterFunctions.Filter(source, include, ExcelMissing.Value));

        Assert.Equal(new object[,] { { "b", "c" }, { 2d, 3d } }, result);
    }

    [Fact]
    public void AcceptsScalarCriteriaForSingleRowsAndColumns()
    {
        var row = new object[,] { { "a", "b" } };
        var column = new object[,] { { "a" }, { "b" } };

        Assert.Equal(row, Assert.IsType<object[,]>(FilterFunctions.Filter(row, true, ExcelMissing.Value)));
        Assert.Equal(column, Assert.IsType<object[,]>(FilterFunctions.Filter(column, true, ExcelMissing.Value)));
        var blank = new object[,] { { ExcelEmpty.Value } };
        Assert.Equal(blank, FilterFunctions.Filter(ExcelEmpty.Value, true, ExcelMissing.Value));
        var reference = new ExcelReference(0, 0, 0, 0, new IntPtr(42));
        Assert.Equal(blank, FilterFunctions.Filter(reference, true, ExcelMissing.Value, _ => ExcelEmpty.Value,
            is64BitProcess: true));
    }

    [Fact]
    public void AcceptsBooleanTextInWorksheetCriteria()
    {
        var source = new object[,] { { "first" }, { "second" }, { "third" } };
        var include = new ExcelReference(10, 12, 20, 20, new IntPtr(42));
        var criteria = new object[,] { { "TRUE" }, { "false" }, { "TrUe" } };

        Assert.Equal(new object[,] { { "first" }, { "third" } },
            FilterFunctions.Filter(source, include, ExcelMissing.Value, ReadCriteria, is64BitProcess: true));
        Assert.Equal(ExcelError.ExcelErrorValue, FilterFunctions.Filter("row", "1", ExcelMissing.Value));
        Assert.Equal(ExcelError.ExcelErrorValue, FilterFunctions.Filter("row", " TRUE ", ExcelMissing.Value));

        object ReadCriteria(ExcelReference block)
        {
            Assert.Equal(include, block);
            return criteria;
        }
    }

    [Fact]
    public void ReturnsTheRequestedFallbackWhenThereAreNoMatches()
    {
        var source = new object[,] { { "a", "b" }, { 1d, 2d } };
        var include = new object[,] { { false }, { 0d } };

        Assert.Equal(ExcelError.ExcelErrorNA, FilterFunctions.Filter(source, include, ExcelMissing.Value));
        Assert.Equal("", FilterFunctions.Filter(source, include, ""));
        Assert.Equal(0d, FilterFunctions.Filter(source, include, ExcelEmpty.Value));
        Assert.Equal(new object[,] { { "none", 0d } },
            FilterFunctions.Filter(source, include, new object[,] { { "none", 0d } }));
    }

    [Fact]
    public void SourceExpressionErrorsTakePrecedenceOverFallbackAndCriteriaErrors()
    {
        const ExcelError sourceError = ExcelError.ExcelErrorDiv0;

        Assert.Equal(sourceError, FilterFunctions.Filter(sourceError, false, "none"));
        Assert.Equal(sourceError, FilterFunctions.Filter(sourceError, ExcelError.ExcelErrorRef, "none"));
    }

    [Fact]
    public void ReadsWorksheetSourceOnlyWhenSelectedAndPreservesItsErrors()
    {
        const ExcelError sourceError = ExcelError.ExcelErrorDiv0;
        var source = new ExcelReference(10, 10, 20, 20, new IntPtr(42));

        Assert.Equal("none", FilterFunctions.Filter(source, false, "none",
            _ => throw new InvalidOperationException("The source must stay unread when there are no matches."),
            is64BitProcess: true));
        Assert.Equal(new object[,] { { sourceError } },
            FilterFunctions.Filter(source, true, "none", block =>
            {
                Assert.Equal(source, block);
                return sourceError;
            }, is64BitProcess: true));
    }

    [Fact]
    public void RejectsInvalidCriteriaAndPropagatesErrorsEvenAfterAMatch()
    {
        var source = new object[,] { { "a", "b" }, { 1d, 2d } };

        Assert.Equal(ExcelError.ExcelErrorValue, FilterFunctions.Filter(source, true, "none"));
        Assert.Equal(ExcelError.ExcelErrorValue,
            FilterFunctions.Filter(source, new object[,] { { true, false }, { true, false } }, "none"));
        Assert.Equal(ExcelError.ExcelErrorValue,
            FilterFunctions.Filter(source, new object[,] { { true }, { "text" } }, "none"));
        Assert.Equal(ExcelError.ExcelErrorDiv0,
            FilterFunctions.Filter(source, new object[,] { { true }, { ExcelError.ExcelErrorDiv0 } }, "none"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ReadsWorksheetIncludeInBoundedBlocks(bool byRow)
    {
        const int length = 1_024;
        var source = new object[byRow ? length : 2, byRow ? 2 : length];
        for (int row = 0; row < source.GetLength(0); row++)
            for (int column = 0; column < source.GetLength(1); column++)
                source[row, column] = CellValue(row, column);
        IntPtr sheet = new(42);
        var include = new ExcelReference(10, byRow ? 1_033 : 10, 20, byRow ? 20 : 1_043, sheet);
        var reads = new List<ExcelReference>();

        var result = Assert.IsType<object[,]>(
            FilterFunctions.Filter(source, include, ExcelMissing.Value, ReadBlock, is64BitProcess: true));

        ExcelReference[] expectedReads = byRow
            ? [new(10, 1_032, 20, 20, sheet), new(1_033, 1_033, 20, 20, sheet)]
            : [new(10, 10, 20, 1_042, sheet), new(10, 10, 1_043, 1_043, sheet)];
        Assert.Equal(expectedReads, reads);
        object[,] expected = byRow
            ? new object[,]
            {
                { source[0, 0], source[0, 1] },
                { source[1_022, 0], source[1_022, 1] },
                { source[1_023, 0], source[1_023, 1] }
            }
            : new object[,]
            {
                { source[0, 0], source[0, 1_022], source[0, 1_023] },
                { source[1, 0], source[1, 1_022], source[1, 1_023] }
            };
        Assert.Equal(expected, result);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            int offset = byRow ? block.RowFirst - include.RowFirst : block.ColumnFirst - include.ColumnFirst;
            int count = byRow ? block.RowLast - block.RowFirst + 1 : block.ColumnLast - block.ColumnFirst + 1;
            if (count == 1)
                return CriterionAt(offset);

            var values = new object[byRow ? count : 1, byRow ? 1 : count];
            for (int index = 0; index < count; index++)
                values[byRow ? index : 0, byRow ? 0 : index] = CriterionAt(offset + index);
            return values;
        }

        static object CriterionAt(int offset) => offset switch
        {
            0 or length - 1 => true,
            length - 2 => -2d,
            _ => ExcelEmpty.Value
        };
    }

    [Theory]
    [InlineData("notes", ExcelError.ExcelErrorValue)]
    [InlineData(ExcelError.ExcelErrorDiv0, ExcelError.ExcelErrorDiv0)]
    public void StopsReadingWorksheetIncludeAtTheFirstError(object firstCriterion, ExcelError expected)
    {
        var source = new ExcelReference(0, 99_999, 0, 7, new IntPtr(41));
        var include = new ExcelReference(0, 99_999, 0, 0, new IntPtr(42));
        var reads = new List<ExcelReference>();
        var values = new object[1_023, 1];
        for (int row = 0; row < values.GetLength(0); row++)
            values[row, 0] = false;
        values[0, 0] = firstCriterion;
        values[1, 0] = ExcelError.ExcelErrorRef;

        var result = FilterFunctions.Filter(source, include, "", block =>
        {
            reads.Add(block);
            return values;
        }, is64BitProcess: false);

        Assert.Equal(expected, result);
        Assert.Equal(new ExcelReference(0, 1_022, 0, 0, include.SheetId), Assert.Single(reads));
    }

    [Fact]
    public void RejectsInvalidWorksheetIncludeShapeWithoutReading()
    {
        var source = new ExcelReference(0, 1, 0, 1, new IntPtr(41));
        var include = new ExcelReference(0, 1, 0, 1, new IntPtr(42));

        var result = FilterFunctions.Filter(source, include, "",
            _ => throw new InvalidOperationException("An invalid include range must stay unread."), is64BitProcess: true);

        Assert.Equal(ExcelError.ExcelErrorValue, result);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DisjointSourceOrIncludeReturnsRefError(bool disjointSource)
    {
        var disjoint = new ExcelReference([[0, 0, 0, 0], [2, 2, 0, 0]], new IntPtr(42));
        object source = disjointSource ? disjoint : new object[,] { { 1d }, { 2d } };
        object include = disjointSource ? new object[,] { { true }, { true } } : disjoint;

        Assert.Equal(ExcelError.ExcelErrorRef, FilterFunctions.Filter(source, include, "none"));
    }

    [Fact]
    public void SkipsReadingFallbackRangeWhenRowsMatch()
    {
        var fallback = new ExcelReference(0, 99_999, 0, 7, new IntPtr(42));

        var result = FilterFunctions.Filter(42d, true, fallback,
            _ => throw new InvalidOperationException("The fallback must stay unread when there is a match."),
            is64BitProcess: true);

        Assert.Equal(42d, result);
    }

    [Fact]
    public void ReadsFallbackRangeInBoundedBlocksWhenNoRowsMatch()
    {
        var fallback = new ExcelReference(10, 11, 20, 2_066, new IntPtr(42));
        var readSizes = new List<int>();

        var result = Assert.IsType<object[,]>(
            FilterFunctions.Filter(42d, false, fallback, ReadBlock, is64BitProcess: false));

        Assert.Equal(new[] { 1_023, 1_023, 1, 1_023, 1_023, 1 }, readSizes);
        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(2_047, result.GetLength(1));
        for (int row = 0; row < result.GetLength(0); row++)
            for (int column = 0; column < result.GetLength(1); column++)
                Assert.Equal(CellValue(10 + row, 20 + column), result[row, column]);

        object ReadBlock(ExcelReference block)
        {
            Assert.Equal(fallback.SheetId, block.SheetId);
            readSizes.Add((block.RowLast - block.RowFirst + 1) * (block.ColumnLast - block.ColumnFirst + 1));
            return ReadNumberCells(block);
        }
    }

    [Fact]
    public void TreatsASingleBlankFallbackCellAsZero()
    {
        var fallback = new ExcelReference(10, 10, 20, 20, new IntPtr(42));

        var result = FilterFunctions.Filter(42d, false, fallback, block =>
        {
            Assert.Equal(fallback, block);
            return ExcelEmpty.Value;
        }, is64BitProcess: true);

        Assert.Equal(0d, result);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SparseWorksheetReadsSkipGapsAndKeepOffsets(bool byRow)
    {
        const int length = 205;
        const int breadth = 8;
        IntPtr sheet = new(42);
        var source = new ExcelReference(10, byRow ? 214 : 17, 20, byRow ? 27 : 224, sheet);
        var include = new object[byRow ? length : 1, byRow ? 1 : length];
        int[] selected = [1, 3, 33, 34, 204];
        for (int i = 0; i < length; i++)
            include[byRow ? i : 0, byRow ? 0 : i] = Array.IndexOf(selected, i) >= 0;
        var reads = new List<ExcelReference>();

        var result = Assert.IsType<object[,]>(
            FilterFunctions.Filter(source, include, ExcelMissing.Value, ReadBlock, is64BitProcess: true));

        ExcelReference[] expectedReads = byRow
            ? [new(11, 13, 20, 27, sheet), new(43, 44, 20, 27, sheet), new(214, 214, 20, 27, sheet)]
            : [new(10, 17, 21, 23, sheet), new(10, 17, 53, 54, sheet), new(10, 17, 224, 224, sheet)];
        Assert.Equal(expectedReads, reads);
        Assert.Equal(byRow ? selected.Length : breadth, result.GetLength(0));
        Assert.Equal(byRow ? breadth : selected.Length, result.GetLength(1));
        for (int row = 0; row < result.GetLength(0); row++)
            for (int column = 0; column < result.GetLength(1); column++)
                Assert.Equal(CellValue(10 + (byRow ? selected[row] : row),
                    20 + (byRow ? column : selected[column])), result[row, column]);

        object ReadBlock(ExcelReference block)
        {
            reads.Add(block);
            return ReadNumberCells(block);
        }
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SplitsLongWorksheetSlicesAndKeepsCellOrder(bool byRow)
    {
        const int length = 2_047;
        IntPtr sheet = new(42);
        var source = new ExcelReference(10, 10 + (byRow ? 2 : length - 1),
            20, 20 + (byRow ? length - 1 : 2), sheet);
        object[,] include = byRow
            ? new object[,] { { false }, { true }, { false } }
            : new object[,] { { false, true, false } };
        var readSizes = new List<int>();

        var result = Assert.IsType<object[,]>(
            FilterFunctions.Filter(source, include, ExcelMissing.Value, ReadBlock, is64BitProcess: true));

        Assert.Equal(new[] { 1_023, 1_023, 1 }, readSizes);
        Assert.Equal(byRow ? 1 : length, result.GetLength(0));
        Assert.Equal(byRow ? length : 1, result.GetLength(1));
        for (int i = 0; i < length; i++)
            Assert.Equal(CellValue(10 + (byRow ? 1 : i), 20 + (byRow ? i : 1)),
                result[byRow ? 0 : i, byRow ? i : 0]);

        object ReadBlock(ExcelReference block)
        {
            Assert.Equal(sheet, block.SheetId);
            readSizes.Add((block.RowLast - block.RowFirst + 1) * (block.ColumnLast - block.ColumnFirst + 1));
            return ReadNumberCells(block);
        }
    }

    [Theory]
    [InlineData(16_384, false)]
    [InlineData(65_536, true)]
    public void OversizedWorksheetResultReturnsNumBeforeReadingCells(int rows, bool is64BitProcess)
    {
        var source = new ExcelReference(0, rows - 1, 0, 1_023, new IntPtr(42));
        var include = new object[rows, 1];
        object yes = true;
        for (int row = 0; row < rows; row++)
            include[row, 0] = yes;

        var result = FilterFunctions.Filter(source, include, ExcelMissing.Value, UnexpectedRead, is64BitProcess);
        var fallback = FilterFunctions.Filter(42d, false, source, UnexpectedRead, is64BitProcess);

        Assert.Equal(ExcelError.ExcelErrorNum, result);
        Assert.Equal(ExcelError.ExcelErrorNum, fallback);

        static object UnexpectedRead(ExcelReference _) =>
            throw new InvalidOperationException("The oversized result should be rejected before reading cells.");
    }

    [Theory]
    [InlineData(5_500, 4_000, false)]
    [InlineData(36_000, 32_000, true)]
    public void RejectsWorksheetTextThatExceedsTheResultLimitAcrossReads(int rows, int minCellsRead, bool is64BitProcess)
    {
        var source = new ExcelReference(10, 10 + rows - 1, 20, 20, new IntPtr(42));
        var include = new object[rows, 1];
        object yes = true;
        for (int row = 0; row < rows; row++)
            include[row, 0] = yes;
        string text = new('x', 32_767);
        int cellsRead = 0;

        var result = FilterFunctions.Filter(source, include, ExcelMissing.Value, ReadBlock, is64BitProcess);

        Assert.Equal(ExcelError.ExcelErrorNum, result);
        Assert.InRange(cellsRead, minCellsRead, rows - 1);

        object ReadBlock(ExcelReference block)
        {
            int length = block.RowLast - block.RowFirst + 1;
            Assert.InRange(length, 1, 1_023);
            cellsRead += length;
            var cells = new object[length, 1];
            for (int row = 0; row < length; row++)
                cells[row, 0] = text;
            return cells;
        }
    }

    [Theory]
    [InlineData(false, 4_094, false)]
    [InlineData(false, 4_095, true)]
    [InlineData(true, 4_094, false)]
    [InlineData(true, 4_095, true)]
    public void WorksheetResultsUseTheLower32BitLimit(bool useFallback, int rows, bool tooLarge)
    {
        var source = new ExcelReference(0, rows - 1, 0, 0, new IntPtr(42));
        var include = new object[rows, 1];
        for (int row = 0; row < rows; row++)
            include[row, 0] = true;
        string text = new('x', 32_767);

        object result = useFallback
            ? FilterFunctions.Filter(42d, false, source, ReadBlock, is64BitProcess: false)
            : FilterFunctions.Filter(source, include, ExcelMissing.Value, ReadBlock, is64BitProcess: false);

        if (tooLarge)
            Assert.Equal(ExcelError.ExcelErrorNum, result);
        else
        {
            var cells = Assert.IsType<object[,]>(result);
            Assert.Equal((rows, 1), (cells.GetLength(0), cells.GetLength(1)));
            Assert.All(cells.Cast<object>(), cell => Assert.Equal(text, cell));
        }

        object ReadBlock(ExcelReference block)
        {
            var cells = new object[block.RowLast - block.RowFirst + 1, 1];
            for (int row = 0; row < cells.Length; row++)
                cells[row, 0] = text;
            return cells;
        }
    }

    [Theory]
    [InlineData(false, 8_192, false)]
    [InlineData(true, 8_192, false)]
    [InlineData(false, 32_768, true)]
    [InlineData(true, 32_768, true)]
    public void ChecksTextBudgetForResultsAndFallbacks(bool useFallback, int rows, bool is64BitProcess)
    {
        var values = CreateLongTextColumn(rows);

        Assert.Equal(ExcelError.ExcelErrorNum, Evaluate());

        for (int row = 0; row < 32; row++)
            values[row, 0] = "short";
        var expected = (object[,])values.Clone();

        Assert.Equal(expected, Evaluate());

        object Evaluate() => useFallback
            ? FilterFunctions.Filter(1d, false, values, UnexpectedRead, is64BitProcess)
            : FilterFunctions.Filter(values, true, ExcelMissing.Value, UnexpectedRead, is64BitProcess);
        static object UnexpectedRead(ExcelReference _) =>
            throw new InvalidOperationException("An in-memory result must not read a worksheet.");
    }

    [Fact]
    public void RegisteredFilterUsesTheRunningProcessResultBudget()
    {
        var values = CreateLongTextColumn(8_192);
        var expected = (object[,])values.Clone();

        var result = FilterFunctions.Filter(values, true, ExcelMissing.Value);

        if (Environment.Is64BitProcess)
            Assert.Equal(expected, result);
        else
            Assert.Equal(ExcelError.ExcelErrorNum, result);
    }

    private static object[,] CreateLongTextColumn(int rows)
    {
        var values = new object[rows, 1];
        string text = new('x', 32_767);
        for (int row = 0; row < rows; row++)
            values[row, 0] = text;
        return values;
    }

    private static object ReadNumberCells(ExcelReference block)
    {
        int rows = block.RowLast - block.RowFirst + 1;
        int columns = block.ColumnLast - block.ColumnFirst + 1;
        if (rows == 1 && columns == 1)
            return CellValue(block.RowFirst, block.ColumnFirst);

        var cells = new object[rows, columns];
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < columns; column++)
                cells[row, column] = CellValue(block.RowFirst + row, block.ColumnFirst + column);
        return cells;
    }

    private static double CellValue(int row, int column) => row * 20_000d + column;
}
