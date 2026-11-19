using System.Buffers;
using ExcelDna.Integration;

namespace XLPolyfill;

public static class FilterFunctions
{
    private const int WorksheetReadBytes = 64 * 1024 * 1024;
    private const int MaxWorksheetReadCells = (WorksheetReadBytes - ExcelArrayResult.NativeCellBytes) /
        (ExcelArrayResult.NativeCellBytes + ExcelArrayResult.MaxCellTextBytes);
    private static readonly object InvalidValue = ExcelError.ExcelErrorValue;
    private static readonly object NoMatches = ExcelError.ExcelErrorNA;
    private static readonly object Zero = 0d;

    [ExcelFunction(Name = "FILTER", Category = "Lookup & Reference",
        Description = "Returns the rows or columns selected by an include array.")]
    public static object Filter(
        [ExcelArgument(Name = "array", Description = "The range or array to filter.", AllowReference = true)]
        object source,
        [ExcelArgument(Name = "include", Description = "A column to select rows or a row to select columns.", AllowReference = true)]
        object include,
        [ExcelArgument(Name = "[if_empty]", Description = "The result when nothing matches. Defaults to #N/A.", AllowReference = true)]
        object ifEmpty) =>
        Filter(source, include, ifEmpty, static block => block.GetValue(), Environment.Is64BitProcess);

    internal static object Filter(object source, object include, object ifEmpty, Func<ExcelReference, object> readBlock,
        bool is64BitProcess)
    {
        if (source is ExcelError)
            return source;
        if (include is ExcelError)
            return include;
        if (source is ExcelMissing || include is ExcelMissing)
            return InvalidValue;

        (int rows, int columns) = GetShape(source);
        (int includeRows, int includeColumns) = GetShape(include);
        bool byRow = includeColumns == 1 && includeRows == rows;
        if (rows == 0 || columns == 0 || (!byRow && (includeRows != 1 || includeColumns != columns)))
            return (rows == 0 && source is ExcelReference) || (includeRows == 0 && include is ExcelReference)
                ? ExcelError.ExcelErrorRef : InvalidValue;

        int length = byRow ? rows : columns;
        int[]? buffer = null;
        Span<int> selected = length <= 256
            ? stackalloc int[length]
            : (buffer = ArrayPool<int>.Shared.Rent(length)).AsSpan(0, length);
        try
        {
            int count = 0;
            object? includeError = include is ExcelReference includeRange
                ? SelectIncludeRange(includeRange, byRow, readBlock, selected, ref count)
                : SelectIncludeValues(include, byRow, 0, selected, ref count);
            if (includeError is not null)
                return includeError;

            if (count == 0)
                return ifEmpty switch
                {
                    ExcelMissing => NoMatches,
                    ExcelReference fallback => ReadFallbackRange(fallback, readBlock, is64BitProcess),
                    object[,] fallback => ExcelArrayResult.Check(fallback, is64BitProcess),
                    _ => ValueOrZero(ifEmpty)
                };
            if (source is not (object[,] or ExcelReference))
                return source is ExcelEmpty ? new object[,] { { source } } : source;

            int resultRows = byRow ? count : rows;
            int resultColumns = byRow ? columns : count;
            long textBudget = ExcelArrayResult.TextBudget((long)resultRows * resultColumns, is64BitProcess,
                copiesWorksheetValues: source is ExcelReference);
            if (textBudget < 0)
                return ExcelArrayResult.TooLarge;

            if (source is ExcelReference reference)
                return CopyReference(reference, (resultRows, resultColumns), selected[..count], byRow, readBlock, textBudget);

            var result = new object[resultRows, resultColumns];
            CopyMatches((object[,])source, result, selected[..count], byRow, 0, (0, 0), ref textBudget);
            return textBudget < 0 ? ExcelArrayResult.TooLarge : result;
        }
        finally
        {
            if (buffer is not null)
                ArrayPool<int>.Shared.Return(buffer);
        }
    }

    private static (int Rows, int Columns) GetShape(object input) => input switch
    {
        object[,] values => (values.GetLength(0), values.GetLength(1)),
        ExcelReference reference when reference.InnerReferences.Count == 1 =>
            (reference.RowLast - reference.RowFirst + 1, reference.ColumnLast - reference.ColumnFirst + 1),
        ExcelReference => (0, 0),
        _ => (1, 1)
    };

    private static object? SelectIncludeRange(ExcelReference range, bool byRow, Func<ExcelReference, object> readBlock,
        Span<int> selected, ref int count)
    {
        for (int offset = 0; offset < selected.Length;)
        {
            int length = Math.Min(MaxWorksheetReadCells, selected.Length - offset);
            int row = range.RowFirst + (byRow ? offset : 0);
            int column = range.ColumnFirst + (byRow ? 0 : offset);
            var size = byRow ? (Rows: length, Columns: 1) : (Rows: 1, Columns: length);
            var block = new ExcelReference(row, row + size.Rows - 1, column, column + size.Columns - 1, range.SheetId);
            object values = readBlock(block);
            if (GetShape(values) != size)
                return values is ExcelError ? values : InvalidValue;

            object? error = SelectIncludeValues(values, byRow, offset, selected, ref count);
            if (error is not null)
                return error;
            offset += length;
        }
        return null;
    }

    private static object? SelectIncludeValues(object values, bool byRow, int offset, Span<int> selected, ref int count)
    {
        var criteria = values as object[,];
        int length = criteria?.Length ?? 1;
        for (int index = 0; index < length; index++)
        {
            object criterion = criteria is null ? values : criteria[byRow ? index : 0, byRow ? 0 : index];
            switch (criterion)
            {
                case true:
                case string text when text.Equals("TRUE", StringComparison.OrdinalIgnoreCase):
                case double number when number != 0d:
                    selected[count++] = offset + index;
                    break;
                case false:
                case string text when text.Equals("FALSE", StringComparison.OrdinalIgnoreCase):
                case 0d:
                case ExcelEmpty:
                    break;
                case ExcelError:
                    return criterion;
                default:
                    return InvalidValue;
            }
        }
        return null;
    }

    private static object CopyReference(ExcelReference source, (int Rows, int Columns) size, ReadOnlySpan<int> selected,
        bool byRow, Func<ExcelReference, object> readBlock, long textBudget)
    {
        var result = new object[size.Rows, size.Columns];
        int cellsPerSlice = result.GetLength(byRow ? 1 : 0);
        // Excel-DNA copies the strings for every cell in a read before FILTER can check their size.
        int chunkLength = Math.Min(cellsPerSlice, MaxWorksheetReadCells);
        int slicesPerBlock = MaxWorksheetReadCells / chunkLength;
        // shortcut: cap gaps at 512 skipped cells until worksheet timings can guide the cutoff.
        int maxSkippedSlices = 512 / cellsPerSlice;
        for (int start = 0; start < selected.Length;)
        {
            int first = selected[start];
            int end = start + 1;
            while (end < selected.Length && selected[end] - first < slicesPerBlock &&
                selected[end] - selected[end - 1] - 1 <= maxSkippedSlices &&
                selected[end] - first + 1 <= 2 * (end - start + 1))
                end++;

            int sliceCount = selected[end - 1] - first + 1;
            for (int offset = 0; offset < cellsPerSlice; offset += chunkLength)
            {
                int cellsInChunk = Math.Min(chunkLength, cellsPerSlice - offset);
                int sourceRow = source.RowFirst + (byRow ? first : offset);
                int sourceColumn = source.ColumnFirst + (byRow ? offset : first);
                int blockRows = byRow ? sliceCount : cellsInChunk;
                int blockColumns = byRow ? cellsInChunk : sliceCount;
                var block = new ExcelReference(sourceRow, sourceRow + blockRows - 1,
                    sourceColumn, sourceColumn + blockColumns - 1, source.SheetId);
                int resultRow = byRow ? start : offset;
                int resultColumn = byRow ? offset : start;

                object values = readBlock(block);
                if (values is object[,] cells)
                    CopyMatches(cells, result, selected[start..end], byRow, first, (resultRow, resultColumn), ref textBudget);
                else if (blockRows == 1 && blockColumns == 1)
                    result[resultRow, resultColumn] = CopyCell(values, ref textBudget);
                else
                    return values is ExcelError ? values : InvalidValue;

                if (textBudget < 0)
                    return ExcelArrayResult.TooLarge;
            }
            start = end;
        }
        return result;
    }

    private static object ReadFallbackRange(ExcelReference fallback, Func<ExcelReference, object> readBlock, bool is64BitProcess)
    {
        (int rows, int columns) = GetShape(fallback);
        if (rows == 0)
            return InvalidValue;

        if (rows == 1 && columns == 1)
            return ValueOrZero(readBlock(fallback));

        long textBudget = ExcelArrayResult.TextBudget((long)rows * columns, is64BitProcess, copiesWorksheetValues: true);
        if (textBudget < 0)
            return ExcelArrayResult.TooLarge;

        bool byRow = rows <= columns;
        int[] selected = Enumerable.Range(0, byRow ? rows : columns).ToArray();
        return CopyReference(fallback, (rows, columns), selected, byRow, readBlock, textBudget);
    }

    private static void CopyMatches(object[,] values, object[,] result, ReadOnlySpan<int> selected, bool byRow,
        int sourceOffset, (int Row, int Column) resultOffset, ref long textBudget)
    {
        if (byRow)
        {
            int columns = values.GetLength(1);
            for (int row = 0; row < selected.Length; row++)
            {
                int sourceRow = selected[row] - sourceOffset;
                for (int column = 0; column < columns; column++)
                    result[resultOffset.Row + row, resultOffset.Column + column] =
                        CopyCell(values[sourceRow, column], ref textBudget);
            }
            return;
        }

        int rows = values.GetLength(0);
        for (int row = 0; row < rows; row++)
            for (int column = 0; column < selected.Length; column++)
                result[resultOffset.Row + row, resultOffset.Column + column] =
                    CopyCell(values[row, selected[column] - sourceOffset], ref textBudget);
    }

    private static object CopyCell(object value, ref long textBudget)
    {
        textBudget -= ExcelArrayResult.TextBytes(value);
        return value;
    }

    private static object ValueOrZero(object value) => ReferenceEquals(value, ExcelEmpty.Value) ? Zero : value;
}
