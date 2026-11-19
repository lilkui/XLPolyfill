using ExcelDna.Integration;

namespace XLPolyfill;

internal static class ExcelArrayResult
{
    internal const int NativeCellBytes = 32;
    private const int MaxCellTextLength = 32_767;
    internal const int MaxCellTextBytes = (MaxCellTextLength + 1) * sizeof(char);
    internal static readonly object TooLarge = ExcelError.ExcelErrorNum;

    internal static long TextBudget(long cells, bool is64BitProcess, bool copiesWorksheetValues = false)
    {
        // Worksheet reads retain a managed copy while Excel-DNA builds its native buffers.
        long byteLimit = is64BitProcess ? int.MaxValue : (copiesWorksheetValues ? 256 : 512) * 1024L * 1024 - 1;
        return byteLimit - (cells + 1) * NativeCellBytes;
    }

    internal static int TextBytes(object value) =>
        value is string text ? (Math.Min(text.Length, MaxCellTextLength) + 1) * sizeof(char) : 0;

    internal static object Check(object[,] values, bool is64BitProcess)
    {
        long textBudget = TextBudget(values.LongLength, is64BitProcess);
        if (textBudget < 0)
            return TooLarge;
        if (values.LongLength * MaxCellTextBytes <= textBudget)
            return values;

        foreach (object value in values)
        {
            textBudget -= TextBytes(value);
            if (textBudget < 0)
                return TooLarge;
        }
        return values;
    }
}
