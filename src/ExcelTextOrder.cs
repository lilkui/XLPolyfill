using System.Buffers;
using System.Text;
using System.Text.RegularExpressions;

namespace XLPolyfill;

internal readonly partial struct ExcelTextOrder
{
    // Distinguish symbols that Windows can collapse when its comparison ties.
    private const string Symbols = "0123456789 !\"#$%&()*,./:;?@[\\]^_`{|}~+<=>\u200B";
    private const string SubscriptDigits = "₀₁₂₃₄₅₆₇₈₉";
    private const string CircledDigits = "⓪①②③④⑤⑥⑦⑧⑨";
    private const string SquaredAndCubedUnitSymbols = "㎠㎡㎤㎥";
    private const string CollationMarks =
        "\u034F\u061C\u200C\u200D\u200E\u200F\u202A\u202B\u202C\u202D\u202E\u2060\u2066\u2067\u2068\u2069\uFE0E\uFE0F";
    private static readonly SearchValues<char> SortSymbols = SearchValues.Create(Symbols);
    private static readonly SearchValues<char> SortMarks = SearchValues.Create(CollationMarks);
    private static readonly SearchValues<char> UnitSymbols = SearchValues.Create(SquaredAndCubedUnitSymbols);
    private static readonly SearchValues<char> NormalizeCharacters = SearchValues.Create(
        Enumerable.Range(0xFF01, 94).Concat(Enumerable.Range(0xFF61, 63)).Select(code => (char)code)
            .Concat("\u00AD\u037E\u0640\u2024\u205F\u3000\uFE62\uFE63\uFE64\uFE65\uFE66\uFEFF" +
                "⁺₊⓿❶❷❸❹❺❻❼❽❾⓵⓶⓷⓸⓹⓺⓻⓼⓽➀➁➂➃➄➅➆➇➈➊➋➌➍➎➏➐➑➒" +
                WindowsText.SuperscriptDigits + SubscriptDigits + CircledDigits).ToArray());
    private readonly string exactText;
    private readonly string sortText;
    private readonly bool isEmpty;

    public ExcelTextOrder(string text, string localeName)
    {
        isEmpty = text.Length == 0;
        exactText = Normalize(text);
        if (exactText.AsSpan().ContainsAny(SortMarks))
        {
            foreach (char mark in CollationMarks)
            {
                if (!exactText.Contains(mark))
                    continue;
                // Keep marks that affect collation in the current locale.
                string withoutMark = exactText.Replace(mark.ToString(), "");
                if (WindowsText.Compare(exactText, withoutMark, localeName) == 0)
                    exactText = withoutMark;
            }
        }

        sortText = exactText.Replace("'", "").Replace("-", "");
    }

    public int GetSymbolHashCode()
    {
        // The Windows tie-breaker requires the same sequence of explicit symbols for equality.
        var hash = new HashCode();
        ReadOnlySpan<char> remaining = sortText;
        int index;
        while ((index = remaining.IndexOfAny(SortSymbols)) >= 0)
        {
            hash.Add(remaining[index]);
            remaining = remaining[(index + 1)..];
        }
        return hash.ToHashCode();
    }

    public long AdditionalTextBytes(string original) =>
        (ReferenceEquals(original, exactText) ? 0 : StringBytes(exactText)) +
        (ReferenceEquals(exactText, sortText) || ReferenceEquals(original, sortText) ? 0 : StringBytes(sortText));

    public static long StringBytes(string value) => 24L + value.Length * (long)sizeof(char);

    public int CompareTo(ExcelTextOrder other, string localeName)
    {
        int order = WindowsText.Compare(sortText, other.sortText, localeName);
        if (order == 0 && sortText != other.sortText)
            order = CompareWindowsTies(sortText, other.sortText, localeName);
        if (order != 0)
            return order;
        if (isEmpty != other.isEmpty)
            return isEmpty ? -1 : 1;

        // Word sorting breaks ties using the position and order of apostrophes and hyphens.
        bool hasPunctuation = exactText.Length != sortText.Length;
        bool otherHasPunctuation = other.exactText.Length != other.sortText.Length;
        int punctuationOrder = hasPunctuation.CompareTo(otherHasPunctuation);
        return punctuationOrder != 0 ? punctuationOrder : WindowsText.Compare(exactText, other.exactText, localeName);
    }

    private static int CompareWindowsTies(ReadOnlySpan<char> left, ReadOnlySpan<char> right, string localeName)
    {
        while (!left.IsEmpty && !right.IsEmpty)
        {
            int leftSymbol = Symbols.IndexOf(left[0]);
            int rightSymbol = Symbols.IndexOf(right[0]);
            if (leftSymbol >= 0 || rightSymbol >= 0)
            {
                if (leftSymbol != rightSymbol)
                {
                    // An ignorable trailing run sorts like the end of the string.
                    if (leftSymbol < 0)
                        return IsIgnorableTail(left, localeName) ? -1 : 1;
                    if (rightSymbol < 0)
                        return IsIgnorableTail(right, localeName) ? 1 : -1;
                    return leftSymbol - rightSymbol;
                }
                left = left[1..];
                right = right[1..];
                continue;
            }

            int leftLength = left.IndexOfAny(SortSymbols);
            int rightLength = right.IndexOfAny(SortSymbols);
            leftLength = leftLength < 0 ? left.Length : leftLength;
            rightLength = rightLength < 0 ? right.Length : rightLength;
            int order = WindowsText.Compare(left[..leftLength], right[..rightLength], localeName);
            if (order != 0)
                return order;
            left = left[leftLength..];
            right = right[rightLength..];
        }

        return left.Length != right.Length && !IsIgnorableTail(left.IsEmpty ? right : left, localeName)
            ? left.Length.CompareTo(right.Length)
            : 0;
    }

    private static bool IsIgnorableTail(ReadOnlySpan<char> text, string localeName) =>
        !char.IsAsciiLetter(text[0]) && !text.ContainsAny(SortSymbols) && WindowsText.Compare(text, [], localeName) == 0;

    private static string Normalize(string value)
    {
        if (value.AsSpan().ContainsAny(UnitSymbols))
            value = SquaredAndCubedUnits().Replace(value, static unit => unit.Value.Normalize(NormalizationForm.FormKC));
        if (!value.AsSpan().ContainsAny(NormalizeCharacters))
            return value;

        Span<char> characters = value.Length <= 256 ? stackalloc char[value.Length] : new char[value.Length];
        int length = 0;
        bool hasHalfWidthKana = false;
        foreach (char character in value)
        {
            if (character is '\u00AD' or '\u0640' or '\uFEFF')
                continue;
            hasHalfWidthKana |= character is >= '\uFF61' and <= '\uFF9F';
            characters[length++] = character switch
            {
                '\u037E' => ';',
                '\u2024' => '.',
                '\u205F' or '\u3000' => ' ',
                >= '\uFF01' and <= '\uFF5E' => (char)(character - 0xFEE0),
                >= '₀' and <= '₉' => (char)('0' + character - '₀'),
                >= '①' and <= '⑨' => (char)('1' + character - '①'),
                >= '⓵' and <= '⓽' => (char)('1' + character - '⓵'),
                >= '❶' and <= '❾' => (char)('1' + character - '❶'),
                >= '➀' and <= '➈' => (char)('1' + character - '➀'),
                >= '➊' and <= '➒' => (char)('1' + character - '➊'),
                '⓪' or '⓿' => '0',
                '⁰' or '¹' or '²' or '³' or (>= '⁴' and <= '⁹') => (char)('0' + WindowsText.SuperscriptDigits.IndexOf(character)),
                '⁺' or '₊' => '+',
                >= '\uFE62' and <= '\uFE66' => WindowsText.SmallPunctuation[character - '\uFE62'],
                _ => character
            };
        }
        string normalized = new(characters[..length]);
        return hasHalfWidthKana
            ? HalfWidthKana().Replace(normalized, static run => run.Value.Normalize(NormalizationForm.FormKC))
            : normalized;
    }

    [GeneratedRegex("[\\uFF61-\\uFF9F]+")]
    private static partial Regex HalfWidthKana();

    [GeneratedRegex("[" + SquaredAndCubedUnitSymbols + "]")]
    private static partial Regex SquaredAndCubedUnits();
}
