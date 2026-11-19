using System.Numerics;
using System.Text;

namespace XLPolyfill;

internal sealed class ExcelWildcard
{
    private readonly List<Segment> segments = [];
    private readonly bool startsWithStar;
    private readonly bool endsWithStar;

    public ExcelWildcard(string pattern)
    {
        // Excel treats a standalone CHAR(1) as a wildcard for any text.
        pattern = pattern == "\u0001" ? "*" : WindowsText.FoldWildcardText(pattern);
        startsWithStar = pattern.StartsWith('*');
        var literals = new List<Literal>();
        var literal = new StringBuilder();
        int width = 0;
        int totalWidth = 0;
        long literalBytes = 0;

        for (int index = 0; index < pattern.Length; index++)
        {
            char token = pattern[index];
            if (token == '~' && index == pattern.Length - 1)
                break;
            endsWithStar = token == '*';
            switch (token)
            {
                case '*':
                    FinishSegment();
                    break;
                case '?':
                    FinishLiteral();
                    width++;
                    break;
                default:
                    if (token == '~')
                        token = pattern[++index];
                    literal.Append(token);
                    width++;
                    break;
            }
        }
        FinishSegment();
        MinimumLength = totalWidth;
        // Allow for the pattern's index entries and unused capacity in the segment array.
        EstimatedBytes = 256 + segments.Capacity * 24L + literalBytes;

        void FinishLiteral()
        {
            if (literal.Length == 0)
                return;
            string text = literal.ToString();
            literals.Add(new Literal(width - literal.Length, text));
            literalBytes += 64L + text.Length * sizeof(char);
            if (text.Length > (RequiredLiteral?.Length ?? 0))
            {
                AlternateLiteral = RequiredLiteral;
                RequiredLiteral = text;
            }
            else if (text != RequiredLiteral && text.Length > (AlternateLiteral?.Length ?? 0))
                AlternateLiteral = text;
            literal.Clear();
        }

        void FinishSegment()
        {
            FinishLiteral();
            if (width == 0)
                return;
            segments.Add(new Segment(width, literals.ToArray()));
            literals.Clear();
            totalWidth += width;
            width = 0;
        }
    }

    public string? StartingLiteral => startsWithStar || segments.Count == 0 ? null : segments[0].StartingLiteral;
    public string? EndingLiteral => endsWithStar || segments.Count == 0 ? null : segments[^1].EndingLiteral;
    public string? RequiredLiteral { get; private set; }
    public string? AlternateLiteral { get; private set; }
    public int MinimumLength { get; }
    public long EstimatedBytes { get; }

    public static bool SpendCharacters(ref int remainingWork, int characters) =>
        SpendWork(ref remainingWork, 1 + (characters - 1) / 64);

    private static bool SpendWork(ref int remainingWork, int units)
    {
        remainingWork -= units;
        return remainingWork >= 0;
    }

    public bool MatchesFolded(ReadOnlySpan<char> text, ref int remainingWork)
    {
        if (text.Length < MinimumLength)
            return false;
        if (segments.Count == 0)
            return startsWithStar || text.IsEmpty;
        if (!startsWithStar && !endsWithStar && segments.Count == 1)
            return text.Length == segments[0].Width && segments[0].MatchesAtStart(text, ref remainingWork);

        int first = 0;
        int last = segments.Count;
        if (!endsWithStar)
        {
            var suffix = segments[--last];
            if (suffix.Width > text.Length || !suffix.MatchesAtStart(text[^suffix.Width..], ref remainingWork))
                return false;
            text = text[..^suffix.Width];
        }
        if (!startsWithStar)
        {
            var prefix = segments[first++];
            if (!prefix.MatchesAtStart(text, ref remainingWork))
                return false;
            text = text[prefix.Width..];
        }

        for (int index = first; index < last; index++)
        {
            if (!SpendWork(ref remainingWork, 1))
                return false;
            var segment = segments[index];
            int offset = segment.Find(text, ref remainingWork);
            if (offset < 0)
                return false;
            text = text[(offset + segment.Width)..];
        }
        return true;
    }

    private readonly record struct Literal(int Offset, string Text);

    private readonly struct Segment
    {
        private readonly Literal[] literals;
        private readonly int anchor;

        public Segment(int width, Literal[] literals)
        {
            Width = width;
            this.literals = literals;
            for (int index = 1; index < literals.Length; index++)
                if (literals[index].Text.Length > literals[anchor].Text.Length)
                    anchor = index;
        }

        public int Width { get; }
        public string? StartingLiteral => literals.Length > 0 && literals[0].Offset == 0 ? literals[0].Text : null;
        public string? EndingLiteral => literals.Length > 0 && literals[^1].Offset + literals[^1].Text.Length == Width
            ? literals[^1].Text : null;

        public bool MatchesAtStart(ReadOnlySpan<char> text, ref int remainingWork) =>
            text.Length >= Width && FindMismatch(text, -1, ref remainingWork) < 0;

        private int FindMismatch(ReadOnlySpan<char> text, int checkedLiteral, ref int remainingWork)
        {
            for (int index = 0; index < literals.Length; index++)
            {
                if (index == checkedLiteral)
                    continue;
                var literal = literals[index];
                if (!SpendCharacters(ref remainingWork, literal.Text.Length) ||
                    !text[literal.Offset..].StartsWith(literal.Text, StringComparison.Ordinal))
                    return index;
            }
            return -1;
        }

        public int Find(ReadOnlySpan<char> text, ref int remainingWork)
        {
            int lastStart = text.Length - Width;
            if (lastStart < 0)
                return -1;
            if (literals.Length == 0)
                return 0;

            int start = FindLiteral(text, literals[anchor], 0, lastStart, ref remainingWork);
            int retries = 0;
            while (start >= 0)
            {
                int mismatch = FindMismatch(text[start..], anchor, ref remainingWork);
                if (mismatch < 0)
                    return start;
                if (remainingWork < 0)
                    return -1;

                // No later match can start before the failing literal can line up again.
                if (++retries == 8 && literals.Length > 1 && lastStart - start >= 64)
                    return FindUsingMasks(text, start + 1, lastStart, mismatch, ref remainingWork);
                start = FindLiteral(text, literals[mismatch], start + 1, lastStart, ref remainingWork);
                if (start >= 0)
                    start = FindLiteral(text, literals[anchor], start, lastStart, ref remainingWork);
            }
            return -1;
        }

        private int FindUsingMasks(ReadOnlySpan<char> text, int start, int lastStart, int firstLiteral, ref int remainingWork)
        {
            int count = lastStart - start + 1;
            var candidates = new ulong[(count + 63) / 64];
            Array.Fill(candidates, ulong.MaxValue);
            candidates[^1] >>= (64 - count % 64) % 64;
            var masks = new Dictionary<string, ulong[]>(StringComparer.Ordinal);
            ulong[]? scratch = null;

            // Repeated literals share one occurrence mask. Shift by each literal's required offset.
            for (int index = 0; index < literals.Length; index++)
            {
                var literal = literals[(firstLiteral + index) % literals.Length];
                if (!SpendWork(ref remainingWork, candidates.Length))
                    return -1;
                if (!masks.TryGetValue(literal.Text, out var occurrences))
                {
                    if (masks.Count < 16)
                    {
                        occurrences = new ulong[(text.Length + 63) / 64];
                        masks.Add(literal.Text, occurrences);
                    }
                    else
                    {
                        occurrences = scratch ??= new ulong[(text.Length + 63) / 64];
                        Array.Clear(occurrences);
                    }

                    int position = 0;
                    while (position <= text.Length - literal.Text.Length)
                    {
                        int offset = FindText(text[position..], literal.Text, ref remainingWork);
                        if (remainingWork < 0)
                            return -1;
                        if (offset < 0)
                            break;
                        position += offset;
                        occurrences[position / 64] |= 1UL << (position % 64);
                        position++;
                    }
                }

                int word = (start + literal.Offset) / 64;
                int shift = (start + literal.Offset) % 64;
                ulong any = 0;
                for (int block = 0; block < candidates.Length; block++, word++)
                {
                    ulong aligned = occurrences[word] >> shift;
                    if (shift != 0 && word + 1 < occurrences.Length)
                        aligned |= occurrences[word + 1] << (64 - shift);
                    candidates[block] &= aligned;
                    any |= candidates[block];
                }
                if (any == 0)
                    return -1;
            }

            for (int block = 0; block < candidates.Length; block++)
                if (candidates[block] != 0)
                    return start + block * 64 + BitOperations.TrailingZeroCount(candidates[block]);
            return -1;
        }

        private static int FindLiteral(ReadOnlySpan<char> text, Literal literal, int start, int lastStart, ref int remainingWork)
        {
            if (start > lastStart)
                return -1;
            var candidates = text.Slice(start + literal.Offset, lastStart - start + literal.Text.Length);
            int offset = FindText(candidates, literal.Text, ref remainingWork);
            return offset < 0 ? -1 : start + offset;
        }

        private static int FindText(ReadOnlySpan<char> text, string literal, ref int remainingWork)
        {
            int offset = text.IndexOf(literal, StringComparison.Ordinal);
            return SpendCharacters(ref remainingWork, offset < 0 ? text.Length : offset + literal.Length) ? offset : -1;
        }
    }
}
