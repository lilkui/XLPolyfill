using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

namespace XLPolyfill;

internal static partial class WindowsText
{
    internal const string SmallPunctuation = "+-<>=";
    internal const string SuperscriptDigits = "⁰¹²³⁴⁵⁶⁷⁸⁹";
    private const uint IgnoreCaseAndWidth = 0x0002_0001;
    private const uint LinguisticUppercase = 0x0100_0200;
    private const int InsufficientBuffer = 122;

    public static unsafe int Compare(ReadOnlySpan<char> left, ReadOnlySpan<char> right, string localeName)
    {
        if (left.SequenceEqual(right))
            return 0;
        if (left.IsEmpty)
            left = "\0";
        if (right.IsEmpty)
            right = "\0";

        fixed (char* leftPointer = left, rightPointer = right)
        {
            int result = CompareStringEx(localeName, IgnoreCaseAndWidth, leftPointer, left.Length,
                rightPointer, right.Length, 0, 0, 0);
            return result == 0 ? throw new Win32Exception(Marshal.GetLastPInvokeError()) : result - 2;
        }
    }

    public static string FoldWildcardText(string value) => Ascii.IsValid(value)
        ? value.ToUpperInvariant()
        : string.Create(value.Length, value, static (destination, source) => FoldWildcardText(source, destination));

    private static unsafe void FoldWildcardText(ReadOnlySpan<char> source, Span<char> destination)
    {
        fixed (char* input = source)
        fixed (char* output = destination)
        {
            int written = LCMapStringEx("", LinguisticUppercase, input, source.Length, output, destination.Length, 0, 0, 0);
            if (written != source.Length)
            {
                if (written == 0 && Marshal.GetLastPInvokeError() != InsufficientBuffer)
                    throw new Win32Exception(Marshal.GetLastPInvokeError());

                // Preserve '?' offsets if a case mapping would change the UTF-16 length.
                source.CopyTo(destination);
                char* mapped = stackalloc char[2];
                for (int index = 0; index < source.Length; index++)
                {
                    char character = source[index];
                    if (char.IsSurrogate(character))
                        continue;

                    if (char.IsAscii(character))
                    {
                        destination[index] = char.ToUpperInvariant(character);
                        continue;
                    }

                    written = LCMapStringEx("", LinguisticUppercase, input + index, 1, mapped, 2, 0, 0, 0);
                    if (written == 1)
                        destination[index] = mapped[0];
                    else if (written == 0 && Marshal.GetLastPInvokeError() != InsufficientBuffer)
                        throw new Win32Exception(Marshal.GetLastPInvokeError());
                }
            }
        }

        // Excel folds wildcard literals by UTF-16 code unit and preserves fullwidth syntax.
        for (int index = 0; index < source.Length; index++)
        {
            char character = source[index];
            destination[index] = character switch
            {
                >= '\uD800' and <= '\uDFFF' => character,
                '\u00B5' or '\u0131' or '\u017F' or '\u0345' => character,
                '\u00DF' => '\u1E9E',
                'K' => 'K',
                'Ω' => 'Ω',
                'Å' => 'Å',
                '\u037E' => ';',
                '\u2024' => '.',
                '\u205F' => ' ',
                '\u200D' => '\u200C',
                >= '﹢' and <= '﹦' => SmallPunctuation[character - '﹢'],
                '⓿' => '0',
                '⓪' => '⁰',
                >= '①' and <= '⑨' => SuperscriptDigits[character - '①' + 1],
                >= '➊' and <= '➒' => SuperscriptDigits[character - '➊' + 1],
                >= '❶' and <= '❾' => (char)('₁' + character - '❶'),
                >= '➀' and <= '➈' => (char)('1' + character - '➀'),
                _ => destination[index]
            };
        }
    }

    [LibraryImport("kernel32.dll", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    private static unsafe partial int CompareStringEx(string localeName, uint flags, char* left, int leftLength,
        char* right, int rightLength, nint version, nint reserved, nint sortHandle);

    [LibraryImport("kernel32.dll", StringMarshalling = StringMarshalling.Utf16, SetLastError = true)]
    [UnmanagedCallConv(CallConvs = [typeof(CallConvStdcall)])]
    private static unsafe partial int LCMapStringEx(string localeName, uint flags, char* source, int sourceLength,
        void* destination, int destinationLength, nint version, nint reserved, nint sortHandle);
}
