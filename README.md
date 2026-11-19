# XLPolyfill

An Excel-DNA add-in that backports modern Excel functions to Excel 2016 and Excel 2019.

## Installation

### Prerequisites

- Windows with the [.NET 11 Desktop Runtime](https://learn.microsoft.com/en-us/dotnet/core/install/windows#supported-versions) matching Excel's bitness (x64 or x86).
  > Excel can load only one modern .NET runtime per process; another add-in targeting a different major .NET version can prevent XLPolyfill from loading.
- Building from source requires the .NET 11 SDK on Windows.

### Build and load

1. Build the project in Release configuration:

   ```powershell
   dotnet build src\XLPolyfill.csproj -c Release
   ```

2. Pick the packed `.xll` in `src\bin\Release\net11.0-windows\publish\` that matches your Excel bitness:
   - `XLPolyfill-AddIn64-packed.xll` for 64-bit Excel
   - `XLPolyfill-AddIn-packed.xll` for 32-bit Excel

3. Open the `.xll` directly in Excel, or register it under **File > Options > Add-ins > Manage: Excel Add-ins > Go... > Browse...**.

## Implemented functions

- `XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])`
  - Supports `match_mode` `0` (exact, default), `-1` (exact or next smaller), `1` (exact or next larger), and `2` (wildcard). Mode `3` (regex) returns `#VALUE!`.
  - To look up multiple keys against the same range, pass a range for `lookup_value` (e.g., `=XLOOKUP(E2:E101,A2:A10001,B2:B10001,"missing")`), select the output range, and press **Ctrl+Shift+Enter**. Use `search_mode` `2` (ascending) or `-2` (descending) on sorted ranges.
- `FILTER(array, include, [if_empty])`
  - If no cells match and `if_empty` is omitted, XLPolyfill returns `#N/A` while native Excel returns `#CALC!`. Pass a fallback such as `""` so both return the same value.

### Multi-cell formulas and workbook compatibility

- Excel 2016 and 2019 do not support dynamic array spilling, so select the target output range and press **Ctrl+Shift+Enter** to enter an array formula.
- Workbooks saved in newer Excel may show `_xlfn.XLOOKUP` or `_xlfn._xlws.FILTER` as `#NAME?`; re-enter those formulas as `XLOOKUP` or `FILTER` after loading XLPolyfill.
- In newer Excel without XLPolyfill, XLPolyfill formulas may show `#NAME?`; re-enter them as built-in functions or paste values before sharing.

## Development and testing

Run the managed test suite, which also builds and checks both packed `.xll` files:

```powershell
dotnet test tests\XLPolyfill.Tests\XLPolyfill.Tests.csproj -c Release
```

Run the Excel COM smoke test against a local Excel installation (pass `32` for 32-bit Excel):

```powershell
dotnet build tests\XLPolyfill.ExcelSmoke\XLPolyfill.ExcelSmoke.csproj -c Release
.\tests\Invoke-ExcelSmoke.ps1 -ExcelBitness 64
```

## License

[MIT](LICENSE.txt)
