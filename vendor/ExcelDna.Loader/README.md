# Excel-DNA 1.9.0 loader with array cleanup fix

`ExcelDna.Loader.dll` is an altered build of [Excel-DNA v1.9.0](https://github.com/Excel-DNA/ExcelDna/tree/fb3c19e2def68c6c64d6fece13daa9efd36a2748), targeting `net6.0-windows` with the upstream strong name and assembly version. The [patch](array-cleanup.patch) moves the cleanup flag from the last array element to the returned array. Excel can then call `xlAutoFree12` after it copies the result. The license is in [LICENSE.txt](LICENSE.txt).

Rebuild on Windows with the .NET SDK `11.0.100`. Keep Git's original line endings so the source checksums embedded by the compiler match. Run from the XLPolyfill repository root.

```powershell
$vendor = (Resolve-Path 'vendor/ExcelDna.Loader').Path
$checkout = Join-Path ([IO.Path]::GetTempPath()) ('ExcelDna-' + [guid]::NewGuid())
git clone --no-checkout --depth 1 --branch v1.9.0-Release -c core.autocrlf=false https://github.com/Excel-DNA/ExcelDna.git $checkout
if ($LASTEXITCODE -ne 0) { throw 'Excel-DNA clone failed.' }
git -C $checkout checkout --detach fb3c19e2def68c6c64d6fece13daa9efd36a2748
if ($LASTEXITCODE -ne 0) { throw 'Excel-DNA checkout failed.' }
git -C $checkout apply (Join-Path $vendor 'array-cleanup.patch')
if ($LASTEXITCODE -ne 0) { throw 'Excel-DNA loader patch failed.' }
Push-Location $checkout
try {
    if ((dotnet --version) -ne '11.0.100') { throw 'This rebuild requires .NET SDK 11.0.100.' }
    dotnet build Source/ExcelDna.Loader/ExcelDna.Loader.csproj -c Release -f net6.0-windows -p:AssemblyVersion=1.1.0.0 -p:ContinuousIntegrationBuild=true
    if ($LASTEXITCODE -ne 0) { throw 'Excel-DNA loader build failed.' }
    $rebuilt = Get-FileHash 'Source/ExcelDna.Loader/bin/Release/net6.0-windows/ExcelDna.Loader.dll'
    $vendored = Get-FileHash (Join-Path $vendor 'ExcelDna.Loader.dll')
    if ($rebuilt.Hash -ne $vendored.Hash) { throw "Rebuilt loader checksum differs from the vendored DLL ($($rebuilt.Hash))." }
} finally {
    Pop-Location
}
```

Run this check before releasing a loader update. The checked-in DLL has SHA-256 `7e3b4a93481f74e2e8ab58d9459d58f18ca3ba9dab6ea14056346822863d832b`.

Remove this directory and the project's pack include when an official loader passes the array cleanup regression test.
