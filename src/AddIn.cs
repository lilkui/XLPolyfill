using ExcelDna.Integration;
using ExcelDna.IntelliSense;

namespace XLPolyfill;

public sealed class AddIn : IExcelAddIn
{
    public void AutoOpen() => IntelliSenseServer.Install();

    public void AutoClose() => IntelliSenseServer.Uninstall();
}
