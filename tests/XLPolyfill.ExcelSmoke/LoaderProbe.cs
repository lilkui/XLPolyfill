using System.Reflection;
using ExcelDna.Integration;

namespace XLPolyfill.ExcelSmoke;

public static class LoaderProbe
{
    [ExcelFunction(Name = "XLPOLYFILL_SMOKE_BUFFER_ALLOCATED")]
    public static bool IsBufferAllocated(string xllPath)
    {
        var loader = AppDomain.CurrentDomain.GetAssemblies().Single(assembly =>
            assembly.GetName().Name == "ExcelDna.Loader" &&
            string.Equals((string?)GetLoaderAddIn(assembly)
                .GetProperty("PathXll", BindingFlags.Static | BindingFlags.NonPublic)!.GetValue(null),
                xllPath, StringComparison.OrdinalIgnoreCase));
        return IsBufferAllocated(loader);
    }

    internal static bool IsBufferAllocated(Assembly loader)
    {
        var marshaler = loader.GetType("ExcelDna.Loader.XlDirectMarshal", throwOnError: true)!;
        object context = marshaler.GetMethod("GetMarshalContext")!.Invoke(null, null)!;
        const BindingFlags instanceField = BindingFlags.Instance | BindingFlags.NonPublic;
        object array = context.GetType().GetField("_rank2OperArrayContext", instanceField)!.GetValue(context)!;
        return (IntPtr)array.GetType().GetField("_pNative", instanceField)!.GetValue(array)! != IntPtr.Zero ||
            (IntPtr)array.GetType().GetField("_pNativeStrings", instanceField)!.GetValue(array)! != IntPtr.Zero;
    }

    internal static Type GetLoaderAddIn(Assembly loader) => loader.GetType("ExcelDna.Loader.XlAddIn", throwOnError: true)!;
}
