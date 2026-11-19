using System.Reflection;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Registration;
using XLPolyfill.ExcelSmoke;
using Xunit;

namespace XLPolyfill.Tests;

public class ExcelDnaLoaderTests
{
    [Theory]
    [InlineData(typeof(FilterFunctions), nameof(FilterFunctions.Filter), "FILTER")]
    [InlineData(typeof(LookupFunctions), nameof(LookupFunctions.XLookup), "XLOOKUP")]
    public void WorksheetFunctionsRegisterForSerialRecalculationAndAcceptReferences(Type type, string method, string name)
    {
        var worksheetMethod = type.GetMethod(method);
        Assert.NotNull(worksheetMethod);
        var registration = new ExcelFunctionRegistration(worksheetMethod);

        Assert.Equal(name, registration.FunctionAttribute.Name);
        Assert.False(registration.FunctionAttribute.IsThreadSafe);
        Assert.All(worksheetMethod.GetParameters(), parameter =>
        {
            var argument = parameter.GetCustomAttribute<ExcelArgumentAttribute>();
            Assert.NotNull(argument);
            Assert.True(argument.AllowReference, $"{name}.{parameter.Name} must accept worksheet references.");
        });
    }

    [Theory]
    [InlineData("XLPolyfill-AddIn-packed.xll", false)]
    [InlineData("XLPolyfill-AddIn-packed.xll", true)]
    [InlineData("XLPolyfill-AddIn64-packed.xll", false)]
    [InlineData("XLPolyfill-AddIn64-packed.xll", true)]
    public void PackedLoaderRequestsAndPerformsCleanupOnLargeReturnedArrays(string packedAddIn, bool numericArray)
    {
        const int arrayType = 0x40;
        const int numberType = 0x01;
        const int stringType = 0x02;
        const int cleanupFlag = 0x4000;
        const int cellCount = 4096;
        var loader = Assembly.Load(ReadPackedLoader(packedAddIn));
        Type nativeCell = loader.GetType("ExcelDna.Loader.XlOper12", throwOnError: true)!;
        Type marshaler = loader.GetType("ExcelDna.Loader.XlDirectMarshal", throwOnError: true)!;
        Array values = numericArray ? new double[cellCount, 1] : new object[cellCount, 1];
        object cellValue = numericArray ? 42d : "selected";
        for (int row = 0; row < cellCount; row++)
            values.SetValue(cellValue, row, 0);

        string marshalMethod = numericArray ? "DoubleArray2ToXloperReturn" : "ObjectArray2Return";
        var addIn = LoaderProbe.GetLoaderAddIn(loader);
        const BindingFlags loaderMember = BindingFlags.Static | BindingFlags.NonPublic;
        try
        {
            object context = marshaler.GetMethod("GetMarshalContext")!.Invoke(null, null)!;
            var result = (IntPtr)context.GetType().GetMethod(marshalMethod)!.Invoke(context, [values])!;
            Assert.NotEqual(IntPtr.Zero, result);
            int typeOffset = Marshal.OffsetOf(nativeCell, "xlType").ToInt32();
            int lastCellOffset = Marshal.SizeOf(nativeCell) * cellCount;
            Assert.Equal(arrayType | cleanupFlag, Marshal.ReadInt32(result, typeOffset));
            Assert.Equal(numericArray ? numberType : stringType, Marshal.ReadInt32(result, lastCellOffset + typeOffset));
            Assert.True(LoaderProbe.IsBufferAllocated(loader));
            addIn.GetMethod("XlAutoFree12", loaderMember)!.Invoke(null, [result]);
            Assert.False(LoaderProbe.IsBufferAllocated(loader));
        }
        finally
        {
            marshaler.GetMethod("FreeMemory")!.Invoke(null, null);
        }
    }

    private static byte[] ReadPackedLoader(string packedAddIn)
    {
        // Use the same resource lookup and decompression code as the Excel-DNA host.
        var tasks = Assembly.LoadFrom(Path.Combine(AppContext.BaseDirectory, "ExcelDna.AddIn.Tasks.dll"));
        Type resources = tasks.GetType("ResourceHelper", throwOnError: true)!;
        Type updater = tasks.GetType("ResourceHelper+ResourceUpdater", throwOnError: true)!;
        var module = (IntPtr)resources.GetMethod("LoadXllResources")!
            .Invoke(null, [Path.Combine(AppContext.BaseDirectory, packedAddIn)])!;
        Assert.True(module != IntPtr.Zero, $"Could not open resources in {packedAddIn}.");
        try
        {
            return Assert.IsType<byte[]>(updater.GetMethod("LoadResourceBytes", BindingFlags.Static | BindingFlags.NonPublic)!
                .Invoke(null, [module, "ASSEMBLY", "EXCELDNA.LOADER"]));
        }
        finally
        {
            resources.GetMethod("FreeXllResources")!.Invoke(null, [module]);
        }
    }
}
