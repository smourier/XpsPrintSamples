using System;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using DirectN.Extensions.Com;
using WinRT;

namespace CustomPrintDocument.Utilities;

public static class WinRTExtensions
{
    // this is to replace WinRT's As<T> on C#/WinRT object which doesn't work well under AOT once published in release...
    // throws "Target type is not a projected type: DirectN.ICompositorInterop" from WinRT.TypeExtensions.GetHelperType(Type)
    [return: NotNullIfNotNull(nameof(winRTObject))]
    public static IComObject<T>? AsComObject<T>(this object? winRTObject, CreateObjectFlags flags = CreateObjectFlags.UniqueInstance)
    {
        if (winRTObject == null)
            return null;

        var ptr = MarshalInspectable<object>.FromManaged(winRTObject);
        var obj = ComObject.FromPointer<T>(ptr, flags);
        return obj ?? throw new InvalidCastException($"Object of type '{winRTObject.GetType().FullName}' is not of type '{typeof(T).FullName}'.");
    }
}
