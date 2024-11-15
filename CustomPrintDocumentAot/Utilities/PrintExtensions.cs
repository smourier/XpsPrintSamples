using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using DirectN.Extensions.Utilities;
using WinRT;

namespace CustomPrintDocument.Utilities;

internal static class PrintExtensions
{
    public static void Log(string message, [CallerMemberName] string? methodName = null) => EventProvider.Default?.WriteMessageEvent(Environment.CurrentManagedThreadId + ":" + methodName + ":" + message);

    public static nint GetRefAndAdd(this IWinRTObject? obj, bool throwIfNull = true)
    {
        var no = obj?.NativeObject;
        if (throwIfNull && no == null)
            ArgumentNullException.ThrowIfNull(obj);

        return no?.GetRef() ?? 0;
    }

    public static void WithRef(this IWinRTObject? obj, Action<nint> action, bool throwIfNull = true)
    {
        ArgumentNullException.ThrowIfNull(action);
        var unk = GetRefAndAdd(obj, throwIfNull);
        try
        {
            action(unk);
        }
        finally
        {
            if (unk != 0)
            {
                Marshal.Release(unk);
            }
        }
    }

    public static T WithRef<T>(this IWinRTObject? obj, Func<nint, T> action, bool throwIfNull = true)
    {
        ArgumentNullException.ThrowIfNull(action);
        var unk = GetRefAndAdd(obj, throwIfNull);
        try
        {
            return action(unk);
        }
        finally
        {
            if (unk != 0)
            {
                Marshal.Release(unk);
            }
        }
    }
}
