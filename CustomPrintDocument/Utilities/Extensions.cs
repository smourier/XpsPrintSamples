using System;
using System.Threading;
using Windows.Win32.Foundation;

namespace CustomPrintDocument.Utilities
{
    internal static class Extensions
    {
        public unsafe static PCWSTR ToPCWSTR(this string text)
        {
            fixed (char* p = text)
                return new PCWSTR(p);
        }

        public static void Dispose<T>(ref UnknownObject<T> unknownObject) => Interlocked.Exchange(ref unknownObject, null)?.Dispose();
        public static void Dispose(ref IDisposable disposable) { try { Interlocked.Exchange(ref disposable, null)?.Dispose(); } catch { /* do nothing */ } }
    }
}
