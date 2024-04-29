﻿using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace CustomPrintDocument.Utilities
{
    // we don't use OutputDebugString because it's 100% crap, truncating, slow, etc.
    // use WpfTraceSpy https://github.com/smourier/TraceSpy to see these traces (configure an ETW Provider with guid set to 964d4572-adb9-4f3a-8170-fcbecec27467)
    public sealed partial class EventProvider : IDisposable
    {
        public static readonly EventProvider Default = new(new Guid("964d4572-adb9-4f3a-8170-fcbecec27467"));

        private long _handle;
        public Guid Id { get; }

        public EventProvider(Guid id)
        {
            Id = id;
            var hr = EventRegister(id, IntPtr.Zero, IntPtr.Zero, out _handle);
            if (hr != 0)
                throw new Win32Exception(hr);
        }

        public void Log(string message, [CallerMemberName] string methodName = null) => WriteMessageEvent(Environment.CurrentManagedThreadId + ":" + methodName + ":" + message);
        public bool WriteMessageEvent(string text, byte level = 0, long keywords = 0) => EventWriteString(_handle, level, keywords, text) == 0;

        public void Dispose()
        {
            var handle = Interlocked.Exchange(ref _handle, 0);
            if (handle != 0)
            {
                _ = EventUnregister(handle);
            }
        }

        [LibraryImport("advapi32")]
        private static partial int EventRegister(in Guid ProviderId, IntPtr EnableCallback, IntPtr CallbackContext, out long RegHandle);

        [LibraryImport("advapi32")]
        private static partial int EventUnregister(long RegHandle);

        [LibraryImport("advapi32", StringMarshalling = StringMarshalling.Utf16)]
        private static partial int EventWriteString(long RegHandle, byte Level, long Keyword, string String);
    }
}
