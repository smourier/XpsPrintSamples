using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace CustomPrintDocument.Utilities
{
    public sealed class UnknownObject<T> : IDisposable
    {
        private object _instance;

        public UnknownObject(T instance)
        {
            _instance = instance;
        }

        public UnknownObject(object instance)
        {
            _instance = (T)instance;
        }

        public T Object => (T)_instance; // when cast fail if _instance is not null, it's probably a threading issue
        public bool IsDisposed => _instance == null;
        public void Dispose()
        {
            var instance = Interlocked.Exchange(ref _instance, null);
            if (instance != null)
            {
                Marshal.ReleaseComObject(instance);
            }
        }
    }
}
