using System;
using System.IO;
using System.Threading;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.System.Com;

namespace CustomPrintDocument.Model
{
    public sealed partial class UnmanagedMemoryStream : Stream, IStream
    {
        private IStream _stream;

        private UnmanagedMemoryStream(IStream stream)
        {
            _stream = stream;
        }

        unsafe public UnmanagedMemoryStream()
        {
            _stream = PInvoke.SHCreateMemStream(null, 0);
            CheckStream();
        }

        public UnmanagedMemoryStream(string filePath, STGM mode = STGM.STGM_READ)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            PInvoke.SHCreateStreamOnFile(filePath, (uint)mode, out _stream).ThrowOnFailure();
        }

        public UnmanagedMemoryStream(Stream stream, int bufferSize = 81920) // below LOH
            : this()
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentOutOfRangeException.ThrowIfLessThan(bufferSize, 1);
            stream.CopyTo(this, bufferSize);
        }

        public UnmanagedMemoryStream(byte[] bytes)
        {
            ArgumentNullException.ThrowIfNull(bytes);
            _stream = PInvoke.SHCreateMemStream(bytes);
            CheckStream();
        }

        public UnmanagedMemoryStream(nint bytes, uint length)
        {
            if (bytes == 0 && length > 0)
                throw new ArgumentOutOfRangeException(nameof(length));

            unsafe
            {
                _stream = PInvoke.SHCreateMemStream((byte*)bytes, length);
            }
            CheckStream();
        }

        unsafe public UnmanagedMemoryStream(byte* bytes, uint length)
        {
            ArgumentNullException.ThrowIfNull(bytes);
            if (bytes == null && length > 0)
                throw new ArgumentOutOfRangeException(nameof(length));

            _stream = PInvoke.SHCreateMemStream(bytes, length);
            CheckStream();
        }

        public bool CommitOnDispose { get; set; } = true;
        public bool DeepClone { get; set; } = true;

        public IStream NativeStream
        {
            get
            {
                var stream = _stream;
                ObjectDisposedException.ThrowIf(stream == null, this);
                return stream;
            }
        }

        public override long Length
        {
            get
            {
                NativeStream.Stat(out var stat, 0);
                return (long)stat.cbSize;
            }
        }

        public override int ReadTimeout => Timeout.Infinite;
        public override int WriteTimeout => Timeout.Infinite;
        public override bool CanTimeout => false;
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => true;

        public override long Position
        {
            get
            {
                unsafe
                {
                    ulong pos;
                    NativeStream.Seek(0, SeekOrigin.Current, &pos);
                    return (long)pos;
                }
            }
            set => Seek(value, SeekOrigin.Begin);
        }

        public override void Flush() => NativeStream.Commit(STGC.STGC_DEFAULT);
        public override void SetLength(long value) => NativeStream.SetSize((ulong)value);
        public override long Seek(long offset, SeekOrigin origin)
        {
            unsafe
            {
                ulong pos;
                NativeStream.Seek(offset, origin, &pos);
                return (long)pos;
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            ValidateBufferArguments(buffer, offset, count);
            unsafe
            {
                fixed (byte* p = buffer)
                {
                    uint read;
                    NativeStream.Read(p + offset, (uint)count, &read).ThrowOnFailure();
                    return (int)read;
                }
            }
        }

        public override void Write(byte[] buffer, int offset, int count)
        {
            ValidateBufferArguments(buffer, offset, count);
            unsafe
            {
                fixed (byte* p = buffer)
                {
                    NativeStream.Write(p + offset, (uint)count).ThrowOnFailure();
                }
            }
        }

        private void CheckStream()
        {
            if (_stream == null)
                throw new OutOfMemoryException();
        }

        protected override void Dispose(bool disposing)
        {
            var stream = Interlocked.Exchange(ref _stream, null);
            if (stream != null && CommitOnDispose)
            {
                stream.Commit(0);
            }
            base.Dispose(disposing);
        }

        void IStream.Clone(out IStream ppstm)
        {
            if (DeepClone)
            {
                NativeStream.Clone(out var clonedStream);
                if (clonedStream == null)
                {
                    ppstm = null!;
                    return;
                }

                ppstm = new UnmanagedMemoryStream(clonedStream);
                return;
            }

            NativeStream.Clone(out ppstm);
        }

        unsafe HRESULT IStream.Read(void* pv, uint cb, uint* pcbRead) => NativeStream.Read(pv, cb, pcbRead);
        unsafe HRESULT IStream.Write(void* pv, uint cb, uint* pcbWritten) => NativeStream.Write(pv, cb, pcbWritten);
        unsafe void IStream.Seek(long dlibMove, SeekOrigin dwOrigin, ulong* plibNewPosition) => NativeStream.Seek(dlibMove, dwOrigin, plibNewPosition);
        void IStream.SetSize(ulong libNewSize) => NativeStream.SetSize(libNewSize);
        unsafe void IStream.CopyTo(IStream pstm, ulong cb, ulong* pcbRead, ulong* pcbWritten) => NativeStream.CopyTo(pstm, cb, pcbRead, pcbWritten);
        void IStream.Commit(STGC grfCommitFlags) => NativeStream.Commit(grfCommitFlags);
        void IStream.Revert() => NativeStream.Revert();
        void IStream.LockRegion(ulong libOffset, ulong cb, LOCKTYPE dwLockType) => NativeStream.LockRegion(libOffset, cb, dwLockType);
        void IStream.UnlockRegion(ulong libOffset, ulong cb, uint dwLockType) => _stream.UnlockRegion(libOffset, cb, dwLockType);
        unsafe void IStream.Stat(STATSTG* pstatstg, STATFLAG grfStatFlag) => NativeStream.Stat(pstatstg, grfStatFlag);
        unsafe HRESULT ISequentialStream.Read(void* pv, uint cb, uint* pcbRead) => NativeStream.Read(pv, cb, pcbRead);
        unsafe HRESULT ISequentialStream.Write(void* pv, uint cb, uint* pcbWritten) => NativeStream.Write(pv, cb, pcbWritten);
    }
}
