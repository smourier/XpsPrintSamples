using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Windows.Graphics.Printing;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Storage.Xps.Printing;
using Windows.Win32.System.Com;
using Windows.Win32.System.WinRT;

namespace CustomPrintDocument.Model
{
    // note: we can't use DirectN AOT because there are some issue when authoring COM classes in .NET between ComWrappers, AOT and C#/WinRT
    // and we can't use GeneratedComClass either...
    public abstract class BasePrintDocument : IPrintDocumentSource, IInspectable, IPrintDocumentPageSource, IPrintPreviewPageCollection
    {
        public event EventHandler<PackageStatusUpdatedEventArgs> PackageStatusUpdated;
        private IPrintDocumentPackageTarget _docPackageTarget;

        protected BasePrintDocument(string filePath)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            FilePath = filePath;
        }

        public string FilePath { get; }
        public virtual uint? TotalPages { get; protected set; } // when/if known

        public virtual void Cancel() => _docPackageTarget?.Cancel();

        unsafe HRESULT IInspectable.GetRuntimeClassName(HSTRING* className) => HRESULT.E_NOTIMPL; // avoid throwing...
        unsafe void IInspectable.GetTrustLevel(TrustLevel* trustLevel) => throw new NotImplementedException();
        unsafe void IInspectable.GetIids(out uint iidCount, Guid** iids) { iidCount = 0; throw new NotImplementedException(); }
        void IPrintDocumentPageSource.MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget) => MakeDocument(printTaskOptions, docPackageTarget);
        IPrintPreviewPageCollection IPrintDocumentPageSource.GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget) => GetPreviewPageCollection(docPackageTarget);
        void IPrintPreviewPageCollection.MakePage(int desiredJobPage, float width, float height) => MakePage(desiredJobPage, width, height);
        void IPrintPreviewPageCollection.Paginate(int currentJobPage, nint printTaskOptions) => Paginate(currentJobPage, printTaskOptions);

        // overridable methods
        protected virtual IPrintPreviewPageCollection GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget)
        {
            return this;
        }

        protected virtual void MakePage(int desiredJobPage, float width, float height)
        {
        }

        protected virtual void Paginate(int currentJobPage, nint printTaskOptions)
        {
        }

        protected abstract Task MakeDocumentAsync(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget);
        protected virtual async void MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            ArgumentNullException.ThrowIfNull(docPackageTarget);
            _docPackageTarget = docPackageTarget;

            IConnectionPoint connectionPoint = null;
            uint cookie = 0;
            if (docPackageTarget is IConnectionPointContainer container)
            {
                container.FindConnectionPoint(typeof(IPrintDocumentPackageStatusEvent).GUID, out connectionPoint);
                if (connectionPoint != null)
                {
                    var sink = new StatusSink(this);
                    connectionPoint.Advise(sink, out cookie);
                }
            }

            try
            {
                await MakeDocumentAsync(printTaskOptions, docPackageTarget);
            }
            catch (COMException exception)
            {
                // eat cancelled error
                if (exception.ErrorCode != PInvoke.HRESULT_FROM_WIN32(WIN32_ERROR.ERROR_PRINT_CANCELLED))
                    throw;
            }
            finally
            {
                _docPackageTarget = null;
                if (connectionPoint != null && cookie != 0)
                {
                    connectionPoint.Unadvise(cookie);
                }
            }
        }

        // status update
        protected virtual void OnPackageStatusUpdated(object sender, PackageStatusUpdatedEventArgs e) => PackageStatusUpdated?.Invoke(sender, e);

        private class StatusSink(BasePrintDocument document) : IPrintDocumentPackageStatusEvent
        {
            unsafe void IPrintDocumentPackageStatusEvent.PackageStatusUpdated(PrintDocumentPackageStatus* packageStatus)
            {
                document.OnPackageStatusUpdated(document, new PackageStatusUpdatedEventArgs(*packageStatus));
            }
        }

        // tooling
        internal unsafe static PCWSTR ToPCWSTR(string text)
        {
            fixed (char* p = text)
                return new PCWSTR(p);
        }
    }

#pragma warning disable SYSLIB1096 // Convert to 'GeneratedComInterface'
    // these are not in CS/Win32 ...
    [ComImport, Guid("a96bb1db-172e-4667-82b5-ad97a252318f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPrintDocumentPageSource
    {
        IPrintPreviewPageCollection GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget);
        void MakeDocument(IntPtr printTaskOptions, IPrintDocumentPackageTarget docPackageTarget);
    };

    [ComImport, Guid("0b31cc62-d7ec-4747-9d6e-f2537d870f2b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPrintPreviewPageCollection
    {
        void Paginate(int currentJobPage, IntPtr printTaskOptions);
        void MakePage(int desiredJobPage, float width, float height);
    }
#pragma warning restore SYSLIB1096 // Convert to 'GeneratedComInterface'
}
