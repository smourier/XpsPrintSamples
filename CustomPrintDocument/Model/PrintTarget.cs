using System;
using System.Threading;
using CustomPrintDocument.Utilities;
using Windows.Win32.Graphics.Dxgi;
using Windows.Win32.Graphics.Printing;

namespace CustomPrintDocument.Model
{
    public abstract class PrintTarget : IDisposable
    {
        public const uint JOB_PAGE_APPLICATION_DEFINED = 0xFFFFFFFF;

        private UnknownObject<IPrintPreviewDxgiPackageTarget> _previewTarget;

        protected PrintTarget(IPrintPreviewDxgiPackageTarget target)
        {
            ArgumentNullException.ThrowIfNull(target);
            _previewTarget = new UnknownObject<IPrintPreviewDxgiPackageTarget>(target);
        }

        protected UnknownObject<IPrintPreviewDxgiPackageTarget> PreviewTarget => _previewTarget;

        public virtual void InvalidatePreview() => PreviewTarget?.Object?.InvalidatePreview();
        public virtual void SetJobPageCount(PageCountType countType, uint count) => PreviewTarget?.Object?.SetJobPageCount(countType, count);
        public virtual void DrawPreviewPage(uint jobPageNumber, IDXGISurface pageImage, float dpiX, float dpiY) => PreviewTarget?.Object?.DrawPage(jobPageNumber, pageImage, dpiX, dpiY);

        protected abstract internal void PreviewPaginate(uint currentJobPage, nint printTaskOptions);
        protected abstract internal void MakePreviewPage(uint desiredJobPage, float width, float height);

        protected virtual void Dispose(bool disposing)
        {
            var target = Interlocked.Exchange(ref _previewTarget, null);
            if (target != null)
            {
                Extensions.Dispose(ref _previewTarget);
                target.Dispose();
            }
        }

        ~PrintTarget() { Dispose(disposing: false); }
        public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    }
}
