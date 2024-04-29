using System;
using System.Runtime.InteropServices;
using System.Threading;
using CustomPrintDocument.Utilities;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Graphics.Direct2D;
using Windows.Win32.Graphics.Direct3D;
using Windows.Win32.Graphics.Direct3D11;
using Windows.Win32.Graphics.Dxgi;
using Windows.Win32.Graphics.Dxgi.Common;
using Windows.Win32.Graphics.Printing;

namespace CustomPrintDocument.Model
{
    public abstract class PrintTarget : IDisposable
    {
        public const uint JOB_PAGE_APPLICATION_DEFINED = 0xFFFFFFFF;

        private UnknownObject<ID3D11Device> _d3D11Device;
        private UnknownObject<ID2D1Device> _d2D1Device;
        //private IWICImagingFactory _imagingFactory;
        //private ID2D1PrintControl _d2D1PrintControl;
        private UnknownObject<IPrintPreviewDxgiPackageTarget> _previewTarget;

        protected PrintTarget(IPrintPreviewDxgiPackageTarget target)
        {
            ArgumentNullException.ThrowIfNull(target);
            _previewTarget = new UnknownObject<IPrintPreviewDxgiPackageTarget>(target);
            Marshal.GetIUnknownForObject(target); // equivalent to AddRef

            var flags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT; // D2D
#if DEBUG
            flags |= D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG;
#endif
            unsafe
            {
                PInvoke.D3D11CreateDevice(
                    null,
                    D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE,
                    HMODULE.Null,
                    flags,
                    null,
                    0,
                    PInvoke.D3D11_SDK_VERSION,
                    out var d3D11Device,
                    null,
                    out var deviceContext).ThrowOnFailure();
                Marshal.ReleaseComObject(deviceContext);
                ((ID3D11Multithread)d3D11Device).SetMultithreadProtected(true);
                _d3D11Device = new UnknownObject<ID3D11Device>(d3D11Device);
            }

            var options = new D2D1_FACTORY_OPTIONS();
#if DEBUG
            options.debugLevel = D2D1_DEBUG_LEVEL.D2D1_DEBUG_LEVEL_WARNING;
#endif
            PInvoke.D2D1CreateFactory(D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_MULTI_THREADED, typeof(ID2D1Factory1).GUID, options, out var obj).ThrowOnFailure();
            var d2D1Factory = (ID2D1Factory1)obj;
            var dxgiDevice = (IDXGIDevice)_d3D11Device.Object;
            d2D1Factory.CreateDevice(dxgiDevice, out var d2D1Device);
            _d2D1Device = new UnknownObject<ID2D1Device>(d2D1Device);
            Marshal.ReleaseComObject(obj);

            //PInvoke.CoCreateInstance(PInvoke.CLSID_WICImagingFactory, null, Windows.Win32.System.Com.CLSCTX.CLSCTX_ALL, typeof(IWICImagingFactory).GUID, out obj).ThrowOnFailure();
            //_imagingFactory = (IWICImagingFactory)obj;
        }

        protected UnknownObject<IPrintPreviewDxgiPackageTarget> PreviewTarget => _previewTarget;
        protected UnknownObject<ID2D1Device> D2D1Device => _d2D1Device;

        protected virtual UnknownObject<IDXGISurface> CreateSurface(uint width, uint height)
        {
            var texture = new D3D11_TEXTURE2D_DESC
            {
                ArraySize = 1,
                BindFlags = D3D11_BIND_FLAG.D3D11_BIND_SHADER_RESOURCE | D3D11_BIND_FLAG.D3D11_BIND_RENDER_TARGET,
                Format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
                MipLevels = 1,
                SampleDesc = new DXGI_SAMPLE_DESC { Count = 1 },
                Width = width,
                Height = height,
            };
            _d3D11Device.Object.CreateTexture2D(texture, null, out var tex);
            return new UnknownObject<IDXGISurface>((IDXGISurface)tex);
        }

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
                var unk = Marshal.GetIUnknownForObject(target);
                Marshal.Release(unk);
                Marshal.Release(unk);

                //_imagingFactory.ReleaseComObject();
                //_d2D1PrintControl.ReleaseComObject();
                Extensions.Dispose(ref _previewTarget);
                Extensions.Dispose(ref _d2D1Device);
                Extensions.Dispose(ref _d3D11Device);
                target.Dispose();
            }
        }

        ~PrintTarget() { Dispose(disposing: false); }
        public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
    }
}
