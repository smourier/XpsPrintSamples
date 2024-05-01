using System;
using System.Runtime.InteropServices;
using System.Threading;
using CustomPrintDocument.Utilities;
using Windows.Graphics.Printing;
using Windows.Win32;
using Windows.Win32.Graphics.Direct2D;
using Windows.Win32.Graphics.Direct2D.Common;
using Windows.Win32.Graphics.Direct3D11;
using Windows.Win32.Graphics.Dxgi;
using Windows.Win32.Graphics.Dxgi.Common;
using Windows.Win32.Graphics.Printing;
using Windows.Win32.Storage.Xps;
using Windows.Win32.Storage.Xps.Printing;
using Windows.Win32.System.Com;
using WinRT;

namespace CustomPrintDocument.Model
{
    public sealed class XpsPrintDocument(string filePath) : BasePrintDocument(filePath)
    {
        private XPSRAS_RENDERING_MODE _previewTextRenderingMode = XPSRAS_RENDERING_MODE.XPSRAS_RENDERING_MODE_ANTIALIASED;
        private XPSRAS_RENDERING_MODE _previewNonTextRenderingMode = XPSRAS_RENDERING_MODE.XPSRAS_RENDERING_MODE_ANTIALIASED;
        private int? _rasterizerMinimalLineWidth;

        public XPSRAS_RENDERING_MODE PreviewTextRenderingMode { get => _previewTextRenderingMode; set { if (_previewTextRenderingMode == value) return; _previewTextRenderingMode = value; PrintTarget?.InvalidatePreview(); } }
        public XPSRAS_RENDERING_MODE PreviewNonTextRenderingMode { get => _previewNonTextRenderingMode; set { if (_previewNonTextRenderingMode == value) return; _previewNonTextRenderingMode = value; PrintTarget?.InvalidatePreview(); } }
        public int? RasterizerMinimalLineWidth { get => _rasterizerMinimalLineWidth; set { if (_rasterizerMinimalLineWidth == value) return; _rasterizerMinimalLineWidth = value; PrintTarget?.InvalidatePreview(); } }

        protected override PrintTarget GetPrintTarget(IPrintPreviewDxgiPackageTarget target) => new XpsPrintTarget(this, target);

        unsafe protected override void MakeDocumentCore(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            ArgumentNullException.ThrowIfNull(docPackageTarget);
            // can use options for various customizations
            // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);

            var IXpsDocumentPackageTargetGuid = typeof(IXpsDocumentPackageTarget).GUID;
            var xpsGuid = PInvoke.ID_DOCUMENTPACKAGETARGET_MSXPS;
            docPackageTarget.GetPackageTarget(&xpsGuid, &IXpsDocumentPackageTargetGuid, out var obj);
            var target = (IXpsDocumentPackageTarget)obj;
            var factory = (IXpsOMObjectFactory1)target.GetXpsOMFactory();

            // reload package (we may not run on same thead as preview)
            var package = factory.CreatePackageFromFile1(FilePath, true);
            var sequence = package.GetDocumentSequence();
            var documents = sequence.GetDocuments();
            var count = documents.GetCount();
            if (count == 0)
                return;

            // get a package writer
            var seqName = factory.CreatePartUri("/seq");
            var discardName = factory.CreatePartUri("/discard");
            var writer = target.GetXpsOMPackageWriter(seqName, discardName);

            // start a printer document
            var name = factory.CreatePartUri("/name");
            writer.StartNewDocument(name, null, null, null, null);

            var doc = documents.GetAt(0);
            var pagesRef = doc.GetPageReferences();

            // add all pages from first document
            var pagesCount = pagesRef.GetCount();
            TotalPages += pagesCount;
            for (uint j = 0; j < pagesCount; j++)
            {
                var pageRef = pagesRef.GetAt(j);
                var page = pageRef.GetPage();
                writer.AddPage(page, null, null, null, null, null);
            }
            writer.Close();
        }

        private sealed class XpsPrintTarget : PrintTarget
        {
            private UnknownObject<ID3D11Device> _d3D11Device;
            private UnknownObject<ID2D1Device> _d2D1Device;
            private UnknownObject<IXpsOMDocument> _document;
            private UnknownObject<ID2D1DeviceContext> _deviceContext;
            private UnknownObject<IDXGISurface> _surface;
            private UnknownObject<ID2D1Bitmap> _pageBitmap;
            private uint? _pageNumber;
            private float? _surfaceWidth;
            private float? _surfaceHeight;
            private readonly XpsPrintDocument _printDocument;
            private readonly object _workingLock = new();
            private uint? _pageBeingWorkedOn;
            private bool _stopped;

            public XpsPrintTarget(XpsPrintDocument printDocument, IPrintPreviewDxgiPackageTarget target)
                : base(target)
            {
                _printDocument = printDocument;

                // load the file as an XPS package (could use a stream too)
                var factory = (IXpsOMObjectFactory1)new XpsOMObjectFactory();
                var package = factory.CreatePackageFromFile1(printDocument.FilePath, true);
                Marshal.ReleaseComObject(factory);
                printDocument.TotalPages = 0;

                var sequence = package.GetDocumentSequence();
                var documents = sequence.GetDocuments();
                Marshal.ReleaseComObject(sequence);
                var count = documents.GetCount();
                if (count < 0)
                    return;

                // consider the first document only
                _document = new UnknownObject<IXpsOMDocument>(documents.GetAt(0));
                var pagesRef = _document.Object.GetPageReferences();

                // add all pages
                var pagesCount = pagesRef.GetCount();
                Marshal.ReleaseComObject(pagesRef);
                printDocument.TotalPages += pagesCount;

                _d3D11Device = Extensions.CreateD3D11Device();
                var dxgiDevice = (IDXGIDevice)_d3D11Device.Object;

                using var d2D1Factory = Extensions.CreateD2D1Factory();
                d2D1Factory.Object.CreateDevice(dxgiDevice, out var d2D1Device);
                _d2D1Device = new UnknownObject<ID2D1Device>(d2D1Device);
            }

            private void EnsureSurface(float width, float height)
            {
                if (width == _surfaceWidth && height == _surfaceHeight)
                    return;

                Extensions.Dispose(ref _deviceContext);
                Extensions.Dispose(ref _surface);

                // create a DC with a surface/texture as target
                _d2D1Device.Object.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS.D2D1_DEVICE_CONTEXT_OPTIONS_NONE, out var deviceContext);
                _deviceContext = new UnknownObject<ID2D1DeviceContext>(deviceContext);
                _surface = _d3D11Device.Object.CreateSurface((uint)width, (uint)height);

                var props = new D2D1_BITMAP_PROPERTIES1
                {
                    pixelFormat = new D2D1_PIXEL_FORMAT { format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, alphaMode = D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_PREMULTIPLIED },
                    bitmapOptions = D2D1_BITMAP_OPTIONS.D2D1_BITMAP_OPTIONS_TARGET,
                };
                _deviceContext.Object.CreateBitmapFromDxgiSurface(_surface.Object, props, out var targetBitmap);
                _deviceContext.Object.SetTarget(targetBitmap);
                Marshal.ReleaseComObject(targetBitmap);

                _surfaceWidth = width;
                _surfaceHeight = height;
            }

            private void EnsurePageBitmap(uint pageNumber)
            {
                if (pageNumber == _pageNumber)
                    return;

                Extensions.Dispose(ref _pageBitmap);
                var document = _document?.Object;
                if (document == null)
                    return;

                var pagesRef = document.GetPageReferences();
                var pagesCount = pagesRef.GetCount();
                if (pageNumber > pagesCount)
                    return;

                var pageRef = pagesRef.GetAt(pageNumber - 1);
                var page = pageRef.GetPage();

                // build rasterizer for XPS
                PInvoke.CoCreateInstance(PInvoke.CLSID_XPSRASTERIZER_FACTORY, null, CLSCTX.CLSCTX_ALL, typeof(IXpsRasterizationFactory).GUID, out var obj).ThrowOnFailure();
                var rasterizationFactory = (IXpsRasterizationFactory)obj;
                rasterizationFactory.CreateRasterizer(page, 96, _printDocument.PreviewNonTextRenderingMode, _printDocument.PreviewTextRenderingMode, out var rasterizer);
                Marshal.ReleaseComObject(rasterizationFactory);
                try
                {
                    if (_printDocument.RasterizerMinimalLineWidth.HasValue)
                    {
                        rasterizer.SetMinimalLineWidth(_printDocument.RasterizerMinimalLineWidth.Value);
                    }

                    var size = page.GetPageDimensions();
                    rasterizer.RasterizeRect(0, 0, (int)size.width, (int)size.height, null, out var wicBitmap);

                    unsafe
                    {
                        _deviceContext.Object.CreateBitmapFromWicBitmap(wicBitmap, null, out var bitmap);
                        _pageBitmap = new UnknownObject<ID2D1Bitmap>(bitmap);
                        Marshal.ReleaseComObject(wicBitmap);
                        EventProvider.Default.Log("loaded bitmap size: " + size.width + " x " + size.height);
                    }

                    _pageNumber = pageNumber;
                }
                finally
                {
                    Marshal.ReleaseComObject(rasterizer);
                    Marshal.ReleaseComObject(pageRef);
                    Marshal.ReleaseComObject(page);
                }
            }

            protected internal override void MakePreviewPage(uint desiredJobPage, float width, float height)
            {
                EventProvider.Default.Log("MakePreviewPage " + desiredJobPage + " " + width + " x " + height);
                if (desiredJobPage == JOB_PAGE_APPLICATION_DEFINED)
                {
                    desiredJobPage = 1;
                }

                // this loop is not super smart but some objects (xps, etc.) here are thread-bound, so we just wait
                ID2D1DeviceContext dc;
                ID2D1Bitmap bitmap;
                IDXGISurface surface;
                IPrintPreviewDxgiPackageTarget target;
                lock (_workingLock)
                {
                    if (_pageBeingWorkedOn == desiredJobPage)
                        return;

                    while (!_stopped && _pageBeingWorkedOn.HasValue)
                    {
                        EventProvider.Default.Log("Sleep...");
                        Thread.Sleep(100);
                    }

                    _pageBeingWorkedOn = desiredJobPage;
                    EnsurePageBitmap(desiredJobPage);
                    dc = _deviceContext?.Object;
                    bitmap = _pageBitmap?.Object;
                    surface = _surface?.Object;
                    target = PreviewTarget?.Object;
                }
                if (dc == null || bitmap == null || surface == null || target == null)
                    return;

                try
                {
                    unsafe
                    {
                        var i = Interlocked.Increment(ref _i);
                        dc.BeginDraw();
                        try
                        {
                            dc.Clear(new D2D1_COLOR_F { a = 1, r = 1, g = 1, b = 1 });
                            dc.DrawBitmap(bitmap, null, 1, D2D1_INTERPOLATION_MODE.D2D1_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC, null, null);
                        }
                        finally
                        {
                            dc.EndDraw().ThrowOnFailure();
                        }
                    }

                    target.DrawPage(desiredJobPage, surface, 96, 96);
                }
                finally
                {
                    _pageBeingWorkedOn = null;
                }
            }

            private static int _i;

            protected internal override void PreviewPaginate(uint currentJobPage, nint printTaskOptions)
            {
                var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);
                var desc = options.GetPageDescription(currentJobPage);
                EventProvider.Default.Log("PreviewPaginate " + currentJobPage + " ImageableRect: " + desc.ImageableRect.ToString() + " PageSize:" + desc.PageSize + " DPI:" + desc.DpiX);
                // note we build our target surface using Dips not pixels
                // this is somewhat incorrect but the ImageableRect is generally already bigger than the preview page
                EnsureSurface((uint)desc.ImageableRect.Width, (uint)desc.ImageableRect.Height);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _stopped = true;
                    Extensions.Dispose(ref _document);
                    Extensions.Dispose(ref _pageBitmap);
                    Extensions.Dispose(ref _surface);
                    Extensions.Dispose(ref _deviceContext);
                    Extensions.Dispose(ref _d2D1Device);
                    Extensions.Dispose(ref _d3D11Device);
                }
                base.Dispose(disposing);
            }
        }
    }
}
