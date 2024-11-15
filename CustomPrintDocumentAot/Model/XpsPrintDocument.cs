using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.Marshalling;
using System.Threading;
using CustomPrintDocument.Utilities;
using DirectN;
using DirectN.Extensions;
using DirectN.Extensions.Com;
using Windows.Graphics.Printing;
using WinRT;

namespace CustomPrintDocument.Model;

[GeneratedComClass]
public sealed partial class XpsPrintDocument(string filePath) : BasePrintDocument(filePath)
{
    private XPSRAS_RENDERING_MODE _previewTextRenderingMode = XPSRAS_RENDERING_MODE.XPSRAS_RENDERING_MODE_ANTIALIASED;
    private XPSRAS_RENDERING_MODE _previewNonTextRenderingMode = XPSRAS_RENDERING_MODE.XPSRAS_RENDERING_MODE_ANTIALIASED;
    private int? _rasterizerMinimalLineWidth;

    public XPSRAS_RENDERING_MODE PreviewTextRenderingMode { get => _previewTextRenderingMode; set { if (_previewTextRenderingMode == value) return; _previewTextRenderingMode = value; PrintTarget?.InvalidatePreview(); } }
    public XPSRAS_RENDERING_MODE PreviewNonTextRenderingMode { get => _previewNonTextRenderingMode; set { if (_previewNonTextRenderingMode == value) return; _previewNonTextRenderingMode = value; PrintTarget?.InvalidatePreview(); } }
    public int? RasterizerMinimalLineWidth { get => _rasterizerMinimalLineWidth; set { if (_rasterizerMinimalLineWidth == value) return; _rasterizerMinimalLineWidth = value; PrintTarget?.InvalidatePreview(); } }

    protected override PrintTarget GetPrintTarget(IComObject<IPrintPreviewDxgiPackageTarget> target) => new XpsPrintTarget(this, target);

    unsafe protected override void MakeDocumentCore(nint printTaskOptions, IComObject<IPrintDocumentPackageTarget> docPackageTarget)
    {
        ArgumentNullException.ThrowIfNull(docPackageTarget);
        // can use options for various customizations
        // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);

        using var target = docPackageTarget.GetPackageTarget(Constants.ID_DOCUMENTPACKAGETARGET_MSXPS, typeof(IXpsDocumentPackageTarget).GUID);
        using var factory = target.GetXpsOMFactory<IXpsOMObjectFactory1>();

        // reload package (we may not run on same thead as preview)
        using var package = factory.CreatePackageFromFile1(FilePath, true);
        using var sequence = package.GetDocumentSequence();
        using var documents = sequence.GetDocuments();
        var count = documents.GetCount();
        if (count == 0)
            return;

        // get a package writer
        using var seqName = factory.CreatePartUri("/seq");
        using var discardName = factory.CreatePartUri("/discard");
        using var writer = target.GetXpsOMPackageWriter(seqName, discardName);

        // start a printer document
        using var name = factory.CreatePartUri("/name");
        writer.Object.StartNewDocument(name.Object, null!, null!, null!, null!);

        using var doc = documents.GetAt(0);
        using var pageRefs = doc.GetPageReferences();

        // add all pages from first document
        var pagesCount = pageRefs.GetCount();
        TotalPages += pagesCount;
        for (uint j = 0; j < pagesCount; j++)
        {
            using var pageRef = pageRefs.GetAt(j);
            using var page = pageRef.GetPage();
            writer.Object.AddPage(page.Object, Unsafe.NullRef<XPS_SIZE>(), null!, null!, null!, null!);
        }
        writer.Object.Close();
    }

    private sealed partial class XpsPrintTarget : PrintTarget
    {
        private readonly IComObject<ID3D11Device>? _d3D11Device;
        private readonly IComObject<ID2D1Device>? _d2D1Device;
        private readonly IComObject<IXpsOMDocument>? _document;
        private IComObject<IDXGISurface>? _surface;
        private ComObject<ID2D1DeviceContext>? _deviceContext;
        private IComObject<ID2D1Bitmap>? _pageBitmap;
        private int? _pageNumber;
        private float? _surfaceWidth;
        private float? _surfaceHeight;
        private readonly XpsPrintDocument _printDocument;
        private readonly object _workingLock = new();
        private int? _pageBeingWorkedOn;
        private bool _stopped;

        public XpsPrintTarget(XpsPrintDocument printDocument, IComObject<IPrintPreviewDxgiPackageTarget> target)
            : base(target)
        {
            _printDocument = printDocument;

            // load the file as an XPS package (could use a stream too)
            using var factory = ComObject<IXpsOMObjectFactory1>.CoCreate(Constants.XpsOMObjectFactory)!;
            using var package = factory.CreatePackageFromFile1(printDocument.FilePath, true);
            printDocument.TotalPages = 0;

            using var sequence = package.GetDocumentSequence();
            using var documents = sequence.GetDocuments();
            var count = documents.GetCount();
            if (count < 0)
                return;

            // consider the first document only
            _document = documents.GetAt(0);
            using var pagesRef = _document.GetPageReferences();

            // add all pages
            var pagesCount = pagesRef.GetCount();
            printDocument.TotalPages += pagesCount;

            var flags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT; // D2D
#if DEBUG
            flags |= D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG;
#endif
            _d3D11Device = D3D11Functions.D3D11CreateDevice(null, D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE, flags);

            var options = new D2D1_FACTORY_OPTIONS();
#if DEBUG
            options.debugLevel = D2D1_DEBUG_LEVEL.D2D1_DEBUG_LEVEL_WARNING;
#endif
            using var d2D1Factory = D2D1Functions.D2D1CreateFactory1(D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_MULTI_THREADED, options);
            var dxgiDevice = _d3D11Device.As<IDXGIDevice>()!;
            _d2D1Device = d2D1Factory.CreateDevice(dxgiDevice)!;
        }

        private unsafe void EnsureSurface(float width, float height)
        {
            if (_d2D1Device == null || _d3D11Device == null || _surface == null)
                return;

            if (width == _surfaceWidth && height == _surfaceHeight)
                return;

            _deviceContext?.Dispose();
            _surface.Dispose();

            // create a DC with a surface/texture as target
            _d2D1Device.Object.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS.D2D1_DEVICE_CONTEXT_OPTIONS_NONE, out var deviceContext);
            _deviceContext = new ComObject<ID2D1DeviceContext>(deviceContext);

            var texture = new D3D11_TEXTURE2D_DESC
            {
                ArraySize = 1,
                BindFlags = (uint)(D3D11_BIND_FLAG.D3D11_BIND_SHADER_RESOURCE | D3D11_BIND_FLAG.D3D11_BIND_RENDER_TARGET),
                Format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
                MipLevels = 1,
                SampleDesc = new DXGI_SAMPLE_DESC { Count = 1 },
                Width = (uint)width,
                Height = (uint)height,
            };
            _surface = _d3D11Device.CreateTexture2D(texture).As<IDXGISurface>()!;

            var props = new D2D1_BITMAP_PROPERTIES1
            {
                pixelFormat = new D2D1_PIXEL_FORMAT { format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM, alphaMode = D2D1_ALPHA_MODE.D2D1_ALPHA_MODE_PREMULTIPLIED },
                bitmapOptions = D2D1_BITMAP_OPTIONS.D2D1_BITMAP_OPTIONS_TARGET,
            };

            _deviceContext.Object.CreateBitmapFromDxgiSurface(_surface.Object, (nint)Unsafe.AsPointer(ref props), out var targetBitmapObj);
            using var targetBitmap = new ComObject<ID2D1Bitmap1>(targetBitmapObj);
            _deviceContext.Object.SetTarget(targetBitmap.Object);

            _surfaceWidth = width;
            _surfaceHeight = height;
        }

        private void EnsurePageBitmap(int pageNumber)
        {
            if (_deviceContext == null)
                return;

            if (pageNumber == _pageNumber)
                return;

            _pageBitmap?.Dispose();
            var document = _document?.Object;
            if (document == null)
                return;

            document.GetPageReferences(out var objPagesRef).ThrowOnError();
            using var pagesRef = new ComObject<IXpsOMPageReferenceCollection>(objPagesRef);
            pagesRef.Object.GetCount(out var pagesCount).ThrowOnError();
            if (pageNumber > pagesCount)
                return;

            pagesRef.Object.GetAt((uint)(pageNumber - 1), out var pageRef).ThrowOnError();
            pageRef.GetPage(out var objPage).ThrowOnError();
            using var page = new ComObject<IXpsOMPage>(objPage);

            // build rasterizer for XPS
            using var rasterizationFactory = ComObject<IXpsRasterizationFactory>.CoCreate(Constants.CLSID_XPSRASTERIZER_FACTORY)!;
            rasterizationFactory.Object.CreateRasterizer(page.Object, 96, _printDocument.PreviewNonTextRenderingMode, _printDocument.PreviewTextRenderingMode, out var objRasterizer).ThrowOnError();
            using var rasterizer = new ComObject<IXpsRasterizer>(objRasterizer);
            if (_printDocument.RasterizerMinimalLineWidth.HasValue)
            {
                rasterizer.Object.SetMinimalLineWidth(_printDocument.RasterizerMinimalLineWidth.Value).ThrowOnError();
            }

            page.Object.GetPageDimensions(out var size).ThrowOnError();
            rasterizer.Object.RasterizeRect(0, 0, (int)size.width, (int)size.height, null, out var wicBmp).ThrowOnError();

            using var wicBitmap = new ComObject<IWICBitmap>(wicBmp);
            _pageBitmap = _deviceContext.CreateBitmapFromWicBitmap(wicBitmap);
            PrintExtensions.Log("loaded bitmap size: " + size.width + " x " + size.height);

            _pageNumber = pageNumber;
        }

        protected internal override void MakePreviewPage(int desiredJobPage, float width, float height)
        {
            PrintExtensions.Log("MakePreviewPage " + desiredJobPage + " " + width + " x " + height);
            if (desiredJobPage == unchecked((int)Constants.JOB_PAGE_APPLICATION_DEFINED))
            {
                desiredJobPage = 1;
            }

            // this loop is not super smart but some objects (xps, etc.) here are thread-bound, so we just wait
            ID2D1DeviceContext? dc;
            ID2D1Bitmap? bitmap;
            IDXGISurface? surface;
            IPrintPreviewDxgiPackageTarget? target;
            lock (_workingLock)
            {
                if (_pageBeingWorkedOn == desiredJobPage)
                    return;

                while (!_stopped && _pageBeingWorkedOn.HasValue)
                {
                    PrintExtensions.Log("Sleep...");
                    Thread.Sleep(100);
                }

                _pageBeingWorkedOn = desiredJobPage;
                EnsurePageBitmap(desiredJobPage);
                dc = _deviceContext?.Object;
                bitmap = _pageBitmap?.Object;
                surface = _surface?.Object;
                target = ComObject?.Object;
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
                        dc.Clear(new D3DCOLORVALUE { a = 1, r = 1, g = 1, b = 1 });
                        dc.DrawBitmap(bitmap, 1, D2D1_INTERPOLATION_MODE.D2D1_INTERPOLATION_MODE_HIGH_QUALITY_CUBIC);
                    }
                    finally
                    {
                        dc.EndDraw();
                    }
                }

                target.DrawPage((uint)desiredJobPage, surface, 96, 96);
            }
            finally
            {
                _pageBeingWorkedOn = null;
            }
        }

        private static int _i;

        protected internal override void PreviewPaginate(int currentJobPage, nint printTaskOptions)
        {
            var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);
            var desc = options.GetPageDescription((uint)currentJobPage);
            PrintExtensions.Log("PreviewPaginate " + currentJobPage + " ImageableRect: " + desc.ImageableRect.ToString() + " PageSize:" + desc.PageSize + " DPI:" + desc.DpiX);
            // note we build our target surface using Dips not pixels
            // this is somewhat incorrect but the ImageableRect is generally already bigger than the preview page
            EnsureSurface((uint)desc.ImageableRect.Width, (uint)desc.ImageableRect.Height);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _stopped = true;
                _document?.Dispose();
                _pageBitmap?.Dispose();
                _surface?.Dispose();
                _deviceContext?.Dispose();
                _d2D1Device?.Dispose();
                _d3D11Device?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
