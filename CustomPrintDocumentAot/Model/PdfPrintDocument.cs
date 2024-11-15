using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices.Marshalling;
using System.Threading;
using System.Threading.Tasks;
using CustomPrintDocument.Utilities;
using DirectN;
using DirectN.Extensions;
using DirectN.Extensions.Com;
using DirectN.Extensions.Utilities;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Graphics.Printing;
using Windows.Storage;
using WinRT;

namespace CustomPrintDocument.Model;

[GeneratedComClass]
public sealed partial class PdfPrintDocument(string filePath) : BasePrintDocument(filePath)
{
    private PdfPrintingMode _printingMode = PdfPrintingMode.Direct2D;

    public PdfPrintingMode PrintingMode { get => _printingMode; set { if (_printingMode == value) return; _printingMode = value; PrintTarget?.InvalidatePreview(); } }
    public bool? Direct2DIgnoreHighContrast { get; set; }
    public float? Direct2DRasterDpi { get; set; }
    public D2D1_COLOR_SPACE? Direct2DColorSpace { get; set; }
    public D2D1_PRINT_FONT_SUBSET_MODE? Direct2DFontSubset { get; set; }
    private new PdfPrintTarget? PrintTarget => (PdfPrintTarget?)base.PrintTarget;

    protected override PrintTarget GetPrintTarget(IComObject<IPrintPreviewDxgiPackageTarget> target) => new PdfPrintTarget(this, target);

    protected override void MakeDocumentCore(nint printTaskOptions, IComObject<IPrintDocumentPackageTarget> docPackageTarget)
    {
        ArgumentNullException.ThrowIfNull(docPackageTarget);
        // can use options for various customizations
        // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);
        if (PrintingMode == PdfPrintingMode.Xps)
        {
            MakeDocumentUsingXpsAsync(printTaskOptions, docPackageTarget).Wait();
        }

        MakeDocumentUsingDirect2D(printTaskOptions, docPackageTarget);
    }

    private void MakeDocumentUsingDirect2D(nint printTaskOptions, IComObject<IPrintDocumentPackageTarget> docPackageTarget)
    {
        var pt = PrintTarget;
        if (pt == null)
            return;

        var pdfDoc = pt._pdfDocument;
        if (pdfDoc == null)
            return;

        var props = new D2D1_PRINT_CONTROL_PROPERTIES();
        if (Direct2DRasterDpi.HasValue)
        {
            props.rasterDPI = Direct2DRasterDpi.Value;
        }

        if (Direct2DFontSubset.HasValue)
        {
            props.fontSubset = Direct2DFontSubset.Value;
        }

        if (Direct2DColorSpace.HasValue)
        {
            props.colorSpace = Direct2DColorSpace.Value;
        }

        var renderParams = new PDF_RENDER_PARAMS
        {
            BackgroundColor = new D3DCOLORVALUE { a = 1, r = 1, g = 1, b = 1 }
        };

        if (Direct2DIgnoreHighContrast.HasValue)
        {
            renderParams.IgnoreHighContrast = new BOOLEAN((byte)(Direct2DIgnoreHighContrast.Value ? 1 : 0));
        }

        var flags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT; // D2D
#if DEBUG
        flags |= D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG;
#endif

        using var d3D11Device = D3D11Functions.D3D11CreateDevice(null, D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE, flags);
        using var imagingFactory = WicImagingFactory.Create()!;
        using var d2D1Factory = D2D1Functions.D2D1CreateFactory1();
        var dxgiDevice = d3D11Device.As<IDXGIDevice>()!;
        using var d2D1Device = d2D1Factory.CreateDevice(dxgiDevice);
        using var printControl = d2D1Device.CreatePrintControl(imagingFactory, docPackageTarget, props);
        using var dc = d2D1Device.CreateDeviceContext();

        unsafe
        {
            for (uint i = 0; i < pdfDoc.PageCount; i++)
            {
                using var commandList = dc.CreateCommandList();
                dc.SetTarget(commandList);

                var page = pdfDoc.GetPage(i);
                page.WithRef(pageUnk =>
                {
                    dc.BeginDraw();

                    pt._renderer.Object.RenderPageToDeviceContext(pageUnk, dc.Object, (nint)Unsafe.AsPointer(ref renderParams));
                    dc.EndDraw();
                });

                commandList.Object.Close();
                printControl.Object.AddPage(commandList.Object, new D2D_SIZE_F { width = (float)page.Size.Width, height = (float)page.Size.Height }, null, 0, 0);
            }
        }

        printControl.Object.Close();
    }

#pragma warning disable IDE0060 // Remove unused parameter
    private async Task MakeDocumentUsingXpsAsync(nint printTaskOptions, IComObject<IPrintDocumentPackageTarget> docPackageTarget)
#pragma warning restore IDE0060 // Remove unused parameter
    {
        var pdfDoc = PrintTarget?._pdfDocument;
        if (pdfDoc == null)
            return;

        using var xpsTarget = docPackageTarget.GetPackageTarget(Constants.ID_DOCUMENTPACKAGETARGET_MSXPS, typeof(IXpsDocumentPackageTarget).GUID);
        using var factory = xpsTarget.GetXpsOMFactory<IXpsOMObjectFactory>();

        // build a writer
        using var seqName = factory.CreatePartUri("/seq");
        using var discardName = factory.CreatePartUri("/discard");
        using var writer = xpsTarget.GetXpsOMPackageWriter(seqName, discardName);

        // start
        using var name = factory.CreatePartUri("/" + Path.GetFileNameWithoutExtension(FilePath));
        writer.Object.StartNewDocument(name.Object, null!, null!, null!, null!);

        var streams = new List<Stream>();
        // browse all PDF pages
        for (uint i = 0; i < pdfDoc.PageCount; i++)
        {
            // render page to stream
            var pdfPage = pdfDoc.GetPage(i);
            using var stream = new MemoryStream();
            await pdfPage.RenderToStreamAsync(stream.AsRandomAccessStream(), new PdfPageRenderOptions { BitmapEncoderId = BitmapEncoder.PngEncoderId });
            var size = new XPS_SIZE { width = (float)pdfPage.Size.Width, height = (float)pdfPage.Size.Height };

            // create image from stream
            stream.Position = 0;

            // note we don't dispose streams here (close would fail)
            var ustream = new DirectN.Extensions.Utilities.UnmanagedMemoryStream(stream);
            streams.Add(stream);
            using var imageUri = factory.CreatePartUri("/image" + i);
            using var image = factory.CreateImageResource(ustream, XPS_IMAGE_TYPE.XPS_IMAGE_TYPE_PNG, imageUri);

            // create a brush from image
            var viewBox = new XPS_RECT { width = size.width, height = size.height };
            using var imageBrush = factory.CreateImageBrush(image, viewBox, viewBox);

            // create a rect figure
            using var rectFigure = factory.CreateGeometryFigure(new XPS_POINT());
            rectFigure.Object.SetIsClosed(true);
            rectFigure.Object.SetIsFilled(true);
            var segmentTypes = new XPS_SEGMENT_TYPE[] { XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE };
            var segmentData = new float[] { 0, size.height, size.width, size.height, size.width, 0 };
            var segmentStrokes = new bool[] { true, true, true };

            rectFigure.Object.SetSegments(segmentTypes.Length(), segmentData.Length(), segmentTypes, segmentData, segmentStrokes);

            // create a rect geometry from figure
            using var rectGeo = factory.CreateGeometry();
            using var figures = rectGeo.GetFigures();
            figures.Object.Append(rectFigure.Object);

            // create a path and set rect geometry
            using var rectPath = factory.CreatePath();
            rectPath.Object.SetGeometryLocal(rectGeo.Object);

            // set image/brush as brush for rect
            rectPath.Object.SetFillBrushLocal(imageBrush.Object);

            // create a page & add add rect to page
            using var pageUri = factory.CreatePartUri("/page" + i);
            using var page = factory.CreatePage(size, "en", pageUri); // note: language is NOT optional
            using var visuals = page.GetVisuals();
            visuals.Object.Append(rectPath.Object);

            writer.Object.AddPage(page.Object, size, null!, null!, null!, null!);
        }

        writer.Object.Close();
        streams.Dispose();
    }

    private sealed partial class PdfPrintTarget : PrintTarget
    {
        private readonly IComObject<ID3D11Device>? _d3D11Device;
        private IComObject<IDXGISurface>? _surface;
        internal IComObject<IPdfRendererNative> _renderer;
        internal PdfDocument? _pdfDocument;
        private float? _surfaceWidth;
        private float? _surfaceHeight;
        private readonly PdfPrintDocument _printDocument;
        private readonly object _workingLock = new();
        private int? _pageBeingWorkedOn;
        private bool _stopped;

        public PdfPrintTarget(PdfPrintDocument printDocument, IComObject<IPrintPreviewDxgiPackageTarget> target)
            : base(target)
        {
            _printDocument = printDocument;

            var flags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT; // D2D
#if DEBUG
            flags |= D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG;
#endif

            _d3D11Device = D3D11Functions.D3D11CreateDevice(null, D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE, flags);
            var dxgiDevice = _d3D11Device.As<IDXGIDevice>()!;
            Functions.PdfCreateRenderer(dxgiDevice.Object, out var renderer).ThrowOnError();
            _renderer = new ComObject<IPdfRendererNative>(renderer);
        }

        private async Task EnsureDocumentAsync()
        {
            if (_pdfDocument != null)
                return;

            var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(_printDocument.FilePath));
            _pdfDocument = await PdfDocument.LoadFromFileAsync(file);
            _printDocument.TotalPages = _pdfDocument.PageCount;
            SetJobPageCount(PageCountType.FinalPageCount, _printDocument.TotalPages.Value);
        }

        private void EnsureSurface(float width, float height)
        {
            if (_d3D11Device == null || _surface == null)
                return;

            if (width == _surfaceWidth && height == _surfaceHeight)
                return;

            EnsureDocumentAsync().Wait();
            _surface.Dispose();

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
            _surface = _d3D11Device.CreateTexture2D(texture).As<IDXGISurface>();
            _surfaceWidth = width;
            _surfaceHeight = height;
        }

        protected internal override void MakePreviewPage(int desiredJobPage, float width, float height)
        {
            if (_pdfDocument == null)
                return;

            PrintExtensions.Log("MakePreviewPage " + desiredJobPage + " " + width + " x " + height);
            if (desiredJobPage == unchecked((int)Constants.JOB_PAGE_APPLICATION_DEFINED))
            {
                desiredJobPage = 1;
            }

            EnsureSurface(width, height);
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
                surface = _surface?.Object;
                target = ComObject?.Object;
            }
            if (surface == null || target == null)
                return;

            try
            {
                var page = _pdfDocument.GetPage((uint)(desiredJobPage - 1));
                page.WithRef(pageUnk =>
                {
                    unsafe
                    {
                        var renderParams = new PDF_RENDER_PARAMS
                        {
                            DestinationWidth = (uint)width,
                            DestinationHeight = (uint)height,
                            BackgroundColor = new D3DCOLORVALUE { a = 1, r = 1, g = 1, b = 1 }
                        };
                        _renderer.Object.RenderPageToSurface(pageUnk, surface, new POINT(), (nint)Unsafe.AsPointer(ref renderParams));
                    }
                });

                target.DrawPage((uint)desiredJobPage, surface, 96, 96);
            }
            finally
            {
                _pageBeingWorkedOn = null;
            }
        }

        protected internal override void PreviewPaginate(int currentJobPage, nint printTaskOptions)
        {
            var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);
            var desc = options.GetPageDescription((uint)currentJobPage);
            PrintExtensions.Log("PreviewPaginate " + currentJobPage + " ImageableRect: " + desc.ImageableRect.ToString() + " PageSize:" + desc.PageSize + " DPI:" + desc.DpiX);
            // do nothing
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _stopped = true;
                _surface?.Dispose();
                _renderer.Dispose();
                _d3D11Device?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
