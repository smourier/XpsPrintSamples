using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using CustomPrintDocument.Utilities;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Graphics.Printing;
using Windows.Storage;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Graphics.Direct2D;
using Windows.Win32.Graphics.Direct2D.Common;
using Windows.Win32.Graphics.Direct3D11;
using Windows.Win32.Graphics.Dxgi;
using Windows.Win32.Graphics.Printing;
using Windows.Win32.Storage.Packaging.Opc;
using Windows.Win32.Storage.Xps;
using Windows.Win32.Storage.Xps.Printing;
using Windows.Win32.System.WinRT.Pdf;
using WinRT;

namespace CustomPrintDocument.Model
{
    public sealed class PdfPrintDocument : BasePrintDocument
    {
        private PdfPrintingMode _printingMode;

        public PdfPrintDocument(string filePath)
            : base(filePath)
        {
            _printingMode = PdfPrintingMode.Direct2D;
        }

        public PdfPrintingMode PrintingMode { get => _printingMode; set { if (_printingMode == value) return; _printingMode = value; PrintTarget?.InvalidatePreview(); } }
        public bool? Direct2DIgnoreHighContrast { get; set; }
        public float? Direct2DRasterDpi { get; set; }
        public D2D1_COLOR_SPACE? Direct2DColorSpace { get; set; }
        public D2D1_PRINT_FONT_SUBSET_MODE? Direct2DFontSubset { get; set; }
        private new PdfPrintTarget PrintTarget => (PdfPrintTarget)base.PrintTarget;

        protected override PrintTarget GetPrintTarget(IPrintPreviewDxgiPackageTarget target) => new PdfPrintTarget(this, target);

        protected override void MakeDocumentCore(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            ArgumentNullException.ThrowIfNull(docPackageTarget);
            // can use options for various customizations
            // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);
            if (PrintingMode == PdfPrintingMode.Xps)
            {
                MakeDocumentUsingXpsAsync(printTaskOptions, docPackageTarget).Wait();
            }

            MakeDocumentUsingNative(printTaskOptions, docPackageTarget);
        }

        private void MakeDocumentUsingNative(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
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
                BackgroundColor = new D2D_COLOR_F { a = 1, r = 1, g = 1, b = 1 }
            };

            if (Direct2DIgnoreHighContrast.HasValue)
            {
                renderParams.IgnoreHighContrast = Direct2DIgnoreHighContrast.Value;
            }

            using var imagingFactory = Extensions.CreateWICImagingFactory();
            using var d3D11Device = Extensions.CreateD3D11Device();
            using var d2D1Factory = Extensions.CreateD2D1Factory();
            var dxgiDevice = (IDXGIDevice)d3D11Device.Object;
            d2D1Factory.Object.CreateDevice(dxgiDevice, out var d2D1Device);
            d2D1Device.CreatePrintControl(imagingFactory.Object, docPackageTarget, props, out var printControl);
            d2D1Device.CreateDeviceContext(D2D1_DEVICE_CONTEXT_OPTIONS.D2D1_DEVICE_CONTEXT_OPTIONS_NONE, out var dc);

            unsafe
            {
                for (uint i = 0; i < PrintTarget._pdfDocument.PageCount; i++)
                {
                    dc.CreateCommandList(out var commandList);
                    dc.SetTarget(commandList);

                    var page = PrintTarget._pdfDocument.GetPage(i);
                    dc.BeginDraw();
                    PrintTarget._renderer.Object.RenderPageToDeviceContext(page, dc, renderParams);
                    dc.EndDraw();
                    commandList.Close();
                    printControl.AddPage(commandList, new D2D_SIZE_F { width = (float)page.Size.Width, height = (float)page.Size.Height }, null);
                    Marshal.ReleaseComObject(commandList);
                }
            }

            printControl.Close();
            Marshal.ReleaseComObject(printControl);
            Marshal.ReleaseComObject(dc);
            Marshal.ReleaseComObject(d2D1Device);
        }

        private async Task MakeDocumentUsingXpsAsync(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            var xpsTarget = GetXpsDocumentPackageTarget(docPackageTarget);
            var factory = xpsTarget.GetXpsOMFactory();

            // build a writer
            var seqName = factory.CreatePartUri("/seq");
            var discardName = factory.CreatePartUri("/discard");
            var writer = xpsTarget.GetXpsOMPackageWriter(seqName, discardName);

            // start
            var name = factory.CreatePartUri("/" + Path.GetFileNameWithoutExtension(FilePath));
            writer.StartNewDocument(name, null!, null!, null!, null!);

            var streams = new List<Stream>();
            // browse all PDF pages
            for (uint i = 0; i < PrintTarget._pdfDocument.PageCount; i++)
            {
                // render page to stream
                var pdfPage = PrintTarget._pdfDocument.GetPage(i);
                using var stream = new MemoryStream();
                await pdfPage.RenderToStreamAsync(stream.AsRandomAccessStream(), new PdfPageRenderOptions { BitmapEncoderId = BitmapEncoder.PngEncoderId });
                var size = new XPS_SIZE { width = (float)pdfPage.Size.Width, height = (float)pdfPage.Size.Height };

                // create image from stream
                stream.Position = 0;

                // note we don't dispose streams here (close would fail)
                var ustream = new UnmanagedMemoryStream(stream);
                streams.Add(stream);
                var imageUri = factory.CreatePartUri("/image" + i);
                var image = factory.CreateImageResource(ustream, XPS_IMAGE_TYPE.XPS_IMAGE_TYPE_PNG, imageUri);

                // create a brush from image
                var viewBox = new XPS_RECT { width = size.width, height = size.height };
                var imageBrush = factory.CreateImageBrush(image, viewBox, viewBox);

                // create a rect figure
                var rectFigure = factory.CreateGeometryFigure(new XPS_POINT());
                rectFigure.SetIsClosed(true);
                rectFigure.SetIsFilled(true);
                var segmentTypes = new XPS_SEGMENT_TYPE[] { XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE };
                var segmentData = new float[] { 0, size.height, size.width, size.height, size.width, 0 };
                var segmentStrokes = new bool[] { true, true, true };

                unsafe
                {
                    // SetSegments def is wrong https://github.com/microsoft/win32metadata/issues/1889
                    fixed (float* f = segmentData)
                    fixed (XPS_SEGMENT_TYPE* st = segmentTypes)
                    fixed (bool* ss = segmentStrokes)
                        rectFigure.SetSegments((uint)segmentTypes.Length, (uint)segmentData.Length, st, *f, (BOOL*)ss);
                }

                // create a rect geometry from figure
                var rectGeo = factory.CreateGeometry();
                var figures = rectGeo.GetFigures();
                figures.Append(rectFigure);

                // create a path and set rect geometry
                var rectPath = factory.CreatePath();
                rectPath.SetGeometryLocal(rectGeo);

                // set image/brush as brush for rect
                rectPath.SetFillBrushLocal(imageBrush);

                // create a page & add add rect to page
                var pageUri = factory.CreatePartUri("/page" + i);
                var page = CreatePage(factory, size, pageUri);
                var visuals = page.GetVisuals();
                visuals.Append(rectPath);

                writer.AddPage(page, size, null!, null!, null!, null!);
            }
            writer.Close();
            foreach (var stream in streams)
            {
                stream.Dispose();
            }

            // unsafes in async
            static unsafe IXpsOMPage CreatePage(IXpsOMObjectFactory factory, XPS_SIZE size, IOpcPartUri pageUri) => factory.CreatePage(size, "en", pageUri); // note: language is NOT optional
            static unsafe IXpsDocumentPackageTarget GetXpsDocumentPackageTarget(IPrintDocumentPackageTarget docPackageTarget)
            {
                var IXpsDocumentPackageTargetGuid = typeof(IXpsDocumentPackageTarget).GUID;
                var xpsGuid = PInvoke.ID_DOCUMENTPACKAGETARGET_MSXPS;
                docPackageTarget.GetPackageTarget(&xpsGuid, &IXpsDocumentPackageTargetGuid, out var obj);
                return (IXpsDocumentPackageTarget)obj;
            }
        }

        private sealed class PdfPrintTarget : PrintTarget
        {
            private UnknownObject<ID3D11Device> _d3D11Device;
            private UnknownObject<IDXGISurface> _surface;
            internal UnknownObject<IPdfRendererNative> _renderer;
            internal PdfDocument _pdfDocument;
            private float? _surfaceWidth;
            private float? _surfaceHeight;
            private readonly PdfPrintDocument _printDocument;
            private readonly object _workingLock = new();
            private uint? _pageBeingWorkedOn;
            private bool _stopped;

            public PdfPrintTarget(PdfPrintDocument printDocument, IPrintPreviewDxgiPackageTarget target)
                : base(target)
            {
                _printDocument = printDocument;

                _d3D11Device = Extensions.CreateD3D11Device();
                var dxgiDevice = (IDXGIDevice)_d3D11Device.Object;
                PInvoke.PdfCreateRenderer(dxgiDevice, out var renderer).ThrowOnFailure();
                _renderer = new UnknownObject<IPdfRendererNative>(renderer);
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
                if (width == _surfaceWidth && height == _surfaceHeight)
                    return;

                EnsureDocumentAsync().Wait();
                Extensions.Dispose(ref _surface);

                _surface = _d3D11Device.Object.CreateSurface((uint)width, (uint)height);
                _surfaceWidth = width;
                _surfaceHeight = height;
            }

            protected internal override void MakePreviewPage(uint desiredJobPage, float width, float height)
            {
                EventProvider.Default.Log("MakePreviewPage " + desiredJobPage + " " + width + " x " + height);
                if (desiredJobPage == JOB_PAGE_APPLICATION_DEFINED)
                {
                    desiredJobPage = 1;
                }

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
                    surface = _surface?.Object;
                    target = PreviewTarget?.Object;
                }
                if (surface == null || target == null)
                    return;

                try
                {
                    var page = _pdfDocument.GetPage(desiredJobPage - 1);
                    var renderParams = new PDF_RENDER_PARAMS
                    {
                        BackgroundColor = new D2D_COLOR_F { a = 1, r = 1, g = 1, b = 1 }
                    };
                    _renderer.Object.RenderPageToSurface(page, _surface.Object, new Point(), renderParams);
                    target.DrawPage(desiredJobPage, surface, 96, 96);
                }
                finally
                {
                    _pageBeingWorkedOn = null;
                }
            }

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
                    Extensions.Dispose(ref _surface);
                    Extensions.Dispose(ref _renderer);
                    Extensions.Dispose(ref _d3D11Device);
                }
                base.Dispose(disposing);
            }
        }
    }
}
