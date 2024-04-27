using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.Versioning;
using DirectN;
using DirectNAot.Extensions.Utilities;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Storage;

[assembly: DisableRuntimeMarshalling]
[assembly: SupportedOSPlatform("windows10.0.19041.0")]

namespace PdfFilePrint
{
    // note: this sample is not AOT-compatible yet, because we need StorageFile & PdfDocument API
    // for which inetrop is built by C#/WinRT which is not yet AOT-compatible https://github.com/microsoft/CsWinRT/discussions/1590
    internal class Program
    {
        static async Task Main()
        {
            //await Print("sample.pdf", "Samsung C480W");
            //await Print("sample.pdf", "Microsoft XPS Document Writer");
            await Print("sample.pdf", "Microsoft Print to PDF");
        }

        static async Task Print(string pdfFilePath, string printerName)
        {
            // load pdf
            var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(pdfFilePath));
            var pdf = await PdfDocument.LoadFromFileAsync(file);

            Functions.CoCreateInstance(Constants.CLSID_PrintDocumentPackageTargetFactory, 0, CLSCTX.CLSCTX_INPROC_SERVER, typeof(IPrintDocumentPackageTargetFactory).GUID, out object obj).ThrowOnError();
            var factory = (IPrintDocumentPackageTargetFactory)obj;

            factory.CreateDocumentPackageTargetForPrintJob(
                new Pwstr(printerName),
                new Pwstr(file.Name),
                null!, // null, send to printer
                null!,
                out var packageTarget).ThrowOnError();

            // register for status changes
            var container = (IConnectionPointContainer)packageTarget;
            container.FindConnectionPoint(typeof(IPrintDocumentPackageStatusEvent).GUID, out var cp).ThrowOnError();

            var status = new PrintDocumentPackageStatus();
            var sink = new StatusSink();
            sink.PackageStatusUpdated += (s, e) => status = e.Status;
            var cw = new StrategyBasedComWrappers();
            var sinkPtr = cw.GetOrCreateComInterfaceForObject(sink, CreateComInterfaceFlags.None);
            cp.Advise(sinkPtr, out var cookie).ThrowOnError();

            // get package target
            // this will usually get us ID_DOCUMENTPACKAGETARGET_MSXPS & ID_DOCUMENTPACKAGETARGET_OPENXPS
            //packageTarget.GetPackageTargetTypes(out var count, out var types).ThrowOnError();

            packageTarget.GetPackageTarget(Constants.ID_DOCUMENTPACKAGETARGET_MSXPS, typeof(IXpsDocumentPackageTarget).GUID, out obj).ThrowOnError();
            var xpsTarget = (IXpsDocumentPackageTarget)obj;
            xpsTarget.GetXpsOMFactory(out var xpsFactory).ThrowOnError();

            // build a writer
            xpsFactory.CreatePartUri(new Pwstr("/seq"), out var seqName);
            xpsFactory.CreatePartUri(new Pwstr("/discard"), out var discardName);
            xpsTarget.GetXpsOMPackageWriter(seqName, discardName, out var writer).ThrowOnError();

            // start
            xpsFactory.CreatePartUri(new Pwstr("/" + file.DisplayName), out var name).ThrowOnError();
            writer.StartNewDocument(name, null!, null!, null!, null!).ThrowOnError();

            var streams = new List<Stream>();
            // browse all PDF pages
            for (uint i = 0; i < pdf.PageCount; i++)
            {
                // render page to stream
                var pdfPage = pdf.GetPage(i);
                using var stream = new MemoryStream();
                await pdfPage.RenderToStreamAsync(stream.AsRandomAccessStream(), new PdfPageRenderOptions { BitmapEncoderId = BitmapEncoder.PngEncoderId });
                var size = new XPS_SIZE { width = (float)pdfPage.Size.Width, height = (float)pdfPage.Size.Height };

                // create image from stream
                stream.Position = 0;

                // note we don't dispose streams here (close would fail)
                var ustream = new DirectNAot.Extensions.Utilities.UnmanagedMemoryStream(stream);
                streams.Add(stream);
                xpsFactory.CreatePartUri(new Pwstr("/image" + i), out var imageUri).ThrowOnError();
                xpsFactory.CreateImageResource(ustream, XPS_IMAGE_TYPE.XPS_IMAGE_TYPE_PNG, imageUri, out var image).ThrowOnError();

                // create a brush from image
                var viewBox = new XPS_RECT { width = size.width, height = size.height };
                xpsFactory.CreateImageBrush(image, viewBox, viewBox, out var imageBrush).ThrowOnError();

                // create a rect figure
                xpsFactory.CreateGeometryFigure(new XPS_POINT(), out var rectFigure).ThrowOnError();
                rectFigure.SetIsClosed(true).ThrowOnError();
                rectFigure.SetIsFilled(true).ThrowOnError();
                var segmentTypes = new XPS_SEGMENT_TYPE[] { XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE };
                var segmentData = new float[] { 0, size.height, size.width, size.height, size.width, 0 };
                rectFigure.SetSegments((uint)segmentTypes.Length, (uint)segmentData.Length, segmentTypes, segmentData, [true, true, true]).ThrowOnError();

                // create a rect geometry from figure
                xpsFactory.CreateGeometry(out var rectGeo).ThrowOnError();
                rectGeo.GetFigures(out var figures).ThrowOnError();
                figures.Append(rectFigure).ThrowOnError();

                // create a path and set rect geometry
                xpsFactory.CreatePath(out var rectPath).ThrowOnError();
                rectPath.SetGeometryLocal(rectGeo).ThrowOnError();

                // set image/brush as brush for rect
                rectPath.SetFillBrushLocal(imageBrush).ThrowOnError();

                // create a page & add add rect to page
                xpsFactory.CreatePartUri(new Pwstr("/page" + i), out var pageUri).ThrowOnError();
                xpsFactory.CreatePage(size, new Pwstr("en"), pageUri, out var page).ThrowOnError(); // note: language is NOT optional

                page.GetVisuals(out var visuals).ThrowOnError();
                visuals.Append(rectPath).ThrowOnError();

                Console.WriteLine("Printing page #" + i + " :" + size.width + " x " + size.height);
                writer.AddPage(page, size, null!, null!, null!, null!).ThrowOnError();
            }
            writer.Close().ThrowOnError();
            streams.Dispose();

            // wait job for completion
            while (status.Completion == PrintDocumentPackageCompletion.PrintDocumentPackageCompletion_InProgress)
            {
                Console.WriteLine("printing ...");
                Thread.Sleep(200);
            }
            cp.Unadvise(cookie).ThrowOnError();
        }
    }

    [GeneratedComClass]
    public partial class StatusSink : IPrintDocumentPackageStatusEvent
    {
        public event EventHandler<PackageStatusUpdatedEventArgs>? PackageStatusUpdated;

        HRESULT IPrintDocumentPackageStatusEvent.PackageStatusUpdated(in PrintDocumentPackageStatus packageStatus)
        {
            PackageStatusUpdated?.Invoke(this, new PackageStatusUpdatedEventArgs(packageStatus));
            Console.WriteLine("Completion: " + packageStatus.Completion);
            return 0;
        }

        // IDispatch
        public HRESULT GetIDsOfNames(in Guid riid, [MarshalUsing(CountElementName = "cNames")] in PWSTR[] rgszNames, uint cNames, uint lcid, [MarshalUsing(CountElementName = "cNames")] out int[] rgDispId) => throw new NotSupportedException();
        public HRESULT GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo) => throw new NotSupportedException();
        public HRESULT GetTypeInfoCount(out uint pctinfo) => throw new NotSupportedException();
        public HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr) => throw new NotSupportedException();
    }

    public class PackageStatusUpdatedEventArgs(PrintDocumentPackageStatus status) : EventArgs
    {
        public PrintDocumentPackageStatus Status => status;
    }
}
