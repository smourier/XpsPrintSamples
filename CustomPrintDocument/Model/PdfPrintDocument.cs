using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Storage.Packaging.Opc;
using Windows.Win32.Storage.Xps;
using Windows.Win32.Storage.Xps.Printing;
using Windows.Win32.System.Com;

namespace CustomPrintDocument.Model
{
    public class PdfPrintDocument(string filePath) : BasePrintDocument(filePath)
    {
        protected override async Task MakeDocumentAsync(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            ArgumentNullException.ThrowIfNull(docPackageTarget);
            // can use options for various customizations
            // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);

            var container = (IConnectionPointContainer)docPackageTarget;

            // load pdf
            var file = await StorageFile.GetFileFromPathAsync(Path.GetFullPath(FilePath));
            var pdf = await PdfDocument.LoadFromFileAsync(file);
            TotalPages = pdf.PageCount;

            var xpsTarget = GetXpsDocumentPackageTarget(docPackageTarget);
            var factory = xpsTarget.GetXpsOMFactory();

            // build a writer
            var seqName = factory.CreatePartUri("/seq");
            var discardName = factory.CreatePartUri("/discard");
            var writer = xpsTarget.GetXpsOMPackageWriter(seqName, discardName);

            // start
            var name = factory.CreatePartUri("/" + file.DisplayName);
            writer.StartNewDocument(name, null!, null!, null!, null!);

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
    }
}
