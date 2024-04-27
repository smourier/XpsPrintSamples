using System;
using System.Collections.Generic;
using System.IO;
using Windows.Data.Pdf;
using Windows.Graphics.Imaging;
using Windows.Storage;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Storage.Xps;
using Windows.Win32.Storage.Xps.Printing;

namespace CustomPrintDocument.Model
{
    public class PdfPrintDocument(string filePath) : BasePrintDocument(filePath)
    {
        unsafe public override void MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            ArgumentNullException.ThrowIfNull(docPackageTarget);
            // can use options for various customizations
            // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);

            // load pdf
            var file = StorageFile.GetFileFromPathAsync(Path.GetFullPath(FilePath)).GetResults();
            var pdf = PdfDocument.LoadFromFileAsync(file).GetResults();

            var IXpsDocumentPackageTargetGuid = typeof(IXpsDocumentPackageTarget).GUID;
            var xpsGuid = PInvoke.ID_DOCUMENTPACKAGETARGET_MSXPS;
            docPackageTarget.GetPackageTarget(&xpsGuid, &IXpsDocumentPackageTargetGuid, out var obj);
            var xpsTarget = (IXpsDocumentPackageTarget)obj;

            var xpsFactory = xpsTarget.GetXpsOMFactory();

            // build a writer
            var seqName = xpsFactory.CreatePartUri(ToPCWSTR("/seq"));
            var discardName = xpsFactory.CreatePartUri(ToPCWSTR("/discard"));
            var writer = xpsTarget.GetXpsOMPackageWriter(seqName, discardName);

            // start
            var name = xpsFactory.CreatePartUri(ToPCWSTR("/" + file.DisplayName));
            writer.StartNewDocument(name, null!, null!, null!, null!);

            var streams = new List<Stream>();
            // browse all PDF pages
            for (uint i = 0; i < pdf.PageCount; i++)
            {
                // render page to stream
                var pdfPage = pdf.GetPage(i);
                using var stream = new MemoryStream();
                pdfPage.RenderToStreamAsync(stream.AsRandomAccessStream(), new PdfPageRenderOptions { BitmapEncoderId = BitmapEncoder.PngEncoderId }).GetResults();
                var size = new XPS_SIZE { width = (float)pdfPage.Size.Width, height = (float)pdfPage.Size.Height };

                // create image from stream
                stream.Position = 0;

                // note we don't dispose streams here (close would fail)
                var ustream = new UnmanagedMemoryStream(stream);
                streams.Add(stream);
                var imageUri = xpsFactory.CreatePartUri(ToPCWSTR("/image" + i));
                var image = xpsFactory.CreateImageResource(ustream, XPS_IMAGE_TYPE.XPS_IMAGE_TYPE_PNG, imageUri);

                // create a brush from image
                var viewBox = new XPS_RECT { width = size.width, height = size.height };
                var imageBrush = xpsFactory.CreateImageBrush(image, viewBox, viewBox);

                // create a rect figure
                var rectFigure = xpsFactory.CreateGeometryFigure(new XPS_POINT());
                rectFigure.SetIsClosed(true);
                rectFigure.SetIsFilled(true);
                var segmentTypes = new XPS_SEGMENT_TYPE[] { XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE, XPS_SEGMENT_TYPE.XPS_SEGMENT_TYPE_LINE };
                var segmentData = new float[] { 0, size.height, size.width, size.height, size.width, 0 };
                var segmentStrokes = new bool[] { true, true, true };

                // SetSegments def is wrong https://github.com/microsoft/win32metadata/issues/1889
                fixed (float* f = segmentData)
                fixed (XPS_SEGMENT_TYPE* st = segmentTypes)
                fixed (bool* ss = segmentStrokes)
                    rectFigure.SetSegments((uint)segmentTypes.Length, (uint)segmentData.Length, st, *f, (BOOL*)ss);

                // create a rect geometry from figure
                var rectGeo = xpsFactory.CreateGeometry();
                var figures = rectGeo.GetFigures();
                figures.Append(rectFigure);

                // create a path and set rect geometry
                var rectPath = xpsFactory.CreatePath();
                rectPath.SetGeometryLocal(rectGeo);

                // set image/brush as brush for rect
                rectPath.SetFillBrushLocal(imageBrush);

                // create a page & add add rect to page
                var pageUri = xpsFactory.CreatePartUri(ToPCWSTR("/page" + i));
                var page = xpsFactory.CreatePage(&size, ToPCWSTR("en"), pageUri); // note: language is NOT optional

                var visuals = page.GetVisuals();
                visuals.Append(rectPath);

                writer.AddPage(page, size, null!, null!, null!, null!);
            }
            writer.Close();
            foreach (var stream in streams)
            {
                stream.Dispose();
            }
        }
    }
}
