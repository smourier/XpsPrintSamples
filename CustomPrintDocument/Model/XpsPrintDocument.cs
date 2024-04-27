using System;
using Windows.Win32;
using Windows.Win32.Storage.Xps;
using Windows.Win32.Storage.Xps.Printing;

namespace CustomPrintDocument.Model
{
    public class XpsPrintDocument(string filePath) : BasePrintDocument(filePath)
    {
        unsafe public override void MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
        {
            ArgumentNullException.ThrowIfNull(docPackageTarget);
            // can use options for various customizations
            // var options = MarshalInterface<PrintTaskOptions>.FromAbi(printTaskOptions);

            var IXpsDocumentPackageTargetGuid = typeof(IXpsDocumentPackageTarget).GUID;
            var xpsGuid = PInvoke.ID_DOCUMENTPACKAGETARGET_MSXPS;
            docPackageTarget.GetPackageTarget(&xpsGuid, &IXpsDocumentPackageTargetGuid, out var obj);

            var target = (IXpsDocumentPackageTarget)obj;
            var fac = target.GetXpsOMFactory();

            // load the file as an XPS package (could use a stream too)
            var pack = fac.CreatePackageFromFile(ToPCWSTR(FilePath), true);

            // build a package writer for printer
            var discard = pack.GetDiscardControlPartName();
            var seq = pack.GetDocumentSequence();
            var seqName = seq.GetPartName();
            var writer = target.GetXpsOMPackageWriter(seqName, discard);

            // start a printer document
            var name = fac.CreatePartUri(ToPCWSTR("/name"));
            writer.StartNewDocument(name, null, null, null, null);

            var docs = seq.GetDocuments();
            var count = docs.GetCount();

            // add all documents
            for (uint i = 0; i < count; i++)
            {
                var doc = docs.GetAt(i);
                var pagesRef = doc.GetPageReferences();

                // add all pages
                var pagesCount = pagesRef.GetCount();
                for (uint j = 0; j < pagesCount; j++)
                {
                    var pageRef = pagesRef.GetAt(j);
                    var page = pageRef.GetPage();
                    writer.AddPage(page, null, null, null, null, null);
                }
            }
            writer.Close();
        }
    }
}
