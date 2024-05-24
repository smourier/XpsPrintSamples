using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using DirectN;
using DirectN.Extensions.Com;

[assembly: DisableRuntimeMarshalling]
[assembly: SupportedOSPlatform("windows10.0.19041.0")]

namespace XpsFilePrint
{
    internal partial class Program
    {
        static void Main()
        {
            PrintJob("sample.xps", "Samsung C480W");
            //PrintJob("sample.xps", "Microsoft XPS Document Writer");
            //PrintJob("sample.xps", "Microsoft Print to PDF");

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        static void PrintJob(string filePath, string printerName)
        {
            using var factory = ComObject<IPrintDocumentPackageTargetFactory>.CoCreate(Constants.CLSID_PrintDocumentPackageTargetFactory)!;
            factory.Object.CreateDocumentPackageTargetForPrintJob(
                PWSTR.From(printerName),
                PWSTR.From(Path.GetFileName(filePath)),
                null!, // null, send to printer
                null!,
                out var packageTarget).ThrowOnError();

            // register for status changes
            var container = (IConnectionPointContainer)packageTarget;
            container.FindConnectionPoint(typeof(IPrintDocumentPackageStatusEvent).GUID, out var cp).ThrowOnError();

            var status = new PrintDocumentPackageStatus();
            var sink = new StatusSink();
            sink.PackageStatusUpdated += (s, e) => status = e.Status;
            var sinkPtr = ComObject.ComWrappers.GetOrCreateComInterfaceForObject(sink, CreateComInterfaceFlags.None);
            cp.Advise(sinkPtr, out var cookie).ThrowOnError();

            // get package target
            // this will usually get us ID_DOCUMENTPACKAGETARGET_MSXPS & ID_DOCUMENTPACKAGETARGET_OPENXPS
            //packageTarget.GetPackageTargetTypes(out var count, out var types).ThrowOnError();

            packageTarget.GetPackageTarget(Constants.ID_DOCUMENTPACKAGETARGET_MSXPS, typeof(IXpsDocumentPackageTarget).GUID, out var unk2).ThrowOnError();
            using var xpsTarget = DirectN.Extensions.Com.ComObject.FromPointer<IXpsDocumentPackageTarget>(unk2)!;
            xpsTarget.Object.GetXpsOMFactory(out var xpsFactory).ThrowOnError();

            // load xps file
            xpsFactory.CreatePackageFromFile(PWSTR.From(filePath), true, out var xpsPackage).ThrowOnError();

            // build a writer
            xpsPackage.GetDiscardControlPartName(out var discardName).ThrowOnError();
            xpsPackage.GetDocumentSequence(out var seq).ThrowOnError();
            seq.GetPartName(out var seqName).ThrowOnError();
            xpsTarget.Object.GetXpsOMPackageWriter(seqName, discardName, out var writer).ThrowOnError();

            seq.GetDocuments(out var docs).ThrowOnError();
            docs.GetCount(out var docCount).ThrowOnError();

            xpsFactory.CreatePartUri(PWSTR.From("/" + Path.GetFileNameWithoutExtension(filePath)), out var name).ThrowOnError();
            writer.StartNewDocument(name, null!, null!, null!, null!).ThrowOnError();

            // browse all docs (usually one)
            for (uint i = 0; i < docCount; i++)
            {
                docs.GetAt(i, out var doc).ThrowOnError();
                doc.GetPageReferences(out var pagesRef).ThrowOnError();

                // browse all pages
                pagesRef.GetCount(out var pagesCount).ThrowOnError();
                for (uint j = 0; j < pagesCount; j++)
                {
                    pagesRef.GetAt(j, out var pageRef).ThrowOnError();
                    pageRef.GetPage(out var page).ThrowOnError();
                    page.GetPageDimensions(out var size).ThrowOnError();
                    Console.WriteLine("Printing page #" + j + " :" + size.width + " x " + size.height);
                    writer.AddPage(page, size, null!, null!, null!, null!).ThrowOnError();
                }
            }
            writer.Close().ThrowOnError();

            // wait job for completion
            while (status.Completion == PrintDocumentPackageCompletion.PrintDocumentPackageCompletion_InProgress)
            {
                Console.WriteLine("printing ...");
                Thread.Sleep(200);
            }
            cp.Unadvise(cookie).ThrowOnError();
        }
    }

    [System.Runtime.InteropServices.Marshalling.GeneratedComClass]
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
        public HRESULT GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId) => throw new NotSupportedException();
        public HRESULT GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo) => throw new NotSupportedException();
        public HRESULT GetTypeInfoCount(out uint pctinfo) => throw new NotSupportedException();
        public HRESULT Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr) => throw new NotSupportedException();
    }

    public class PackageStatusUpdatedEventArgs(PrintDocumentPackageStatus status) : EventArgs
    {
        public PrintDocumentPackageStatus Status => status;
    }
}
