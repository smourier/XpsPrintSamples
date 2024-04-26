using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using System.Runtime.Versioning;
using DirectN;
using DirectNAot.Extensions.Utilities;

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
            //GC.WaitForPendingFinalizers(); // this crashes .NET 8's ComObject (waiting for .NET 9...) see https://github.com/dotnet/runtime/issues/96901
        }

        static void PrintJob(string filePath, string printerName)
        {
            Functions.CoCreateInstance(Constants.CLSID_PrintDocumentPackageTargetFactory, 0, CLSCTX.CLSCTX_INPROC_SERVER, typeof(IPrintDocumentPackageTargetFactory).GUID, out object obj).ThrowOnError();
            var factory = (IPrintDocumentPackageTargetFactory)obj;

            factory.CreateDocumentPackageTargetForPrintJob(
                new Pwstr(printerName),
                new Pwstr(Path.GetFileName(filePath)),
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

            // load xps file
            xpsFactory.CreatePackageFromFile(new Pwstr(filePath), true, out var xpsPackage).ThrowOnError();

            // build a writer
            xpsPackage.GetDiscardControlPartName(out var discardName).ThrowOnError();
            xpsPackage.GetDocumentSequence(out var seq).ThrowOnError();
            seq.GetPartName(out var seqName).ThrowOnError();
            xpsTarget.GetXpsOMPackageWriter(seqName, discardName, out var writer).ThrowOnError();

            seq.GetDocuments(out var docs).ThrowOnError();
            docs.GetCount(out var docCount).ThrowOnError();

            xpsFactory.CreatePartUri(new Pwstr("/" + Path.GetFileNameWithoutExtension(filePath)), out var name).ThrowOnError();
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
