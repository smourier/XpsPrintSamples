using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using DirectN;
using DirectN.Extensions.Com;
using Windows.Graphics.Printing;

namespace CustomPrintDocument.Model;

// note: we can't use DirectN AOT because there are some issue when authoring COM classes in .NET between ComWrappers, AOT and C#/WinRT
// and we can't use GeneratedComClass either...
[GeneratedComClass]
public abstract partial class BasePrintDocument : IPrintDocumentSource, IInspectable, IPrintDocumentPageSource, IPrintPreviewPageCollection, IDisposable
{
    public event EventHandler<PackageStatusUpdatedEventArgs>? PackageStatusUpdated;
    private ComObject<IPrintDocumentPackageTarget>? _docPackageTarget;

    protected BasePrintDocument(string filePath)
    {
        ArgumentNullException.ThrowIfNull(filePath);
        FilePath = filePath;
    }

    public string FilePath { get; }
    public virtual uint? TotalPages { get; protected set; } // when/if known
    protected virtual PrintTarget? PrintTarget { get; set; }

    public virtual void Cancel() => _docPackageTarget?.Object?.Cancel();

    HRESULT IInspectable.GetTrustLevel(out TrustLevel trustLevel) { trustLevel = TrustLevel.BaseTrust; return Constants.S_OK; }
    HRESULT IInspectable.GetRuntimeClassName(out HSTRING className) { className = default; return Constants.S_OK; }
    HRESULT IInspectable.GetIids(out uint iidCount, out nint iids) => throw new NotImplementedException();

    HRESULT IPrintDocumentPageSource.MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
    {
        MakeDocument(printTaskOptions, docPackageTarget);
        return Constants.S_OK;
    }

    HRESULT IPrintDocumentPageSource.GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget, out IPrintPreviewPageCollection docPageCollection)
    {
        docPageCollection = GetPreviewPageCollection(docPackageTarget);
        return Constants.S_OK;
    }

    HRESULT IPrintPreviewPageCollection.MakePage(int desiredJobPage, float width, float height)
    {
        MakePreviewPage(desiredJobPage, width, height);
        return Constants.S_OK;
    }

    HRESULT IPrintPreviewPageCollection.Paginate(int currentJobPage, nint printTaskOptions)
    {
        PreviewPaginate(currentJobPage, printTaskOptions);
        return Constants.S_OK;
    }

    // overridable methods
    protected virtual IPrintPreviewPageCollection GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget)
    {
        docPackageTarget.GetPackageTarget(typeof(IPrintPreviewDxgiPackageTarget).GUID, typeof(IPrintPreviewDxgiPackageTarget).GUID, out var unk);
        var target = DirectN.Extensions.Com.ComObject.FromPointer<IPrintPreviewDxgiPackageTarget>(unk);
        if (target != null)
        {
            PrintTarget = GetPrintTarget(target);
            if (TotalPages.HasValue)
            {
                PrintTarget?.SetJobPageCount(PageCountType.FinalPageCount, TotalPages.Value);
            }
        }
        return this;
    }

    protected virtual void MakePreviewPage(int desiredJobPage, float width, float height) => PrintTarget?.MakePreviewPage(desiredJobPage, width, height);
    protected virtual void PreviewPaginate(int currentJobPage, nint printTaskOptions) => PrintTarget?.PreviewPaginate(currentJobPage, printTaskOptions);
    protected virtual PrintTarget? GetPrintTarget(IComObject<IPrintPreviewDxgiPackageTarget> target) => null;
    protected abstract void MakeDocumentCore(nint printTaskOptions, IComObject<IPrintDocumentPackageTarget> docPackageTarget);
    protected virtual void MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget)
    {
        ArgumentNullException.ThrowIfNull(docPackageTarget);
        _docPackageTarget = new ComObject<IPrintDocumentPackageTarget>(docPackageTarget);

        IConnectionPoint? connectionPoint = null;
        uint cookie = 0;
        if (docPackageTarget is IConnectionPointContainer container)
        {
            container.FindConnectionPoint(typeof(IPrintDocumentPackageStatusEvent).GUID, out connectionPoint);
            if (connectionPoint != null)
            {
                var sink = new StatusSink(this);
                var unk = DirectN.Extensions.Com.ComObject.GetOrCreateComInstance(sink);
                connectionPoint.Advise(unk, out cookie);
            }
        }

        try
        {
            MakeDocumentCore(printTaskOptions, _docPackageTarget);
        }
        catch (COMException exception)
        {
            // eat cancelled error
            if (exception.ErrorCode != unchecked((int)(0x80070000 | (uint)WIN32_ERROR.ERROR_PRINT_CANCELLED)))
                throw;
        }
        finally
        {
            Dispose(true);
            if (connectionPoint != null && cookie != 0)
            {
                connectionPoint.Unadvise(cookie);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    // status update
    protected virtual void OnPackageStatusUpdated(object sender, PackageStatusUpdatedEventArgs e) => PackageStatusUpdated?.Invoke(sender, e);

    [GeneratedComClass]
    private sealed partial class StatusSink(BasePrintDocument document) : IPrintDocumentPackageStatusEvent
    {
        HRESULT IDispatch.GetIDsOfNames(in Guid riid, PWSTR[] rgszNames, uint cNames, uint lcid, int[] rgDispId)
        {
            throw new NotImplementedException();
        }

        HRESULT IDispatch.GetTypeInfo(uint iTInfo, uint lcid, out ITypeInfo ppTInfo)
        {
            throw new NotImplementedException();
        }

        HRESULT IDispatch.GetTypeInfoCount(out uint pctinfo)
        {
            throw new NotImplementedException();
        }

        HRESULT IDispatch.Invoke(int dispIdMember, in Guid riid, uint lcid, DISPATCH_FLAGS wFlags, in DISPPARAMS pDispParams, nint pVarResult, nint pExcepInfo, nint puArgErr)
        {
            throw new NotImplementedException();
        }

        HRESULT IPrintDocumentPackageStatusEvent.PackageStatusUpdated(in PrintDocumentPackageStatus packageStatus)
        {
            document.OnPackageStatusUpdated(document, new PackageStatusUpdatedEventArgs(packageStatus));
            return Constants.S_OK;
        }
    }

    protected virtual void Dispose(bool disposing)
    {
        if (disposing)
        {
            var target = PrintTarget;
            PrintTarget = null;
            target?.Dispose();
            _docPackageTarget?.Dispose();
        }
    }

    ~BasePrintDocument() { Dispose(disposing: false); }
    public void Dispose() { Dispose(disposing: true); GC.SuppressFinalize(this); }
}
