using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;
using CustomPrintDocument.Model;
using DirectN;
using DirectN.Extensions.Com;
using DirectN.Extensions.Utilities;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Foundation;
using Windows.Graphics;
using Windows.Graphics.Printing;
using Windows.Storage.Pickers;
using WinRT;

namespace CustomPrintDocument;

public sealed partial class MainWindow : Microsoft.UI.Xaml.Window
{
    private readonly PrintManager _printMananager;
    private BasePrintDocument? _printDocument;

    public MainWindow()
    {
        InitializeComponent();

        // set app icon
        var appIcon = Icon.LoadApplicationIcon()!;
        AppWindow.SetIcon(Win32Interop.GetIconIdFromIcon(appIcon.Handle));

        // size & center
        var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Nearest);
        var width = 800;
        var height = 600;
        var rc = new RectInt32((area.WorkArea.Width - width) / 2, (area.WorkArea.Height - height) / 2, width, height);
        AppWindow.MoveAndResize(rc);

        // register print manager
        _printMananager = PrintManagerInterop.GetForWindow(Win32Interop.GetWindowFromWindowId(AppWindow.Id));
        _printMananager.PrintTaskRequested += OnPrintTaskRequested;
    }

    private void OnPrintTaskRequested(PrintManager sender, PrintTaskRequestedEventArgs args)
    {
        if (_printDocument != null)
        {
            var printTask = args.Request.CreatePrintTask(Path.GetFileName(_printDocument.FilePath), PrintTaskSourceRequested);
            printTask.Completed += (s, e) => DispatcherQueue.TryEnqueue(() => PrintTaskCompleted(e));
        }
    }

    private void PrintTaskSourceRequested(PrintTaskSourceRequestedArgs args)
    {
        if (_printDocument != null)
        {
            args.SetSource(_printDocument);
        }
    }

    private void PrintTaskCompleted(PrintTaskCompletedEventArgs args)
    {
        openButton.Visibility = Visibility.Visible;
        cancelButton.Visibility = Visibility.Collapsed;
        switch (args.Completion)
        {
            case PrintTaskCompletion.Failed:
                status.Title = "Print job has failed.";
                status.Severity = InfoBarSeverity.Error;
                break;

            case PrintTaskCompletion.Abandoned:
                status.Title = "Print job was abandoned.";
                status.Severity = InfoBarSeverity.Warning;
                break;

            case PrintTaskCompletion.Canceled:
                status.Title = "Print job was canceled.";
                status.Severity = InfoBarSeverity.Warning;
                break;

            case PrintTaskCompletion.Submitted:
                status.Title = "Print job was submitted.";
                status.Severity = InfoBarSeverity.Success;
                break;
        }

        if (_printDocument != null)
        {
            _printDocument.Dispose();
            _printDocument.PackageStatusUpdated -= OnPackageStatusUpdated;
            _printDocument = null;
        }
    }

    private void OnPackageStatusUpdated(object? sender, PackageStatusUpdatedEventArgs e)
    {
        if (e.Status.Completion == PrintDocumentPackageCompletion.PrintDocumentPackageCompletion_InProgress)
        {
            DispatcherQueue.TryEnqueue(() =>
            {
                if (sender is not BasePrintDocument doc)
                    return;

                var totalPages = doc.TotalPages.HasValue ? "/" + doc.TotalPages.Value.ToString() : null;
                status.Title = $"Print job in progress: {e.Status.CurrentPage}{totalPages} page(s)";
                status.Severity = InfoBarSeverity.Informational;
                openButton.Visibility = Visibility.Collapsed;
                cancelButton.Visibility = Visibility.Visible;
            });
        }
    }

    private async void OnPrintClicked(object sender, RoutedEventArgs e)
    {
        var fop = new FileOpenPicker();
        WinRT.Interop.InitializeWithWindow.Initialize(fop, Win32Interop.GetWindowFromWindowId(AppWindow.Id));
        fop.FileTypeFilter.Add(".pdf");
        fop.FileTypeFilter.Add(".xps");
        fop.FileTypeFilter.Add(".oxps");
        var file = await fop.PickSingleFileAsync();
        if (file == null)
            return;

        var ext = Path.GetExtension(file.Path).ToLowerInvariant();
        if (ext == ".pdf")
        {
            _printDocument = new PdfPrintDocument(file.Path);
        }
        else if (ext == ".xps" || ext == ".oxps")
        {
            _printDocument = new XpsPrintDocument(file.Path);
        }
        else
            return;

        _printDocument.PackageStatusUpdated += OnPackageStatusUpdated;

        // PrintManagerInterop.ShowPrintUIForWindowAsync is broken for AOT!!
        // see https://github.com/microsoft/CsWinRT/issues/1871
        var pm = ActivationFactory.Get("Windows.Graphics.Printing.PrintManager", typeof(IPrintManagerInterop).GUID);
        var pmi = ComObject.FromPointer<IPrintManagerInterop>(pm.ThisPtr)!;

        var hr = pmi.Object.ShowPrintUIForWindowAsync(new HWND(Win32Interop.GetWindowFromWindowId(AppWindow.Id)), typeof(IAsyncOperation_bool).GUID, out var op);

        var ao = new AsyncOperation<bool>(op);

        Thread.Sleep(10000);
        //var ab = ComObject.FromPointer<IAsyncOperation_bool>(op)!;

        //using var cls = new AsyncOperationCompletedHandler_bool();
        //var unk = ComObject.GetOrCreateComInstance(cls);
        //hr = ab.Object.put_Completed(unk);
        //cls.Wait();

        //var asyncOperation = MarshalInspectable<IAsyncOperation<bool>>.FromAbi(op);
        //var ac = asyncOperation.As<IAsyncOperation<bool>>();
        //var status = asyncOperation.Status;

        //PrintManagerInterop.ShowPrintUIForWindowAsync(Win32Interop.GetWindowFromWindowId(AppWindow.Id)).AsTask();
    }

    [System.Runtime.InteropServices.Marshalling.GeneratedComClass]
    partial class AsyncOperationCompletedHandler_bool : IAsyncOperationCompletedHandler_bool, IDisposable
    {
        private ManualResetEvent? _evt = new(false);

        public void Wait()
        {
            _evt.WaitOne();
        }

        public void Dispose()
        {
            Interlocked.Exchange(ref _evt, null)?.Dispose();
        }

        public AsyncStatus Status { get; private set; }

        public HRESULT Invoke(IAsyncOperation_bool value, AsyncStatus status)
        {
            var hr = value.GetResults(out var results);
            Status = status;
            _evt?.Set();
            return 0;
        }
    }

    private void OnCancelClicked(object sender, RoutedEventArgs e)
    {
        if (_printDocument != null)
        {
            _printDocument.Cancel();
            _printDocument.Dispose();
            _printDocument = null;
        }
    }
}

[System.Runtime.InteropServices.Marshalling.GeneratedComInterface, Guid("cdb5efb3-5788-509d-9be1-71ccb8a3362a")]
public partial interface IAsyncOperation_bool
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetIids(out uint iidCount, out nint iids);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetRuntimeClassName(out HSTRING className);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetTrustLevel(out DirectN.TrustLevel trustLevel);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT put_Completed(nint value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT get_Completed(out nint value);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT GetResults(out nint results);
}

[Guid("cdb5efb3-5788-509d-9be1-71ccb8a3362a")]
public class AsyncOperationBoolean(nint asyncOperation) : AsyncOperation<bool>(asyncOperation)
{
}

public class AsyncOperation<T>(nint asyncOperation) : AsyncOperation(asyncOperation)
{
    public T? Results { get; private set; }

    protected sealed override nint AllocateMemory(int size) => ComWrappersSupport.AllocateVtableMemory(typeof(T), size);
}

public abstract unsafe class AsyncOperation
{
    private readonly nint _handlerVtbl;

    public AsyncOperation(nint asyncOperation)
    {
        if (asyncOperation == 0)
            throw new ArgumentNullException(nameof(asyncOperation));

        var unk = *(nint*)asyncOperation;
        // IAsyncOption<T>'s layout is 9 methods:
        // - IUnknown (3 methods)
        // - IInspectable (3 methods)
        // - put_Completed(void* handler)
        // - get_Completed(void** handler)
        // - GetResults(void** results)

        const int put_CompletedSlot = 6;
        var put_CompletedPtr = *(nint*)(unk + put_CompletedSlot * nint.Size);

        var fn = (delegate* unmanaged[Stdcall]<nint, nint, int>)put_CompletedPtr;

        var size = sizeof(IAsyncOperationCompletedHandlerVTable);
        _handlerVtbl = AllocateMemory(size);
        //_handlerVtbl = Marshal.AllocCoTaskMem(size);
        *(IAsyncOperationCompletedHandlerVTable*)_handlerVtbl = new IAsyncOperationCompletedHandlerVTable();

        var hr = fn(unk, _handlerVtbl);

        //const int GetResultsSlot = 8;
        //var GetResultsPtr = *(nint*)(unk + GetResultsSlot * nint.Size);

        //var getResults = (delegate* unmanaged[Stdcall]<nint, nint*, int>)GetResultsPtr;

        //nint n = 0;
        //var hr = getResults(unk, &n);

    }

    protected abstract nint AllocateMemory(int size);

    public void Wait(nint asyncOperationAbi)
    {
    }

    private struct IAsyncOperationCompletedHandlerVTable
    {
        private static readonly delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int> _queryInterfaceFn = &QueryInterfaceImpl;
        private static readonly delegate* unmanaged[Stdcall]<nint, uint> _addRef = &AddRefImpl;
        private static readonly delegate* unmanaged[Stdcall]<nint, uint> _release = &ReleaseImpl;
        private static readonly delegate* unmanaged[Stdcall]<nint, nint, AsyncStatus, int> _invoke = &InvokeImpl;

        public IAsyncOperationCompletedHandlerVTable()
        {
            //ComWrappers.GetIUnknownImpl(out QueryInterface, out AddRef, out Release);
            QueryInterface = _queryInterfaceFn;
            AddRef = _addRef;
            Release = _release;
            Invoke = _invoke;
            //AddRef = &AddRefImpl;
            //Release = &ReleaseImpl;
            //Invoke = &InvokeImpl;
        }

        // IUnknown
        public delegate* unmanaged[Stdcall]<nint, Guid*, nint*, int> QueryInterface;
        public delegate* unmanaged[Stdcall]<nint, uint> AddRef;
        public delegate* unmanaged[Stdcall]<nint, uint> Release;
        //public nint QueryInterface;
        //public nint AddRef;
        //public nint Release;

        // IAsyncOperationCompletedHandler
        public delegate* unmanaged[Stdcall]<nint, nint, AsyncStatus, int> Invoke;

        [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
        private static int QueryInterfaceImpl(nint pThis, Guid* riid, nint* ppv)
        {
            return Constants.E_NOINTERFACE;
        }

        [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
        private static uint AddRefImpl(nint pThis) => 1;

        [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
        private static uint ReleaseImpl(nint pThis) => 1;

        [UnmanagedCallersOnly(CallConvs = [typeof(CallConvStdcall)])]
        private static int InvokeImpl(nint pThis, nint asyncInfo, AsyncStatus status)
        {
            return 0;
        }
    }
}

[System.Runtime.InteropServices.Marshalling.GeneratedComInterface, Guid("c1d3d1a2-ae17-5a5f-b5a2-bdcc8844889a")]
public partial interface IAsyncOperationCompletedHandler_bool
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Invoke(IAsyncOperation_bool value, AsyncStatus status);
}
