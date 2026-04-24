using System;
using System.IO;
using CustomPrintDocument.Interop;
using CustomPrintDocument.Model;
using CustomPrintDocument.Utilities;
using DirectN;
using DirectN.Extensions.Com;
using DirectN.Extensions.Utilities;
using ShellN.Extensions.Utilities;
using Windows.Graphics.Printing;

namespace CustomPrintDocument;

internal class MainWindow : Window
{
    private PrintManager? _printManager;
    private IComObject<IPrintManagerInterop>? _printManagerInterop;
    private BasePrintDocument? _printDocument;
    private string? _message;
    private bool _canClose = true;

    public MainWindow()
        : base("Custom Print Document")
    {
        CreateButton("Open XPS or PDF file to print...", 10, 10, 300, 24, 1);
    }

    protected override LRESULT? WindowProc(HWND hwnd, uint msg, WPARAM wParam, LPARAM lParam)
    {
        if (msg == MessageDecoder.WM_COMMAND)
        {
            var id = wParam.Value.LOWORD();
            switch (id)
            {
                case 1:
                    var fod = new FileOpenDialog();
                    fod.SetOptions(FileOpenDialog.DefaultOptions | ShellN.FILEOPENDIALOGOPTIONS.FOS_FORCEFILESYSTEM);
                    fod.SetFileTypes([
                        $"PDF Files (*.pdf)|*.pdf",
                        $"XPS Files (*.xps)|*.xps",
                        $"All Files (*.*)|*.*"
                        ]);
                    if (fod.Show(hwnd))
                    {
                        var path = fod.GetResult()?.SIGDN_FILESYSPATH;
                        if (path != null)
                        {
                            Print(path);
                        }
                    }
                    break;

                default:
                    return 0; // handled
            }
        }

        return base.WindowProc(hwnd, msg, wParam, lParam);
    }

    protected override bool OnPaint(HDC hdc, PAINTSTRUCT ps)
    {
        if (_message != null)
        {
            var font = Functions.GetStockObject(GET_STOCK_OBJECT_FLAGS.SYSTEM_FONT);
            var oldFont = Functions.SelectObject(hdc, font);
            Functions.TextOutW(hdc, 10, 50, PWSTR.From(_message), _message.Length);
            Functions.SelectObject(hdc, oldFont);
        }
        return base.OnPaint(hdc, ps);
    }

    private void SetMessage(string? message)
    {
        _message = message;
        Invalidate();
    }

    protected override bool OnClosing()
    {
        if (!_canClose)
            return true;

        return base.OnClosing();
    }

    protected override void OnHandleCreated(object? sender, EventArgs e)
    {
        base.OnHandleCreated(sender, e);

        // equivalent of this
        // _printManager = PrintManagerInterop.GetForWindow(Handle);
        _printManagerInterop = ComObject.GetActivationFactory<IPrintManagerInterop>("Windows.Graphics.Printing.PrintManager");
        if (_printManagerInterop != null)
        {
            if (_printManagerInterop.Object.GetForWindow(Handle, typeof(DirectN.IInspectable).GUID, out var pm).IsOk)
            {
                _printManager = PrintManager.FromAbi(pm);
                _printManager.PrintTaskRequested += OnPrintTaskRequested;
            }
        }
    }

    private void Print(string filePath)
    {
        _printDocument?.Dispose();
        _printDocument = null;

        var ext = Path.GetExtension(filePath);
        if (ext.EqualsIgnoreCase(".pdf"))
        {
            _printDocument = new PdfPrintDocument(filePath);
        }
        else if (ext.EqualsIgnoreCase(".xps"))
        {
            _printDocument = new XpsPrintDocument(filePath);
        }
        else
            return;

        _printDocument.PackageStatusUpdated += OnPackageStatusUpdated;

        // equivalent of this but it doesn't work (pb with CsWinRT support https://github.com/microsoft/CsWinRT/issues/1871)
        // var ret = await PrintManagerInterop.ShowPrintUIForWindowAsync(Handle);
        _printManagerInterop!.Object.ShowPrintUIForWindowAsync(Handle, typeof(Interop.IAsyncInfo).GUID, out var op).ThrowOnError();
        var info = ComObject.FromPointer<Interop.IAsyncInfo>(op);
        var ac = info.As<Interop.IAsyncOperationBoolean>(true)!;
        var handler = new Interop.AsyncActionCompletedHandlerBoolean();
        handler.StatusChanged += (s, e) =>
            {
                var status = e.Value;
                Application.TraceInfo($"Print UI status: {status}");
            };

        ac.Object.put_Completed(handler).ThrowOnError();
    }

    private void OnPackageStatusUpdated(object? sender, PackageStatusUpdatedEventArgs e)
    {
        if (e.Status.Completion == PrintDocumentPackageCompletion.PrintDocumentPackageCompletion_InProgress)
        {
            if (!_canClose)
            {
                var menu = Functions.GetSystemMenu(Handle, false);
                Functions.EnableMenuItem(menu, (uint)SC.SC_CLOSE, MENU_ITEM_FLAGS.MF_BYCOMMAND | MENU_ITEM_FLAGS.MF_DISABLED | MENU_ITEM_FLAGS.MF_GRAYED);
            }

            _canClose = false;
            var doc = (BasePrintDocument)sender!;
            var totalPages = doc.TotalPages.HasValue ? "/" + doc.TotalPages.Value.ToString() : null;
            SetMessage($"Print job in progress: {e.Status.CurrentPage}{totalPages} page(s)");
            //Application.TraceInfo($"JobId: {e.Status.JobId} Page: {e.Status.CurrentPage} Total: {e.Status.CurrentPageTotal} Status: {e.Status.PackageStatus} Completion: {e.Status.Completion}");
        }
    }

    private void OnPrintTaskRequested(PrintManager sender, PrintTaskRequestedEventArgs args)
    {
        if (_printDocument != null)
        {
            var printTask = args.Request.CreatePrintTask(Path.GetFileName(_printDocument.FilePath), PrintTaskSourceRequested);
            printTask.Completed += PrintTaskCompleted;
        }
    }

    private void PrintTaskSourceRequested(PrintTaskSourceRequestedArgs args)
    {
        if (_printDocument != null)
        {
            using var a = args.AsComObject<IPrintTaskSourceRequestedArgs>();
            a.Object.SetSource(_printDocument).ThrowOnError();
        }
    }

    private void PrintTaskCompleted(PrintTask sender, PrintTaskCompletedEventArgs args)
    {
        switch (args.Completion)
        {
            case PrintTaskCompletion.Failed:
                SetMessage("Print job has failed.");
                break;

            case PrintTaskCompletion.Abandoned:
                SetMessage("Print job was abandoned.");
                break;

            case PrintTaskCompletion.Canceled:
                SetMessage("Print job was canceled.");
                break;

            case PrintTaskCompletion.Submitted:
                SetMessage("Print job was submitted.");
                break;
        }

        if (_printDocument != null)
        {
            _printDocument.Dispose();
            _printDocument.PackageStatusUpdated -= OnPackageStatusUpdated;
            _printDocument = null;
        }

        var menu = Functions.GetSystemMenu(Handle, false);
        Functions.EnableMenuItem(menu, (uint)SC.SC_CLOSE, MENU_ITEM_FLAGS.MF_BYCOMMAND);
        _canClose = true;
    }
}
