using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using CustomPrintDocument.Model;
using Microsoft.UI;
using Microsoft.UI.Windowing;
using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Windows.Graphics;
using Windows.Graphics.Printing;
using Windows.Storage.Pickers;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Storage.Xps.Printing;
using Windows.Win32.UI.WindowsAndMessaging;
using WinRT;

namespace CustomPrintDocument
{
    public sealed partial class MainWindow : Window
    {
        private readonly PrintManager _printMananager;
        private BasePrintDocument _printDocument;

        public MainWindow()
        {
            InitializeComponent();

            // set app icon
            var exeHandle = PInvoke.GetModuleHandle(Path.GetFileName(Process.GetCurrentProcess().MainModule.FileName));
            var appIcon = PInvoke.LoadImage(new HINSTANCE(exeHandle.DangerousGetHandle()), PInvoke.IDI_APPLICATION, GDI_IMAGE_TYPE.IMAGE_ICON, 16, 16, 0);
            AppWindow.SetIcon(Win32Interop.GetIconIdFromIcon(appIcon));

            // size & center
            var area = DisplayArea.GetFromWindowId(AppWindow.Id, DisplayAreaFallback.Nearest);
            var width = 800; var height = 600;
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
                // we can't pass the document directly for some obscure WinRT/CsWin32, etc. reason
                var unk = Marshal.GetIUnknownForObject(_printDocument);
                var doc = MarshalInterface<IPrintDocumentSource>.FromAbi(unk);
                args.SetSource(doc);
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

        private void OnPackageStatusUpdated(object sender, PackageStatusUpdatedEventArgs e)
        {
            if (e.Status.Completion == PrintDocumentPackageCompletion.PrintDocumentPackageCompletion_InProgress)
            {
                DispatcherQueue.TryEnqueue(() =>
                {
                    var doc = (BasePrintDocument)sender;
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
            await PrintManagerInterop.ShowPrintUIForWindowAsync(Win32Interop.GetWindowFromWindowId(AppWindow.Id));
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
}
