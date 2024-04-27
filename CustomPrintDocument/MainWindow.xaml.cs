using System;
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
            status.IsOpen = true;
        }

        private async void print_Click(object sender, RoutedEventArgs e)
        {
            var fop = new FileOpenPicker();
            WinRT.Interop.InitializeWithWindow.Initialize(fop, Win32Interop.GetWindowFromWindowId(AppWindow.Id));
            fop.FileTypeFilter.Add(".pdf");
            fop.FileTypeFilter.Add(".xps");
            var file = await fop.PickSingleFileAsync();
            if (file == null)
                return;

            var ext = Path.GetExtension(file.Path).ToLowerInvariant();
            if (ext == ".pdf")
            {
                _printDocument = new PdfPrintDocument(file.Path);
            }
            else if (ext == ".xps")
            {
                _printDocument = new XpsPrintDocument(file.Path);
            }
            else
                return;

            await PrintManagerInterop.ShowPrintUIForWindowAsync(Win32Interop.GetWindowFromWindowId(AppWindow.Id));
        }
    }
}
