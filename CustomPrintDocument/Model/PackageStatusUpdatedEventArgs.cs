using System;
using Windows.Win32.Storage.Xps.Printing;

namespace CustomPrintDocument.Model
{
    public class PackageStatusUpdatedEventArgs(PrintDocumentPackageStatus status) : EventArgs
    {
        public PrintDocumentPackageStatus Status => status;
    }
}
