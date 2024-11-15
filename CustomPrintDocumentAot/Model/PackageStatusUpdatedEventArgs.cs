using System;
using DirectN;

namespace CustomPrintDocument.Model;

public class PackageStatusUpdatedEventArgs(PrintDocumentPackageStatus status) : EventArgs
{
    public PrintDocumentPackageStatus Status => status;
}
