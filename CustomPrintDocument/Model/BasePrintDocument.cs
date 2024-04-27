using System;
using System.Runtime.InteropServices;
using Windows.Graphics.Printing;
using Windows.Win32.Foundation;
using Windows.Win32.Storage.Xps.Printing;
using Windows.Win32.System.WinRT;

namespace CustomPrintDocument.Model
{
    // note: we can't use DirectN AOT because there are some issue when authoring COM classes in .NET between ComWrappers, AOT and C#/WinRT
    // and we can't use GeneratedComClass either...
    public abstract class BasePrintDocument : IPrintDocumentSource, IInspectable, IPrintDocumentPageSource, IPrintPreviewPageCollection
    {
        protected BasePrintDocument(string filePath)
        {
            ArgumentNullException.ThrowIfNull(filePath);
            FilePath = filePath;
        }

        public string FilePath { get; }

        unsafe void IInspectable.GetIids(out uint iidCount, Guid** iids)
        {
            iidCount = 0;
            throw new NotImplementedException();
        }

        unsafe HRESULT IInspectable.GetRuntimeClassName(HSTRING* className)
        {
            return HRESULT.E_NOTIMPL; // avoid throwing...
        }

        unsafe void IInspectable.GetTrustLevel(TrustLevel* trustLevel)
        {
            throw new NotImplementedException();
        }

        public virtual IPrintPreviewPageCollection GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget)
        {
            return this;
        }

        public virtual void MakePage(int desiredJobPage, float width, float height)
        {
        }

        public virtual void Paginate(int currentJobPage, nint printTaskOptions)
        {
        }

        public abstract void MakeDocument(nint printTaskOptions, IPrintDocumentPackageTarget docPackageTarget);

        internal unsafe static PCWSTR ToPCWSTR(string text)
        {
            fixed (char* p = text)
                return new PCWSTR(p);
        }
    }

#pragma warning disable SYSLIB1096 // Convert to 'GeneratedComInterface'
    // these are not in CS/Win32 ...
    [ComImport, Guid("a96bb1db-172e-4667-82b5-ad97a252318f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPrintDocumentPageSource
    {
        IPrintPreviewPageCollection GetPreviewPageCollection(IPrintDocumentPackageTarget docPackageTarget);
        void MakeDocument(IntPtr printTaskOptions, IPrintDocumentPackageTarget docPackageTarget);
    };

    [ComImport, Guid("0b31cc62-d7ec-4747-9d6e-f2537d870f2b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IPrintPreviewPageCollection
    {
        void Paginate(int currentJobPage, IntPtr printTaskOptions);
        void MakePage(int desiredJobPage, float width, float height);
    }
#pragma warning restore SYSLIB1096 // Convert to 'GeneratedComInterface'
}
