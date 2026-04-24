using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace CustomPrintDocument.Interop;

// redefined because the one in Windows.Graphics.Printing is not usable due to the IPrintDocumentSource interface
[GeneratedComInterface, Guid("f9f067be-f456-41f0-9c98-5ce73e851410")]
public partial interface IPrintTaskSourceRequestedArgs : IInspectable
{
#pragma warning disable IDE1006 // Naming Styles
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT get_Deadline(out /* Windows::Foundation::DateTime */ long id);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT SetSource([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IPrintDocumentSource>))] IPrintDocumentSource source);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT GetDeferral(out /* IPrintTaskSourceRequestedDeferral */ nint deferral);
#pragma warning restore IDE1006 // Naming Styles
}
