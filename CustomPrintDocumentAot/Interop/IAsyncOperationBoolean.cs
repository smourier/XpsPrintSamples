using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace CustomPrintDocument.Interop;

// 	IAsyncOperation<bool>
[GeneratedComInterface, Guid("cdb5efb3-5788-509d-9be1-71ccb8a3362a")]
public partial interface IAsyncOperationBoolean : IInspectable
{
#pragma warning disable IDE1006 // Naming Styles
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT put_Completed([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IAsyncActionCompletedHandlerBoolean>))] IAsyncActionCompletedHandlerBoolean handler);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT get_Completed([MarshalUsing(typeof(UniqueComInterfaceMarshaller<IAsyncActionCompletedHandlerBoolean>))] out IAsyncActionCompletedHandlerBoolean handler);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT GetResults([MarshalAs(UnmanagedType.Bool)] out bool result);
#pragma warning restore IDE1006 // Naming Styles
}
