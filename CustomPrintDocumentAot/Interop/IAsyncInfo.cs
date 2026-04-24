using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace CustomPrintDocument.Interop;

[GeneratedComInterface, Guid("00000036-0000-0000-C000-000000000046")]
public partial interface IAsyncInfo : IInspectable
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT get_Id(out uint id);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT get_Status(out DirectN.AsyncStatus status);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT get_ErrorCode(out DirectN.HRESULT errorCode);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT Cancel();

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT Close();
}
