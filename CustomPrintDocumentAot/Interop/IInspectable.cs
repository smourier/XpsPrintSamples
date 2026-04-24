using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace CustomPrintDocument.Interop;

[GeneratedComInterface, Guid("af86e2e0-b12d-4c6a-9c5a-d7aa65101e90")]
public partial interface IInspectable
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT GetIids(out uint iidCount, out nint iids);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT GetRuntimeClassName(out DirectN.HSTRING className);

    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    DirectN.HRESULT GetTrustLevel(out DirectN.TrustLevel trustLevel);
}
