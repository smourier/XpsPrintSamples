using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;
using DirectN;

namespace CustomPrintDocument.Interop;

[GeneratedComInterface, Guid("c1d3d1a2-ae17-5a5f-b5a2-bdcc8844889a")]
public partial interface IAsyncActionCompletedHandlerBoolean
{
    [PreserveSig]
    [return: MarshalAs(UnmanagedType.Error)]
    HRESULT Invoke(nint asyncInfo, DirectN.AsyncStatus asyncStatus);
}
