using System;
using DirectN;
using DirectN.Extensions.Utilities;

namespace CustomPrintDocument.Interop;

[System.Runtime.InteropServices.Marshalling.GeneratedComClass]
public sealed partial class AsyncActionCompletedHandlerBoolean : IAsyncActionCompletedHandlerBoolean
{
    public event EventHandler<ValueEventArgs<AsyncStatus>>? StatusChanged;

    HRESULT IAsyncActionCompletedHandlerBoolean.Invoke(nint asyncInfo, AsyncStatus asyncStatus)
    {
        StatusChanged?.Invoke(this, new ValueEventArgs<AsyncStatus>(asyncStatus));
        return Constants.S_OK;
    }
}
