using System;
using System.Runtime.InteropServices;
using System.Threading;
using Windows.Win32;
using Windows.Win32.Foundation;
using Windows.Win32.Graphics.Direct2D;
using Windows.Win32.Graphics.Direct3D;
using Windows.Win32.Graphics.Direct3D11;
using Windows.Win32.Graphics.Dxgi;
using Windows.Win32.Graphics.Dxgi.Common;
using Windows.Win32.Graphics.Imaging;
using Windows.Win32.System.Com;

namespace CustomPrintDocument.Utilities
{
    internal static class Extensions
    {
        public static UnknownObject<IDXGISurface> CreateSurface(this ID3D11Device device, uint width, uint height)
        {
            ArgumentNullException.ThrowIfNull(device);
            var texture = new D3D11_TEXTURE2D_DESC
            {
                ArraySize = 1,
                BindFlags = D3D11_BIND_FLAG.D3D11_BIND_SHADER_RESOURCE | D3D11_BIND_FLAG.D3D11_BIND_RENDER_TARGET,
                Format = DXGI_FORMAT.DXGI_FORMAT_B8G8R8A8_UNORM,
                MipLevels = 1,
                SampleDesc = new DXGI_SAMPLE_DESC { Count = 1 },
                Width = width,
                Height = height,
            };
            device.CreateTexture2D(texture, null, out var tex);
            return new UnknownObject<IDXGISurface>((IDXGISurface)tex);
        }

        public static UnknownObject<ID3D11Device> CreateD3D11Device()
        {
            var flags = D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_BGRA_SUPPORT; // D2D
#if DEBUG
            flags |= D3D11_CREATE_DEVICE_FLAG.D3D11_CREATE_DEVICE_DEBUG;
#endif
            unsafe
            {
                PInvoke.D3D11CreateDevice(
                    null,
                    D3D_DRIVER_TYPE.D3D_DRIVER_TYPE_HARDWARE,
                    HMODULE.Null,
                    flags,
                    null,
                    0,
                    PInvoke.D3D11_SDK_VERSION,
                    out var d3D11Device,
                    null,
                    out var deviceContext).ThrowOnFailure();
                Marshal.ReleaseComObject(deviceContext);
                ((ID3D11Multithread)d3D11Device).SetMultithreadProtected(true);
                return new UnknownObject<ID3D11Device>(d3D11Device);
            }
        }

        public static UnknownObject<ID2D1Factory1> CreateD2D1Factory()
        {
            var options = new D2D1_FACTORY_OPTIONS();
#if DEBUG
            options.debugLevel = D2D1_DEBUG_LEVEL.D2D1_DEBUG_LEVEL_WARNING;
#endif
            PInvoke.D2D1CreateFactory(D2D1_FACTORY_TYPE.D2D1_FACTORY_TYPE_MULTI_THREADED, typeof(ID2D1Factory1).GUID, options, out var obj).ThrowOnFailure();
            return new UnknownObject<ID2D1Factory1>((ID2D1Factory1)obj);
        }

        public static UnknownObject<IWICImagingFactory> CreateWICImagingFactory()
        {
            PInvoke.CoCreateInstance(PInvoke.CLSID_WICImagingFactory, null, CLSCTX.CLSCTX_ALL, typeof(IWICImagingFactory).GUID, out var obj).ThrowOnFailure();
            return new UnknownObject<IWICImagingFactory>((IWICImagingFactory)obj);
        }

        public unsafe static PCWSTR ToPCWSTR(this string text)
        {
            fixed (char* p = text)
                return new PCWSTR(p);
        }

        public static void Dispose<T>(ref UnknownObject<T> unknownObject) => Interlocked.Exchange(ref unknownObject, null)?.Dispose();
        public static void Dispose(ref IDisposable disposable) { try { Interlocked.Exchange(ref disposable, null)?.Dispose(); } catch { /* do nothing */ } }
    }
}
