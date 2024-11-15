using DirectN;
using DirectN.Extensions.Com;

namespace CustomPrintDocument.Model;

public abstract class PrintTarget(IComObject<IPrintPreviewDxgiPackageTarget> target) : InterlockedComObject<IPrintPreviewDxgiPackageTarget>(target)
{
    public virtual void InvalidatePreview() => NativeObject.InvalidatePreview();
    public virtual void SetJobPageCount(PageCountType countType, uint count) => NativeObject.SetJobPageCount(countType, count);
    public virtual void DrawPreviewPage(uint jobPageNumber, IDXGISurface pageImage, float dpiX, float dpiY) => NativeObject.DrawPage(jobPageNumber, pageImage, dpiX, dpiY);

    protected abstract internal void PreviewPaginate(int currentJobPage, nint printTaskOptions);
    protected abstract internal void MakePreviewPage(int desiredJobPage, float width, float height);
}
