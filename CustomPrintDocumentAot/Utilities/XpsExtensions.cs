using System;
using DirectN;
using DirectN.Extensions.Com;

namespace CustomPrintDocument.Utilities;

public static class XpsExtensions
{
    public static IComObject<IOpcPartUri> CreatePartUri(this IComObject<IXpsOMObjectFactory> factory, string uri) => CreatePartUri(factory?.Object!, uri);
    public static IComObject<IOpcPartUri> CreatePartUri(this IXpsOMObjectFactory factory, string uri)
    {
        ArgumentNullException.ThrowIfNull(factory);
        ArgumentNullException.ThrowIfNull(uri);
        factory.CreatePartUri(PWSTR.From(uri), out var part).ThrowOnError();
        return new ComObject<IOpcPartUri>(part);
    }

    public static IComObject<IXpsOMImageResource> CreateImageResource(this IComObject<IXpsOMObjectFactory> factory, IStream acquiredStream, XPS_IMAGE_TYPE contentType, IComObject<IOpcPartUri> partUri) => CreateImageResource(factory?.Object!, acquiredStream, contentType, partUri?.Object!);
    public static IComObject<IXpsOMImageResource> CreateImageResource(this IXpsOMObjectFactory factory, IStream acquiredStream, XPS_IMAGE_TYPE contentType, IOpcPartUri partUri)
    {
        ArgumentNullException.ThrowIfNull(factory);
        ArgumentNullException.ThrowIfNull(acquiredStream);
        ArgumentNullException.ThrowIfNull(partUri);
        factory.CreateImageResource(acquiredStream, contentType, partUri, out var resource).ThrowOnError();
        return new ComObject<IXpsOMImageResource>(resource);
    }

    public static IComObject<IXpsOMImageBrush> CreateImageBrush(this IComObject<IXpsOMObjectFactory> factory, IComObject<IXpsOMImageResource> image, XPS_RECT viewBox, XPS_RECT viewPort) => CreateImageBrush(factory?.Object!, image?.Object!, viewBox, viewPort);
    public static IComObject<IXpsOMImageBrush> CreateImageBrush(this IXpsOMObjectFactory factory, IXpsOMImageResource image, XPS_RECT viewBox, XPS_RECT viewPort)
    {
        ArgumentNullException.ThrowIfNull(factory);
        ArgumentNullException.ThrowIfNull(image);
        factory.CreateImageBrush(image, viewBox, viewPort, out var brush).ThrowOnError();
        return new ComObject<IXpsOMImageBrush>(brush);
    }

    public static IComObject<IXpsOMGeometryFigure> CreateGeometryFigure(this IComObject<IXpsOMObjectFactory> factory, XPS_POINT startPoint) => CreateGeometryFigure(factory?.Object!, startPoint);
    public static IComObject<IXpsOMGeometryFigure> CreateGeometryFigure(this IXpsOMObjectFactory factory, XPS_POINT startPoint)
    {
        ArgumentNullException.ThrowIfNull(factory);
        factory.CreateGeometryFigure(startPoint, out var figure).ThrowOnError();
        return new ComObject<IXpsOMGeometryFigure>(figure);
    }

    public static IComObject<IXpsOMPage> CreatePage(this IComObject<IXpsOMObjectFactory> factory, XPS_SIZE pageDimensions, string language, IComObject<IOpcPartUri> partUri) => CreatePage(factory?.Object!, pageDimensions, language, partUri?.Object!);
    public static IComObject<IXpsOMPage> CreatePage(this IXpsOMObjectFactory factory, XPS_SIZE pageDimensions, string language, IOpcPartUri partUri)
    {
        ArgumentNullException.ThrowIfNull(factory);
        ArgumentNullException.ThrowIfNull(partUri);
        factory.CreatePage(pageDimensions, PWSTR.From(language), partUri, out var page).ThrowOnError();
        return new ComObject<IXpsOMPage>(page);
    }

    public static IComObject<IXpsOMPath> CreatePath(this IComObject<IXpsOMObjectFactory> factory) => CreatePath(factory?.Object!);
    public static IComObject<IXpsOMPath> CreatePath(this IXpsOMObjectFactory factory)
    {
        ArgumentNullException.ThrowIfNull(factory);
        factory.CreatePath(out var path).ThrowOnError();
        return new ComObject<IXpsOMPath>(path);
    }

    public static IComObject<IXpsOMGeometry> CreateGeometry(this IComObject<IXpsOMObjectFactory> factory) => CreateGeometry(factory?.Object!);
    public static IComObject<IXpsOMGeometry> CreateGeometry(this IXpsOMObjectFactory factory)
    {
        ArgumentNullException.ThrowIfNull(factory);
        factory.CreateGeometry(out var geometry).ThrowOnError();
        return new ComObject<IXpsOMGeometry>(geometry);
    }

    public static IComObject<IXpsOMPackageWriter> GetXpsOMPackageWriter(this IComObject<IXpsDocumentPackageTarget> target, IComObject<IOpcPartUri> documentSequencePartName, IComObject<IOpcPartUri> discardControlPartName) => GetXpsOMPackageWriter(target?.Object!, documentSequencePartName?.Object!, discardControlPartName?.Object!);
    public static IComObject<IXpsOMPackageWriter> GetXpsOMPackageWriter(this IXpsDocumentPackageTarget target, IOpcPartUri documentSequencePartName, IOpcPartUri discardControlPartName)
    {
        ArgumentNullException.ThrowIfNull(target);
        ArgumentNullException.ThrowIfNull(documentSequencePartName);
        target.GetXpsOMPackageWriter(documentSequencePartName, discardControlPartName, out var writer).ThrowOnError();
        return new ComObject<IXpsOMPackageWriter>(writer);
    }

    public static IComObject<T> GetXpsOMFactory<T>(this IComObject<IXpsDocumentPackageTarget> target) where T : IXpsOMObjectFactory => GetXpsOMFactory<T>(target?.Object!);
    public static IComObject<T> GetXpsOMFactory<T>(this IXpsDocumentPackageTarget target) where T : IXpsOMObjectFactory
    {
        ArgumentNullException.ThrowIfNull(target);
        target.GetXpsOMFactory(out var factory).ThrowOnError();
        return new ComObject<T>(factory);
    }

    public static IComObject<IXpsDocumentPackageTarget> GetPackageTarget(this IComObject<IPrintDocumentPackageTarget> target, Guid guidTargetType, Guid riid) => GetPackageTarget(target?.Object!, guidTargetType, riid);
    public static IComObject<IXpsDocumentPackageTarget> GetPackageTarget(this IPrintDocumentPackageTarget target, Guid guidTargetType, Guid riid)
    {
        ArgumentNullException.ThrowIfNull(target);
        target.GetPackageTarget(guidTargetType, riid, out var unk).ThrowOnError();
        return ComObject.FromPointer<IXpsDocumentPackageTarget>(unk)!;
    }

    public static IComObject<IXpsOMPackage1> CreatePackageFromFile1(this IComObject<IXpsOMObjectFactory1> factory, string fileName, bool reuseObjects) => CreatePackageFromFile1(factory?.Object!, fileName, reuseObjects);
    public static IComObject<IXpsOMPackage1> CreatePackageFromFile1(this IXpsOMObjectFactory1 factory, string fileName, bool reuseObjects)
    {
        ArgumentNullException.ThrowIfNull(factory);
        factory.CreatePackageFromFile1(PWSTR.From(fileName), reuseObjects, out var package).ThrowOnError();
        return new ComObject<IXpsOMPackage1>(package);
    }

    public static IComObject<IXpsOMDocumentSequence> GetDocumentSequence(this IComObject<IXpsOMPackage> package) => GetDocumentSequence(package?.Object!);
    public static IComObject<IXpsOMDocumentSequence> GetDocumentSequence(this IXpsOMPackage package)
    {
        ArgumentNullException.ThrowIfNull(package);
        package.GetDocumentSequence(out var sequence).ThrowOnError();
        return new ComObject<IXpsOMDocumentSequence>(sequence);
    }

    public static IComObject<IXpsOMDocumentCollection> GetDocuments(this IComObject<IXpsOMDocumentSequence> sequence) => GetDocuments(sequence?.Object!);
    public static IComObject<IXpsOMDocumentCollection> GetDocuments(this IXpsOMDocumentSequence sequence)
    {
        ArgumentNullException.ThrowIfNull(sequence);
        sequence.GetDocuments(out var documents).ThrowOnError();
        return new ComObject<IXpsOMDocumentCollection>(documents);
    }

    public static IComObject<IXpsOMDocument> GetAt(this IComObject<IXpsOMDocumentCollection> collection, uint index) => GetAt(collection?.Object!, index);
    public static IComObject<IXpsOMDocument> GetAt(this IXpsOMDocumentCollection collection, uint index)
    {
        ArgumentNullException.ThrowIfNull(collection);
        collection.GetAt(index, out var document).ThrowOnError();
        return new ComObject<IXpsOMDocument>(document);
    }

    public static IComObject<IXpsOMPageReferenceCollection> GetPageReferences(this IComObject<IXpsOMDocument> collection) => GetPageReferences(collection?.Object!);
    public static IComObject<IXpsOMPageReferenceCollection> GetPageReferences(this IXpsOMDocument document)
    {
        ArgumentNullException.ThrowIfNull(document);
        document.GetPageReferences(out var pageRefs).ThrowOnError();
        return new ComObject<IXpsOMPageReferenceCollection>(pageRefs);
    }

    public static uint GetCount(this IComObject<IXpsOMDocumentCollection> collection) => GetCount(collection?.Object!);
    public static uint GetCount(this IXpsOMDocumentCollection collection)
    {
        ArgumentNullException.ThrowIfNull(collection);
        collection.GetCount(out var count).ThrowOnError();
        return count;
    }

    public static uint GetCount(this IComObject<IXpsOMPageReferenceCollection> collection) => GetCount(collection?.Object!);
    public static uint GetCount(this IXpsOMPageReferenceCollection collection)
    {
        ArgumentNullException.ThrowIfNull(collection);
        collection.GetCount(out var count).ThrowOnError();
        return count;
    }

    public static IComObject<IXpsOMPageReference> GetAt(this IComObject<IXpsOMPageReferenceCollection> collection, uint index) => GetAt(collection?.Object!, index);
    public static IComObject<IXpsOMPageReference> GetAt(this IXpsOMPageReferenceCollection collection, uint index)
    {
        ArgumentNullException.ThrowIfNull(collection);
        collection.GetAt(index, out var pageRef).ThrowOnError();
        return new ComObject<IXpsOMPageReference>(pageRef);
    }

    public static IComObject<IXpsOMPage> GetPage(this IComObject<IXpsOMPageReference> collection) => GetPage(collection?.Object!);
    public static IComObject<IXpsOMPage> GetPage(this IXpsOMPageReference collection)
    {
        ArgumentNullException.ThrowIfNull(collection);
        collection.GetPage(out var page).ThrowOnError();
        return new ComObject<IXpsOMPage>(page);
    }

    public static IComObject<IXpsOMVisualCollection> GetVisuals(this IComObject<IXpsOMPage> page) => GetVisuals(page?.Object!);
    public static IComObject<IXpsOMVisualCollection> GetVisuals(this IXpsOMPage page)
    {
        ArgumentNullException.ThrowIfNull(page);
        page.GetVisuals(out var visuals).ThrowOnError();
        return new ComObject<IXpsOMVisualCollection>(visuals);
    }

    public static IComObject<IXpsOMGeometryFigureCollection> GetFigures(this IComObject<IXpsOMGeometry> page) => GetFigures(page?.Object!);
    public static IComObject<IXpsOMGeometryFigureCollection> GetFigures(this IXpsOMGeometry page)
    {
        ArgumentNullException.ThrowIfNull(page);
        page.GetFigures(out var figures).ThrowOnError();
        return new ComObject<IXpsOMGeometryFigureCollection>(figures);
    }
}
