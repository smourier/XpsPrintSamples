using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace CustomPrintDocument.Interop;

// we can't use the existing one, so it's easier to redefine one,
// see here https://github.com/microsoft/CsWinRT/issues/1722
[GeneratedComInterface, Guid("dedc0c30-f1eb-47df-aae6-ed5427511f01")]
public partial interface IPrintDocumentSource : IInspectable
{
}
