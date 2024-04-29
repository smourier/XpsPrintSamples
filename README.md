# XpsPrintSamples
Xps and PDF file printing samples

* **XpsFilePrint**: a project (AOT friendly) that demonstrates how to print an XPS file (to any printer, including XPS, PDF, ...)
* **PdfFilePrint**: a project (AOT friendly) that demonstrates how to print a PDF file (to any printer, including XPS, PDF, ...)
* **CustomPrintDocument**: a WinUI3 project (which is not AOT-friendly) that implements a custom WinRT [IPrintDocumentSource](https://learn.microsoft.com/en-us/uwp/api/windows.graphics.printing.iprintdocumentsource), usable without XAML. It supports printing XPS or PDF files (including preview). For PDF, it implements two printing mode: one based on **XPS** and another based on **Direct2D** which outputs higher quality PDF prints.

![image](https://github.com/smourier/XpsPrintSamples/assets/5328574/5536ca26-eb92-46f3-812b-55eed74a1244)
