using System;
using DirectN.Extensions.Utilities;

namespace CustomPrintDocument;

internal static class Program
{
    [STAThread]
    static void Main()
    {
        using var app = new Application();
        using var win = new MainWindow();
        win.ResizeClient(800, 600);
        win.Center();
        win.Show();
        win.SetForeground();
        app.Run();
    }
}
