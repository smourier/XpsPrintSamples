using System.Runtime.CompilerServices;
using Microsoft.UI.Xaml;

[assembly: DisableRuntimeMarshalling]

namespace CustomPrintDocument;

public partial class App : Application
{
    private Window? m_window;

    public App()
    {
        InitializeComponent();
    }

    protected override void OnLaunched(LaunchActivatedEventArgs args)
    {
        m_window = new MainWindow();
        m_window.Activate();
    }
}
