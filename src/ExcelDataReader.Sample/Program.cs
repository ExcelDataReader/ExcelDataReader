using System.Text;
using Avalonia;

namespace ExcelDataReader.Sample;

internal static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    public static AppBuilder BuildAvaloniaApp()
    {
        return AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .LogToTrace();
    }
}
