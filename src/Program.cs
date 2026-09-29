using System;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

namespace OfflineDocxCompare
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            Application application = new Application();
            application.ShutdownMode = ShutdownMode.OnMainWindowClose;
            application.DispatcherUnhandledException += delegate(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs exceptionArgs)
            {
                MessageBox.Show(
                    "程序遇到未处理错误：\n\n" + exceptionArgs.Exception.Message,
                    "离线文档对比",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                exceptionArgs.Handled = true;
            };

            MainWindow window = new MainWindow();
            if (args.Length == 2 && args[0].Equals("--screenshot", StringComparison.OrdinalIgnoreCase))
            {
                application.ShutdownMode = ShutdownMode.OnExplicitShutdown;
                window.Loaded += delegate
                {
                    window.Dispatcher.BeginInvoke(new Action(delegate
                    {
                        SaveWindowPreview(window, args[1]);
                        window.Close();
                        application.Shutdown();
                    }), DispatcherPriority.ApplicationIdle);
                };
            }
            application.Run(window);
        }

        private static void SaveWindowPreview(Window window, string outputPath)
        {
            int width = Math.Max(1, (int)Math.Ceiling(window.ActualWidth));
            int height = Math.Max(1, (int)Math.Ceiling(window.ActualHeight));
            RenderTargetBitmap bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
            bitmap.Render(window);
            PngBitmapEncoder encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(bitmap));

            string fullPath = Path.GetFullPath(outputPath);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath));
            using (FileStream stream = new FileStream(fullPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                encoder.Save(stream);
            }
        }
    }
}
