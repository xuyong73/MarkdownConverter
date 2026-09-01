using System.IO;
using System.Windows;

namespace MarkdownConverter
{
    public partial class App : Application
    {
        protected override async void OnStartup(StartupEventArgs e)
        {
            try
            {
                base.OnStartup(e);

                // 清理上次异常退出残留的临时文件
                FileService.CleanupTempFiles();

                var main = new MainWindow();
                main.Show();

                await main.InitWebViewAndConvertAsync();

                if (e.Args.Length > 0)
                {
                    var p = e.Args[0];
                    if (File.Exists(p) && FileService.IsMarkdownFile(p))
                        await main.LoadMarkdownFileAsync(p);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnStartup 异常: {ex}");
                MessageBox.Show($"启动异常: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            FileService.CleanupTempFiles();
            base.OnExit(e);
        }
    }
}
