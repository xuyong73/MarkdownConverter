using System.IO;
using System.Windows;

namespace MarkdownConverter
{
    public partial class App : Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            base.OnStartup(e);
            StartupLogic(e);
        }

        private void StartupLogic(StartupEventArgs e)
        {
            var main = new MainWindow();
            if (e.Args.Length > 0)
            {
                var p = e.Args[0];
                if (File.Exists(p) && (Path.GetExtension(p).Equals(".md", System.StringComparison.OrdinalIgnoreCase) || Path.GetExtension(p).Equals(".markdown", System.StringComparison.OrdinalIgnoreCase)))
                    main.LoadMarkdownFile(p);
            }
            main.Show();
        }
    }
}