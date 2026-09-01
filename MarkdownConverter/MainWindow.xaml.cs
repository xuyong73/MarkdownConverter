using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Threading;
using System.Windows.Controls;

namespace MarkdownConverter
{
    public partial class MainWindow : Window
    {
        private bool _darkTheme;
        private bool _scrollToTop;
        private bool _isDirty;
        private bool _isConverting;
        private bool _navigateToCursorAfterRender;
        private bool _pendingAutoConvert;
        private DispatcherTimer? _debounceTimer;
        private string? _workDir;
        private string? _sourceFileName;
        private string? _originalFilePath;
        private CancellationTokenSource? _convertCts;
        private readonly TaskCompletionSource<bool> _webView2InitTcs = new();
        private bool _webView2Initialized;
        private readonly IncrementalConverter _incrementalConverter = new();
        private readonly HashSet<int> _renderedSectionIndices = [];
        private const int ConvertDebounceMs = 300;

        private static readonly string _previewJs;

        static MainWindow()
        {
            var assembly = typeof(MainWindow).Assembly;
            using var previewStream = assembly.GetManifestResourceStream("MarkdownConverter.Preview.js")!;
            using var previewReader = new StreamReader(previewStream);
            _previewJs = previewReader.ReadToEnd();
        }

        private string LastStatusText
        {
            get => field ?? "";
            set
            {
                if (field == value) return;
                field = value;
                txtStatus.Text = value;
            }
        }

        private void SetStatus(string text, string? bgColor = null)
        {
            if (!Dispatcher.CheckAccess())
            {
                _ = Dispatcher.InvokeAsync(() => SetStatus(text, bgColor));
                return;
            }
            LastStatusText = text;
            if (bgColor != null)
                statusBorder.Background = ThemeManager.GetBrush(bgColor);
        }

        public MainWindow()
        {
            InitializeComponent();
            var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
            Title = ver != null ? $"Markdown 转换器 v{ver.Major}.{ver.Minor}" : "Markdown 转换器";
            ApplyThemeColors();
            SetStatus("正在初始化...", ColorConstants.StatusBlue);
            // 后台预热 pandoc 进程，使首次转换跳过进程加载延迟
            WarmupPandocProcess();
            var ctxMenu = CreateEditorContextMenu();
            txtMarkdown.TextChanged += (_, _) => { _isDirty = true; RequestConvert(); };
            txtMarkdown.ContextMenu = ctxMenu;
            txtMarkdown.PreviewMouseRightButtonDown += (s, e) =>
            {
                var point = e.GetPosition(txtMarkdown);
                var pos = txtMarkdown.GetPositionFromPoint(point);
                if (pos.HasValue)
                {
                    int charIndex = txtMarkdown.Document.GetOffset(pos.Value.Location);
                    var selStart = txtMarkdown.SelectionStart;
                    var selLen = txtMarkdown.SelectionLength;
                    if (charIndex < selStart || charIndex > selStart + selLen)
                        txtMarkdown.CaretOffset = charIndex;
                }
            };
            txtMarkdown.PreviewMouseLeftButtonDown += (s, e) =>
            {
                if (e.ClickCount < 2 || !_webView2Initialized) return;
                e.Handled = true;

                // 先将 caret 移到点击位置，因为 e.Handled=true 阻止了默认处理
                var point = e.GetPosition(txtMarkdown);
                var pos = txtMarkdown.GetPositionFromPoint(point);
                if (pos.HasValue)
                    txtMarkdown.CaretOffset = txtMarkdown.Document.GetOffset(pos.Value.Location);

                if (_isConverting && !IsSectionReadyAtCaret())
                {
                    _ = MessageBox.Show("请等待当前转换完成后，再执行定位操作。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                _navigateToCursorAfterRender = true;
                _ = NavigatePreviewToCursorAsync();
            };

            KeyDown += OnMainWindowKeyDown;
        }

        protected override async void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (_isDirty && !string.IsNullOrWhiteSpace(txtMarkdown.Text))
                {
                    var action = await ConfirmDiscardChangesAsync();
                    if (action == DiscardAction.Cancel)
                    {
                        e.Cancel = true;
                        return;
                    }
                    if (action == DiscardAction.SaveAndContinue)
                    {
                        e.Cancel = true;
                        await SaveCurrentFileAsync();
                        if (_isDirty) return;
                        Close();
                    }
                }

                FileService.CleanupTempFiles();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"OnClosing 异常: {ex.Message}");
            }

            base.OnClosing(e);
        }

        private enum DiscardAction { Discard, SaveAndContinue, Cancel }

        private async Task<DiscardAction> ConfirmDiscardChangesAsync()
        {
            if (!_isDirty || string.IsNullOrWhiteSpace(txtMarkdown.Text))
                return DiscardAction.Discard;

            var result = MessageBox.Show(
                $"是否保存对 \"{_sourceFileName ?? "未命名"}\" 的更改？",
                "Markdown 转换器",
                MessageBoxButton.YesNoCancel,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Cancel)
                return DiscardAction.Cancel;

            if (result == MessageBoxResult.Yes)
            {
                await SaveCurrentFileAsync();
                return DiscardAction.SaveAndContinue;
            }

            return DiscardAction.Discard;
        }

        private async Task SaveCurrentFileAsync()
        {
            SetStatus("正在保存...", ColorConstants.StatusAmber);
            var filePath = _originalFilePath;
            if (string.IsNullOrEmpty(filePath))
            {
                var saveDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "保存 Markdown 文件",
                    Filter = "Markdown 文件 (*.md)|*.md|所有文件 (*.*)|*.*",
                    FileName = _sourceFileName ?? "untitled.md"
                };
                if (saveDialog.ShowDialog() != true) return;
                filePath = saveDialog.FileName;
            }

            await File.WriteAllTextAsync(filePath, txtMarkdown.Text, Encoding.UTF8);
            _originalFilePath = filePath;
            _sourceFileName = Path.GetFileName(filePath);
            _workDir = Path.GetDirectoryName(filePath);
            _isDirty = false;
            SetStatus("保存成功", ColorConstants.StatusGreen);
        }

        private static CancellationTokenSource ReplaceCts(ref CancellationTokenSource? field)
        {
            var old = Interlocked.Exchange(ref field, new CancellationTokenSource());
            old?.Cancel();
            old?.Dispose();
            return field;
        }

        private static async Task ExecuteButtonAction(object sender, Func<Task> action, string loadingText)
        {
            var btn = sender as Button;
            var originalContent = btn?.Content;
            btn?.Dispatcher.Invoke(() => { btn.Content = loadingText; btn.IsEnabled = false; });
            try { await action(); }
            finally { btn?.Dispatcher.Invoke(() => { btn.Content = originalContent; btn.IsEnabled = true; }); }
        }

        /// <summary>后台预热 pandoc 进程：跑一次 --version 让 OS 缓存二进制</summary>
        private static void WarmupPandocProcess()
        {
            Task.Run(() =>
            {
                try
                {
                    using var proc = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = "pandoc",
                            Arguments = "--version",
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    if (proc.Start())
                        proc.WaitForExit(2000);
                }
                catch
                {
                    // pandoc 未安装或不在 PATH，首次转换时会提示用户
                }
            });
        }

        /// <summary>当前光标所在段是否已转换完成（不受转换中状态影响）</summary>
        private bool IsSectionReadyAtCaret() =>
            !_isConverting || _renderedSectionIndices.Contains(GetSectionIndexAtCaret());
    }
}

/// <summary>System.Text.Json 源生成上下文，避免运行时反射</summary>
[System.Text.Json.Serialization.JsonSerializable(typeof(string))]
internal partial class AppJsonContext : System.Text.Json.Serialization.JsonSerializerContext
{
}
