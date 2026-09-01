using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace MarkdownConverter
{
    /// <summary>MainWindow 的编辑器操作：右键菜单、快捷键、文件操作、主题切换</summary>
    public partial class MainWindow
    {
        private void OnMainWindowKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyboardDevice.Modifiers != ModifierKeys.Control) return;

            switch (e.Key)
            {
                case Key.S:
                    e.Handled = true;
                    _ = SaveCurrentFileAsync();
                    break;
                case Key.P:
                    e.Handled = true;
                    _ = CustomPasteAsync();
                    break;
                case Key.W:
                    e.Handled = true;
                    Dispatcher.InvokeAsync(ToggleWordWrap);
                    break;
                case Key.Q:
                    e.Handled = true;
                    _navigateToCursorAfterRender = true;
                    if (!IsSectionReadyAtCaret())
                        _ = Dispatcher.InvokeAsync(() => MessageBox.Show("请等待当前转换完成后，再执行定位操作。", "提示", MessageBoxButton.OK, MessageBoxImage.Information));
                    else
                        _ = NavigatePreviewToCursorAsync();
                    break;
                case Key.C:
                    if (Keyboard.FocusedElement == webView2)
                    {
                        e.Handled = true;
                        _ = webView2.CoreWebView2.ExecuteScriptAsync("document.execCommand('copy');");
                    }
                    break;
                case Key.F:
                    // WebView2 原生支持 Ctrl+F，无需拦截模拟
                    break;
            }
        }

        private async Task CustomPasteAsync()
        {
            if (_isDirty)
            {
                var action = await ConfirmDiscardChangesAsync();
                if (action == DiscardAction.Cancel) return;
            }
            txtMarkdown.Clear();
            txtMarkdown.Paste();
            _originalFilePath = null;
            _sourceFileName = null;
            _workDir = null;
        }

        private ContextMenu CreateEditorContextMenu()
        {
            var menu = new ContextMenu();

            var navItem = new MenuItem { Header = "定位到预览(_N)", InputGestureText = "Ctrl+Q" };
            navItem.Click += (_, _) => _ = NavigatePreviewToCursorAsync();
            menu.Items.Add(navItem);

            var importPasteItem = new MenuItem { Header = "导入粘贴(_P)", InputGestureText = "Ctrl+P" };
            importPasteItem.Click += async (_, _) => await CustomPasteAsync();
            menu.Items.Add(importPasteItem);

            var wrapItem = new MenuItem { Header = "自动换行(_W)", InputGestureText = "Ctrl+W", IsCheckable = true, IsChecked = txtMarkdown.WordWrap };
            wrapItem.Click += (_, _) => ToggleWordWrap();
            menu.Items.Add(wrapItem);

            menu.Items.Add(new Separator());

            var undoItem = new MenuItem { Header = "撤销(_U)", InputGestureText = "Ctrl+Z" };
            undoItem.Click += (_, _) => txtMarkdown.Undo();
            menu.Items.Add(undoItem);

            menu.Items.Add(new Separator());

            var cutItem = new MenuItem { Header = "剪切(_T)", InputGestureText = "Ctrl+X" };
            cutItem.Click += (_, _) => txtMarkdown.Cut();
            menu.Items.Add(cutItem);

            var copyItem = new MenuItem { Header = "复制(_C)", InputGestureText = "Ctrl+C" };
            copyItem.Click += (_, _) => txtMarkdown.Copy();
            menu.Items.Add(copyItem);

            var systemPasteItem = new MenuItem { Header = "粘贴(_V)", InputGestureText = "Ctrl+V" };
            systemPasteItem.Click += (_, _) => txtMarkdown.Paste();
            menu.Items.Add(systemPasteItem);

            menu.Items.Add(new Separator());

            var selectAllItem = new MenuItem { Header = "全选 (_A)", InputGestureText = "Ctrl+A" };
            selectAllItem.Click += (_, _) => txtMarkdown.SelectAll();
            menu.Items.Add(selectAllItem);

            menu.Opened += (_, _) =>
            {
                bool canNavigate = _webView2Initialized && !string.IsNullOrEmpty(txtMarkdown.Text)
                    && IsSectionReadyAtCaret();
                navItem.IsEnabled = canNavigate;
                importPasteItem.IsEnabled = Clipboard.ContainsText();
                wrapItem.IsChecked = txtMarkdown.WordWrap;
                undoItem.IsEnabled = txtMarkdown.CanUndo;
                cutItem.IsEnabled = !string.IsNullOrEmpty(txtMarkdown.SelectedText);
                copyItem.IsEnabled = !string.IsNullOrEmpty(txtMarkdown.SelectedText);
                systemPasteItem.IsEnabled = Clipboard.ContainsText();
                selectAllItem.IsEnabled = true;
            };

            return menu;
        }

        private async void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (await ConfirmDiscardChangesAsync() == DiscardAction.Cancel) return;
                var openDialog = new Microsoft.Win32.OpenFileDialog { Title = "打开 Markdown 文件", Filter = "Markdown 文件 (*.md;*.markdown)|*.md;*.markdown|所有文件 (*.*)|*.*" };
                if (openDialog.ShowDialog() == true)
                    await LoadMarkdownFileAsync(openDialog.FileName);
            }
            catch (Exception ex) { SetStatus("打开文件失败", ColorConstants.StatusRed); MessageBox.Show($"打开文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var defaultName = _sourceFileName ?? "markdown_export.md";
                var saveDialog = new Microsoft.Win32.SaveFileDialog { Title = "保存文件", Filter = "Markdown 文件 (*.md)|*.md|HTML 文件 (*.html)|*.html|Word 文件 (*.docx)|*.docx|所有文件 (*.*)|*.*", FileName = defaultName };
                if (saveDialog.ShowDialog() != true) return;
                await ExecuteButtonAction(sender, async () =>
                {
                    SetStatus("正在保存...", ColorConstants.StatusAmber);
                    var rawMarkdown = txtMarkdown.Text;
                    if (string.IsNullOrWhiteSpace(rawMarkdown)) { MessageBox.Show("没有内容可供保存。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                    switch (Path.GetExtension(saveDialog.FileName)?.ToLowerInvariant())
                    {
                        case ".md" or ".markdown":
                            await FileService.SaveMarkdownAsync(saveDialog.FileName, rawMarkdown);
                            _originalFilePath = saveDialog.FileName;
                            _sourceFileName = Path.GetFileName(saveDialog.FileName);
                            _workDir = Path.GetDirectoryName(saveDialog.FileName);
                            _isDirty = false;
                            SetStatus("保存成功（Markdown）", ColorConstants.StatusGreen);
                            MessageBox.Show("保存成功（Markdown 文件）！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                            break;
                        case ".docx":
                            await SaveAsWordAsync(saveDialog.FileName, rawMarkdown);
                            break;
                        default:
                            await FileService.SaveHtmlAsync(saveDialog.FileName, rawMarkdown, _darkTheme);
                            SetStatus("保存成功（HTML）", ColorConstants.StatusGreen);
                            MessageBox.Show("保存成功（HTML 文件）！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                            break;
                    }
                }, "保存中...");
            }
            catch (Exception ex) { SetStatus("导出失败", ColorConstants.StatusRed); MessageBox.Show($"导出时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private async void BtnStyles_Click(object sender, RoutedEventArgs e)
        {
            try { await ExecuteButtonAction(sender, async () =>
            {
                _darkTheme = !_darkTheme;
                ApplyThemeColors();
                webView2.DefaultBackgroundColor = _darkTheme
                    ? System.Drawing.Color.FromArgb(255, 0, 0, 0)
                    : System.Drawing.Color.FromArgb(255, 255, 255, 255);
                if (webView2.CoreWebView2 is { } wv)
                    await wv.ExecuteScriptAsync(ThemeManager.GetThemeJsCss(_darkTheme));
                SetStatus(_darkTheme ? "深色主题" : "亮色主题", ColorConstants.StatusBlue);
            }, "切换中..."); }
            catch (Exception) { SetStatus("切换主题失败", ColorConstants.StatusRed); }
        }

        private void ToggleWordWrap()
        {
            var newValue = !txtMarkdown.WordWrap;
            txtMarkdown.WordWrap = newValue;
            _ = Dispatcher.InvokeAsync(() =>
                SetStatus(newValue ? "已启用自动换行 (Ctrl+W)" : "已关闭自动换行 (Ctrl+W)", ColorConstants.StatusBlue));
        }

        private void ApplyThemeColors()
        {
            var colors = ThemeManager.GetColors(_darkTheme);

            ApplyControlColor(this, colors.Main, colors.Text);
            ApplyControlColor(txtMarkdown, colors.Editor, colors.Text, colors.Border);

            if (gridSplitter != null) gridSplitter.Background = ThemeManager.GetBrush(colors.Splitter);

            if (webView2.Parent is Border webViewBorder)
                ApplyControlColor(webViewBorder, colors.Editor, borderBrush: colors.Border);
        }

        private static void ApplyControlColor(Control control, string? background = null, string? foreground = null, string? borderBrush = null)
        {
            if (background != null) control.Background = ThemeManager.GetBrush(background);
            if (foreground != null) control.Foreground = ThemeManager.GetBrush(foreground);
            if (borderBrush != null) control.BorderBrush = ThemeManager.GetBrush(borderBrush);
        }

        private static void ApplyControlColor(Border border, string? background = null, string? borderBrush = null)
        {
            if (background != null) border.Background = ThemeManager.GetBrush(background);
            if (borderBrush != null) border.BorderBrush = ThemeManager.GetBrush(borderBrush);
        }

        private async void TxtMarkdown_PreviewDrop(object sender, DragEventArgs e)
        {
            try
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0)
                    {
                        if (FileService.IsMarkdownFile(files[0]))
                        {
                            if (await ConfirmDiscardChangesAsync() == DiscardAction.Cancel) { e.Handled = true; return; }
                            await LoadMarkdownFileAsync(files[0]);
                        }
                    }
                    e.Handled = true;
                }
            }
            catch (Exception) { SetStatus("拖放文件失败", ColorConstants.StatusRed); }
        }

        private void TxtMarkdown_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.Move;
            }
        }

        public async Task LoadMarkdownFileAsync(string filePath)
        {
            if (!File.Exists(filePath)) return;
            try
            {
                SetStatus($"正在打开文件...", ColorConstants.StatusBlue);
                var (text, workDir, fileName) = await FileService.LoadFileAsync(filePath);
                _workDir = workDir;
                _sourceFileName = fileName;
                _originalFilePath = filePath;
                var ver = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version;
                Title = ver != null
                    ? $"Markdown 转换器 v{ver.Major}.{ver.Minor} - " + Path.GetFileNameWithoutExtension(fileName)
                    : "Markdown 转换器 - " + Path.GetFileNameWithoutExtension(fileName);
                _incrementalConverter.Clear();
                _navigateToCursorAfterRender = false;
                txtMarkdown.Text = text;
                txtMarkdown.CaretOffset = 0;
                txtMarkdown.SelectionLength = 0;
                txtMarkdown.ScrollToHome();
                _isDirty = false;
                _scrollToTop = true;
                // 取消正在进行的转换，避免旧转换干扰新文件预览
                _convertCts?.Cancel();
                _debounceTimer?.Stop();

                // 等待前一次转换完全退出（已取消，应很快完成）
                for (int i = 0; i < 100 && _isConverting; i++)
                    await Task.Delay(5);

                _convertCts?.Dispose();
                _pendingAutoConvert = false;
                _convertCts = new CancellationTokenSource();
                _renderedSectionIndices.Clear();

                // 跳过防抖：初始加载直接转换，减少等待时间
                _ = StartConvertNowAsync();
            }
            catch (Exception ex) { SetStatus("打开文件失败", ColorConstants.StatusRed); MessageBox.Show($"打开文件时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
        }

        private async Task SaveAsWordAsync(string filePath, string markdown)
        {
            try
            {
                SetStatus("正在导出 Word...", ColorConstants.StatusAmber);
                if (!await FileService.ExportWordAsync(filePath, markdown, _workDir, _darkTheme))
                    throw new InvalidOperationException("Word转换失败");
                SetStatus("Word 导出成功", ColorConstants.StatusGreen);
                MessageBox.Show("保存成功（Word 文件）！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) { SetStatus("Word 导出失败", ColorConstants.StatusRed); MessageBox.Show($"保存 Word 文件失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
        }
    }
}
