using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Diagnostics;
using System.Windows.Media;
using System.Threading.Tasks;
using System.Threading;

namespace MarkdownConverter
{
    public static class TaskListProcessor
    {
        public static string PreprocessTaskLists(string markdown)
        {
            if (string.IsNullOrEmpty(markdown)) return markdown;

            markdown = Regex.Replace(markdown, @"^(\s*)-\[\]$", "$1- [ ]", RegexOptions.Multiline);
            markdown = Regex.Replace(markdown, @"^(\s*)-\[(x|X)\]$", "$1- [$2]", RegexOptions.Multiline);
            
            var lines = markdown.Split('\n');
            var result = new StringBuilder();
            
            for (int i = 0; i < lines.Length; i++)
            {
                string currentLine = lines[i].TrimEnd();
                result.AppendLine(currentLine);
                
                if (i < lines.Length - 1)
                {
                    string nextLine = lines[i + 1].TrimEnd();
                    if (NeedsEmptyLine(currentLine, nextLine))
                    {
                        result.AppendLine();
                    }
                }
            }
            
            return result.ToString().TrimEnd();
        }

        private static bool NeedsEmptyLine(string current, string next)
        {
            if (string.IsNullOrWhiteSpace(current) || string.IsNullOrWhiteSpace(next)) return false;

            bool currentIsText = !Regex.IsMatch(current, @"^(\s*[-*]\s+|\s*#+\s+|\s*```|\s*>|\s*\d+\.\s+)");
            bool nextIsList = Regex.IsMatch(next, @"^(\s*[-*]\s+)");
            bool nextIsTaskList = Regex.IsMatch(next, @"^(\s*[-*]\s+\[( |x|X)\])");
            
            if (currentIsText && (nextIsList || nextIsTaskList)) return true;

            bool currentIsList = Regex.IsMatch(current, @"^(\s*[-*]\s+)");
            bool currentIsTaskList = Regex.IsMatch(current, @"^(\s*[-*]\s+\[( |x|X)\])");
            bool nextIsText = !Regex.IsMatch(next, @"^(\s*[-*]\s+|\s*#+\s+|\s*```|\s*>|\s*\d+\.\s+)");

            if ((currentIsList || currentIsTaskList) && nextIsText) return true;

            return false;
        }

        public static string PreprocessForWord(string markdown) => PreprocessTaskLists(markdown);

        public static TaskListStats GetStats(string markdown)
        {
            if (string.IsNullOrEmpty(markdown)) return new TaskListStats();
            
            var taskRegex = new Regex(@"^(\s*[-*]\s+\[( |x|X)\]\s+.*)$", RegexOptions.Multiline);
            var completedRegex = new Regex(@"^(\s*[-*]\s+\[(x|X)\]\s+.*)$", RegexOptions.Multiline);
            
            var total = taskRegex.Matches(markdown).Count;
            var completed = completedRegex.Matches(markdown).Count;
            
            return new TaskListStats
            {
                TotalTasks = total,
                CompletedTasks = completed,
                UncompletedTasks = total - completed,
                CompletionPercentage = total > 0 ? (double)completed / total * 100 : 0
            };
        }

        public static string GetTaskListCss(bool darkTheme = false) => 
            ".task-list{margin:4px 0;padding-left:20px}.task-list li{margin:2px 0}.task-list input[type='checkbox']{margin-right:6px;width:14px;height:14px;cursor:pointer}.task-list li.completed{text-decoration:line-through;opacity:.7}";
    }

    public class TaskListStats
    {
        public int TotalTasks { get; set; }
        public int CompletedTasks { get; set; }
        public int UncompletedTasks { get; set; }
        public double CompletionPercentage { get; set; }

        public override string ToString() => TotalTasks == 0 ? "无任务" : $"任务: {CompletedTasks}/{TotalTasks} ({CompletionPercentage:F1}%)";
    }

    public partial class MainWindow : Window
    {
        private bool _darkTheme = false;
        private string? _workDir;
        private static readonly SemaphoreSlim _semaphore = new(2, 2);
        private static readonly Regex _imgRegex = new(@"!\[(.*?)\]\((.*?)\)", RegexOptions.Compiled);
        private static readonly Regex _tagRegex = new(@"<img\s+[^>]*?src\s*=""([^""]*)""[^>]*>", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public MainWindow() => InitializeComponent();

        private async Task InitWebViewAsync()
        {
            try { await webView2.EnsureCoreWebView2Async(null); await ConvertAsync(); }
            catch (Exception ex) { ShowError($"WebView2初始化失败:\n{ex.Message}"); }
        }

        private async Task<string> ConvertMarkdownToHtmlAsync(string markdown)
        {
            var taskListProcessed = TaskListProcessor.PreprocessTaskLists(markdown);
            var preprocessed = Regex.Replace(
                Regex.Replace(Regex.Replace(taskListProcessed, @"\$\s+([^$]+?)\s+\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline),
                @"\$([^\$]+?)\s+\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline),
                @"\$\s+([^\$]+?)\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline);

            await _semaphore.WaitAsync();
            try
            {
                using var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "pandoc",
                        Arguments = "--mathml -f markdown+tex_math_dollars+table_captions+pipe_tables+grid_tables+raw_html -t html5",
                        RedirectStandardInput = true, RedirectStandardOutput = true, RedirectStandardError = true,
                        UseShellExecute = false, CreateNoWindow = true,
                        StandardErrorEncoding = Encoding.UTF8, StandardOutputEncoding = Encoding.UTF8
                    }
                };
                process.Start();
                using (var writer = new StreamWriter(process.StandardInput.BaseStream, Encoding.UTF8))
                {
                    await writer.WriteAsync(preprocessed);
                    await writer.FlushAsync();
                }
                var outputTask = process.StandardOutput.ReadToEndAsync();
                var errorTask = process.StandardError.ReadToEndAsync();
                if (await Task.WhenAny(outputTask, Task.Delay(30000)) == outputTask)
                {
                    await Task.Run(() => process.WaitForExit());
                    if (process.ExitCode != 0) throw new Exception($"Pandoc转换失败: {await errorTask}");
                    return await outputTask;
                }
                try { process.Kill(); } catch { }
                throw new Exception("Pandoc转换超时（30秒）");
            }
            catch (Exception ex) { ShowError($"Pandoc转换出错: {ex.Message}"); return $"<h1>转换错误</h1><p>{ex.Message}</p>"; }
            finally { _semaphore.Release(); }
        }

        private async Task SaveAsWordAsync(string filePath, string markdown)
        {
            try
            {
                var taskListProcessed = TaskListProcessor.PreprocessForWord(markdown);
                var preprocessed = Regex.Replace(
                    Regex.Replace(Regex.Replace(taskListProcessed, @"\$\s+([^$]+?)\s+\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline),
                    @"\$([^\$]+?)\s+\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline),
                    @"\$\s+([^\$]+?)\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline);

                var processedMarkdown = !string.IsNullOrEmpty(_workDir) ? 
                    _imgRegex.Replace(preprocessed, m =>
                    {
                        var path = m.Groups[2].Value;
                        if (string.IsNullOrEmpty(path) || path.StartsWith("http:") || path.StartsWith("https:") || Path.IsPathRooted(path)) return m.Value;
                        var fullPath = Path.GetFullPath(Path.Combine(_workDir, path));
                        return File.Exists(fullPath) ? $"![{m.Groups[1].Value}]({fullPath})" : m.Value;
                    }) : preprocessed;

                // 使用与WebView2相同的转换流程：先转换为纯HTML，然后添加完整HTML结构
                var htmlContent = await ConvertMarkdownToHtmlAsync(processedMarkdown);
                var fullHtml = GetWebViewHtmlContent(htmlContent);
                
                // 保存为临时HTML文件
                var tempHtmlFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.html");
                await File.WriteAllTextAsync(tempHtmlFile, fullHtml, Encoding.UTF8);

                // 将完整的HTML转换为Word（确保HTML表格被正确转换）
                if (!await RunPandoc(tempHtmlFile, filePath, "-f html -t docx"))
                    throw new Exception("Word转换失败");

                try { File.Delete(tempHtmlFile); } catch { }
                MessageBox.Show("保存成功（Word 文件）！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex) { ShowError($"保存Word文件失败: {ex.Message}"); }
        }

        private async Task<bool> RunPandoc(string inputFile, string outputFile, string extraArgs)
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "pandoc",
                    Arguments = $"\"{inputFile}\" -o \"{outputFile}\" {extraArgs}",
                    RedirectStandardOutput = true, RedirectStandardError = true,
                    UseShellExecute = false, CreateNoWindow = true,
                    StandardErrorEncoding = Encoding.UTF8
                }
            };
            process.Start();
            var errorTask = process.StandardError.ReadToEndAsync();
            if (await Task.WhenAny(errorTask, Task.Delay(30000)) == errorTask)
            {
                await Task.Run(() => process.WaitForExit());
                var error = await errorTask;
                if (process.ExitCode != 0)
                {
                    if (error.Contains("[WARNING]"))
                    {
                        MessageBox.Show($"保存成功（Word 文件）！\n\n注意：{error}", "成功", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return true;
                    }
                    return false;
                }
                // 即使退出码为0，也要检查是否有严重错误
                if (error.Contains("Error") || error.Contains("error:") || error.Contains("fatal:"))
                {
                    return false;
                }
                return true;
            }
            try { process.Kill(); } catch { }
            return false;
        }

        public void LoadMarkdownContent(string content)
        {
            txtMarkdown.Document.Blocks.Clear();
            if (!string.IsNullOrWhiteSpace(content)) txtMarkdown.Document.Blocks.Add(new Paragraph(new Run(content)));
            _ = ConvertAsync();
        }

        public void LoadMarkdownFile(string filePath)
        {
            if (!File.Exists(filePath)) return;
            try { _workDir = Path.GetDirectoryName(filePath); LoadMarkdownContent(File.ReadAllText(filePath, Encoding.UTF8)); }
            catch (Exception ex) { ShowError($"读取文件失败: {ex.Message}"); }
        }

        private async Task ConvertAsync()
        {
            try
            {
                if (webView2.CoreWebView2 == null) await webView2.EnsureCoreWebView2Async(null);
                var rawMarkdown = new TextRange(txtMarkdown.Document.ContentStart, txtMarkdown.Document.ContentEnd).Text;
                if (string.IsNullOrWhiteSpace(rawMarkdown)) { await DisplayHtmlAsync("<html><body></body></html>"); return; }

                var markdownToConvert = !string.IsNullOrEmpty(_workDir) ? PreprocessImagePaths(rawMarkdown) : rawMarkdown;
                var htmlContent = await ConvertMarkdownToHtmlAsync(markdownToConvert);
                var finalHtml = GetWebViewHtmlContent(htmlContent);

                if (!string.IsNullOrEmpty(_workDir) && webView2.CoreWebView2 != null)
                {
                    webView2.CoreWebView2.SetVirtualHostNameToFolderMapping("markdown.local", _workDir, Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.Allow);
                    finalHtml = ReplaceImgSrcInHtml(finalHtml, _workDir);
                }
                await DisplayHtmlAsync(finalHtml);
            }
            catch (Exception ex) { ShowError($"转换错误: {ex.Message}"); }
        }

        private string PreprocessImagePaths(string markdown)
        {
            if (string.IsNullOrEmpty(_workDir)) return markdown;
            return _imgRegex.Replace(markdown, m =>
            {
                var path = m.Groups[2].Value;
                if (string.IsNullOrEmpty(path) || path.StartsWith("http:") || path.StartsWith("https:") || Path.IsPathRooted(path)) return m.Value;
                var fullPath = Path.GetFullPath(Path.Combine(_workDir, path));
                return File.Exists(fullPath) ? $"![{m.Groups[1].Value}](http://markdown.local/{path.Replace("\\", "/")})" : m.Value;
            });
        }

        private string ReplaceImgSrcInHtml(string html, string? workDir)
        {
            if (string.IsNullOrEmpty(workDir) || string.IsNullOrEmpty(html)) return html;
            return _tagRegex.Replace(html, m =>
            {
                var src = m.Groups[1].Value;
                if (string.IsNullOrEmpty(src) || src.StartsWith("http://") || src.StartsWith("https://") || src.StartsWith("http://markdown.local/")) return m.Value;
                return m.Value.Replace(src, $"http://markdown.local/{src.TrimStart('/').Replace("\\", "/")}");
            });
        }

        private string GetWebViewHtmlContent(string content)
        {
            var colors = _darkTheme ? new {
                Text = "#D4D4D4", Heading = "#93C5FD", Border = "#444", Link = "#60A5FA",
                BlockquoteBg = "#2D3748", BlockquoteText = "#A0AEC0", BlockquoteBorder = "#60A5FA",
                CodeBg = "#2D2D30", CodeText = "#E2E8F0", CodeBorder = "#4A5568",
                TableBorder = "#444", TableBg = "#2D2D30", BodyBg = "#1E1E1E"
            } : new {
                Text = "#1E1E1E", Heading = "#2563EB", Border = "#eee", Link = "#2563EB",
                BlockquoteBg = "rgba(37,99,235,0.05)", BlockquoteText = "#555", BlockquoteBorder = "#2563EB",
                CodeBg = "#F8F9FA", CodeText = "#C7254E", CodeBorder = "#E9ECEF",
                TableBorder = "#E0E0E0", TableBg = "#FFFFFF", BodyBg = "#F0F2F5"
            };

            var taskListCss = TaskListProcessor.GetTaskListCss(_darkTheme);
            var themeCss = $@"body{{color:{colors.Text};background-color:{colors.BodyBg}}}
                           h1,h2,h3{{color:{colors.Heading};border-bottom-color:{colors.Border}}}
                           a{{color:{colors.Link}}}
                           blockquote{{background:{colors.BlockquoteBg};color:{colors.BlockquoteText};border-left-color:{colors.BlockquoteBorder}}}
                           pre,code{{background:{colors.CodeBg};color:{colors.CodeText};border-color:{colors.CodeBorder}}}
                           table,th,td{{border-color:{colors.TableBorder};background:{colors.TableBg}}}
                           {taskListCss}";

            return $"<!DOCTYPE html><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><meta charset=\"UTF-8\"><style>html{{color:#1a1a1a;background-color:transparent}}body{{font-family:'Segoe UI',sans-serif;line-height:1.6;padding:20px;max-width:850px;margin:0 auto}}h1{{color:{colors.Heading};border-bottom:2px solid {colors.Border};padding-bottom:10px}}h2{{color:{colors.Heading}}}h3{{color:{colors.Heading}}}a{{color:{colors.Link};text-decoration:none}}a:hover{{text-decoration:underline}}img{{max-width:100%;height:auto;display:block;margin:10px 0;border-radius:4px}}blockquote{{border-left:4px solid {colors.BlockquoteBorder};padding-left:16px;margin-left:0;color:{colors.BlockquoteText};background:{colors.BlockquoteBg}}}table{{border-collapse:collapse;width:100%;margin:16px 0}}th,td{{border:1px solid {colors.TableBorder};padding:10px;text-align:left;background:{colors.TableBg}}}ul,ol{{padding-left:24px}}pre{{white-space:pre-wrap;background:{colors.CodeBg};border:1px solid {colors.CodeBorder};padding:12px;border-radius:5px}}code{{font-family:'Consolas',monospace;background:{colors.CodeBg};padding:2px 4px;border-radius:3px;color:{colors.CodeText}}}{themeCss}math{{font-size:1.1em;font-family:'STIX Two Math',serif}}</style></head><body>{content}</body></html>";
        }

        private async Task DisplayHtmlAsync(string html)
        {
            if (string.IsNullOrEmpty(html)) { webView2.NavigateToString("<html><body></body></html>"); return; }
            if (webView2.CoreWebView2 == null) { await webView2.EnsureCoreWebView2Async(null); if (webView2.CoreWebView2 == null) { ShowError("WebView2初始化失败"); return; } }
            if (html.Length <= 200_000) webView2.NavigateToString(html);
            else
            {
                try
                {
                    var tmpFile = Path.Combine(Path.GetTempPath(), $"mdconv_preview_{Environment.ProcessId}.html");
                    await File.WriteAllTextAsync(tmpFile, html, Encoding.UTF8);
                    webView2.CoreWebView2.Navigate(new Uri(tmpFile).AbsoluteUri);
                }
                catch { webView2.NavigateToString(html); }
            }
        }

        private void ShowError(string message, string title = "错误") => MessageBox.Show(message, title, MessageBoxButton.OK, MessageBoxImage.Error);

        private async void BtnConvert_Click(object sender, RoutedEventArgs e) => await ExecuteButtonAction(sender, ConvertAsync, "转换中...");
        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog { Title = "保存文件", Filter = "Markdown 文件 (*.md)|*.md|HTML 文件 (*.html)|*.html|Word 文件 (*.docx)|*.docx|所有文件 (*.*)|*.*", FileName = "markdown_export.md" };
            if (saveDialog.ShowDialog() != true) return;
            await ExecuteButtonAction(sender, async () =>
            {
                var rawMarkdown = new TextRange(txtMarkdown.Document.ContentStart, txtMarkdown.Document.ContentEnd).Text;
                if (string.IsNullOrWhiteSpace(rawMarkdown)) { MessageBox.Show("没有内容可供保存。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning); return; }
                var ext = Path.GetExtension(saveDialog.FileName)?.ToLowerInvariant();
                if (ext == ".md" || ext == ".markdown") { await File.WriteAllTextAsync(saveDialog.FileName, rawMarkdown, Encoding.UTF8); MessageBox.Show("保存成功（Markdown 文件）！", "成功", MessageBoxButton.OK, MessageBoxImage.Information); return; }
                if (ext == ".docx") { await SaveAsWordAsync(saveDialog.FileName, rawMarkdown); return; }
                var htmlContent = await ConvertMarkdownToHtmlAsync(rawMarkdown);
                var finalHtml = await GetHtmlWithEmbeddedMathML(htmlContent);
                await File.WriteAllTextAsync(saveDialog.FileName, finalHtml, Encoding.UTF8);
                MessageBox.Show("保存成功（HTML 文件）！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }, "保存中...");
        }

        private async Task<string> GetHtmlWithEmbeddedMathML(string markdown)
        {
            try
            {
                var taskListProcessed = TaskListProcessor.PreprocessTaskLists(markdown);
                var preprocessed = Regex.Replace(
                    Regex.Replace(Regex.Replace(taskListProcessed, @"\$\s+([^$]+?)\s+\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline),
                    @"\$([^\$]+?)\s+\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline),
                    @"\$\s+([^\$]+?)\$", m => $"${m.Groups[1].Value}$", RegexOptions.Singleline);

                // 使用与WebView2相同的转换流程
                var htmlContent = await ConvertMarkdownToHtmlAsync(preprocessed);
                return GetWebViewHtmlContent(htmlContent);
            }
            catch (Exception ex) { ShowError($"生成HTML时出错: {ex.Message}"); return $"<!DOCTYPE html><html xmlns=\"http://www.w3.org/1999/xhtml\"><head><meta charset=\"UTF-8\"><title>Error</title></head><body><h1>生成HTML时出错</h1><p>{ex.Message}</p></body></html>"; }
        }

        private async void BtnStyles_Click(object sender, RoutedEventArgs e) => await ExecuteButtonAction(sender, async () => 
        { 
            _darkTheme = !_darkTheme;
            var colors = _darkTheme ? new {
                Main = "#1E1E2E", Editor = "#2A2A3A", Text = "#E6E6E6", Border = "#3A3A4A", Splitter = "#3A3A4A"
            } : new {
                Main = "#F0F2F5", Editor = "#FFFFFF", Text = "#1E1E1E", Border = "#E0E0E0", Splitter = "#E0E0E0"
            };

            this.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Main));
            this.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Text));
            txtMarkdown.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Editor));
            txtMarkdown.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Text));
            txtMarkdown.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Border));

            var gridSplitter = this.FindName("gridSplitter") as GridSplitter;
            if (gridSplitter != null) gridSplitter.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Splitter));

            var contextMenu = txtMarkdown.ContextMenu;
            if (contextMenu != null)
            {
                contextMenu.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Editor));
                contextMenu.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Border));
                foreach (var item in contextMenu.Items)
                {
                    if (item is MenuItem menuItem)
                    {
                        menuItem.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Editor));
                        menuItem.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Text));
                        menuItem.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Border));
                    }
                }
            }

            var webViewBorder = webView2.Parent as Border;
            if (webViewBorder != null)
            {
                webViewBorder.Background = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Editor));
                webViewBorder.BorderBrush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(colors.Border));
            }

            if (txtMarkdown.Document != null && txtMarkdown.Document.Blocks.Count > 0)
            {
                foreach (var block in txtMarkdown.Document.Blocks)
                {
                    if (block is Paragraph paragraph)
                    {
                        foreach (var inline in paragraph.Inlines)
                        {
                            if (inline is Run run)
                            {
                                var isTitle = run.Text.Contains("# Markdown 转换器") || run.Text.Contains("## 使用说明") || 
                                            run.Text.Contains("## 支持的功能") || (run.Text.Contains("#") && run.Text.Length < 20);
                                run.Foreground = new SolidColorBrush((Color)ColorConverter.ConvertFromString(isTitle ? (_darkTheme ? "#93C5FD" : "#2563EB") : colors.Text));
                            }
                        }
                    }
                }
            }
            
            await ConvertAsync(); 
        }, "切换中...");

        private async Task ExecuteButtonAction(object sender, Func<Task> action, string loadingText)
        {
            var btn = sender as Button;
            var originalContent = btn?.Content;
            btn?.Dispatcher.Invoke(() => { btn.Content = loadingText; btn.IsEnabled = false; });
            try { await action(); }
            finally { btn?.Dispatcher.Invoke(() => { btn.Content = originalContent; btn.IsEnabled = true; }); }
        }

        private void PasteMenuItem_Click(object sender, RoutedEventArgs e) => txtMarkdown.Paste();

        private void txtMarkdown_PreviewDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(DataFormats.FileDrop) is string[] files && files.Length > 0 && Path.GetExtension(files[0]).ToLower() is ".md" or ".markdown") LoadMarkdownFile(files[0]);
            e.Handled = true;
        }

        private void txtMarkdown_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }
    }
}
