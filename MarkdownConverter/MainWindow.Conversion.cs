using System.IO;
using System.Windows;
using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;

namespace MarkdownConverter
{
    /// <summary>MainWindow 的转换逻辑：增量/全量转换、WebView2 初始化、预览渲染</summary>
    public partial class MainWindow
    {
        public async Task InitWebViewAndConvertAsync()
        {
            await InitWebViewAsync();
            RequestConvert();
        }

        private async Task InitWebViewAsync()
        {
            if (_webView2Initialized) return;

            try
            {
                SetStatus("正在初始化 WebView2...", ColorConstants.StatusAmber);
                var userDataFolder = Path.Combine(Path.GetTempPath(), $"MDConv_WebView2_{Environment.ProcessId}");
                if (!Directory.Exists(userDataFolder)) Directory.CreateDirectory(userDataFolder);

                var environmentOptions = new CoreWebView2EnvironmentOptions("", userDataFolder);
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, environmentOptions);
                await webView2.EnsureCoreWebView2Async(env);

                webView2.DefaultBackgroundColor = System.Drawing.Color.FromArgb(255, 255, 255);

                webView2.CoreWebView2.WebMessageReceived += OnPreviewWebMessageReceived;

                webView2.CoreWebView2.Settings.IsWebMessageEnabled = true;
                webView2.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = true;

                webView2.CoreWebView2.ContextMenuRequested += (s, e) =>
                {
                    e.Handled = true;

                    Dispatcher.Invoke(() =>
                    {
                        var menu = new System.Windows.Controls.ContextMenu { PlacementTarget = webView2, Placement = System.Windows.Controls.Primitives.PlacementMode.Mouse };

                        var copyItem = new System.Windows.Controls.MenuItem { Header = "复制", InputGestureText = "Ctrl+C" };
                        copyItem.Click += (_, _) => _ = webView2.CoreWebView2.ExecuteScriptAsync("document.execCommand('copy');");
                        menu.Items.Add(copyItem);

                        var printItem = new System.Windows.Controls.MenuItem { Header = "打印", InputGestureText = "Ctrl+P" };
                        printItem.Click += (_, _) => _ = webView2.CoreWebView2.ExecuteScriptAsync("window.print();");
                        menu.Items.Add(printItem);

                        menu.Items.Add(new System.Windows.Controls.Separator());

                        var searchItem = new System.Windows.Controls.MenuItem { Header = "搜索(_F)", InputGestureText = "Ctrl+F" };
                        searchItem.Click += async (_, _) =>
                        {
                            webView2.Focus();
                            await System.Threading.Tasks.Task.Delay(30);
                            // 用 keybd_event 模拟 Ctrl+F，触发 WebView2 原生查找栏
                            NativeMethods.keybd_event(NativeMethods.VK_CONTROL, 0, 0, System.IntPtr.Zero);
                            NativeMethods.keybd_event(NativeMethods.VK_F, 0, 0, System.IntPtr.Zero);
                            NativeMethods.keybd_event(NativeMethods.VK_F, 0, NativeMethods.KEYEVENTF_KEYUP, System.IntPtr.Zero);
                            NativeMethods.keybd_event(NativeMethods.VK_CONTROL, 0, NativeMethods.KEYEVENTF_KEYUP, System.IntPtr.Zero);
                        };
                        menu.Items.Add(searchItem);

                        menu.IsOpen = true;
                    });
                };

                await webView2.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync(_previewJs);

                await webView2.CoreWebView2.ExecuteScriptAsync(ThemeManager.GetThemeJsCss(_darkTheme));

                _webView2Initialized = true;
                _webView2InitTcs.TrySetResult(true);
                SetStatus("就绪", ColorConstants.StatusGreen);
            }
            catch (Exception ex)
            {
                _webView2InitTcs.TrySetException(ex);
                SetStatus("WebView2 初始化失败", ColorConstants.StatusRed);
                MessageBox.Show($"WebView2 初始化失败:\n{ex.Message}\n\n请确保已安装 Microsoft Edge WebView2 Runtime\n\n下载地址：https://go.microsoft.com/fwlink/?linkid=2124701", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RequestConvert()
        {
            SetStatus("等待输入...", ColorConstants.StatusAmber);

            var textLen = txtMarkdown.Text.Length;
            var debounceMs = textLen switch
            {
                > 500000 => 800,
                > 200000 => 600,
                > 50000  => 400,
                _        => ConvertDebounceMs
            };

            if (_debounceTimer == null)
            {
                _debounceTimer = new DispatcherTimer();
                _debounceTimer.Tick += OnDebounceTick;
            }
            _debounceTimer.Interval = TimeSpan.FromMilliseconds(debounceMs);
            _debounceTimer.Stop();
            _debounceTimer.Start();
        }

        private void OnDebounceTick(object? sender, EventArgs e)
        {
            _debounceTimer!.Stop();

            if (_isConverting)
            {
                _pendingAutoConvert = true;
            }
            else
            {
                var ct = ReplaceCts(ref _convertCts).Token;
                _ = ConvertAsync(ct);
            }
        }

        /// <summary>跳过防抖直接启动转换（用于初始加载等无需等待的场景）</summary>
        private async Task StartConvertNowAsync()
        {
            if (_isConverting)
            {
                _pendingAutoConvert = true;
                return;
            }
            _pendingAutoConvert = false;
            var ct = ReplaceCts(ref _convertCts).Token;
            await ConvertAsync(ct);
        }

        private async Task ConvertAsync(CancellationToken cancellationToken = default)
        {
            _isConverting = true;
            _lineMapCache.Clear();
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                SetStatus("正在转换，请稍后...", ColorConstants.StatusAmber);
                if (webView2.CoreWebView2 == null)
                {
                    try { await _webView2InitTcs.Task; } catch { return; }
                    cancellationToken.ThrowIfCancellationRequested();
                    if (webView2.CoreWebView2 == null) return;
                }
                var rawMarkdown = txtMarkdown.Text;
                if (string.IsNullOrWhiteSpace(rawMarkdown)) { SetStatus("就绪", ColorConstants.StatusGreen); await DisplayHtmlAsync("<html><body></body></html>"); return; }

                cancellationToken.ThrowIfCancellationRequested();

                var (sections, changedIndices) = await Task.Run(() => _incrementalConverter.Analyze(rawMarkdown), cancellationToken);

                bool shouldNavigate = _navigateToCursorAfterRender;
                if (changedIndices is null)
                {
                    var reuseMap = _incrementalConverter.TryMatchAfterReshuffle(sections);
                    await FullConvertAsync(sections, cancellationToken, reuseMap);
                }
                else if (changedIndices is [])
                {
                    if (shouldNavigate)
                    {
                        _navigateToCursorAfterRender = false;
                        await NavigatePreviewToCursorAsync();
                    }
                    else
                    {
                        SetStatus("内容无变化，跳过转换", ColorConstants.StatusBlue);
                    }
                    return;
                }
                else
                {
                    await IncrementalConvertAsync(sections, changedIndices, cancellationToken);
                }

                if (!shouldNavigate)
                    SetStatus("就绪", ColorConstants.StatusGreen);
            }
            catch (OperationCanceledException) { SetStatus("转换已取消", ColorConstants.StatusAmber); }
            catch (Exception ex) { SetStatus("转换出错", ColorConstants.StatusRed); MessageBox.Show($"转换错误：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error); }
            finally
            {
                _isConverting = false;
                if (_pendingAutoConvert)
                {
                    _pendingAutoConvert = false;
                    RequestConvert();
                }
            }
        }

        private async Task FullConvertAsync(List<string> sections, CancellationToken ct, List<int>? reuseMap = null)
        {
            double? restoreScrollY = null;
            if (!_navigateToCursorAfterRender && !_scrollToTop)
                restoreScrollY = await SaveScrollPositionAsync();

            var workDir = _workDir;
            var htmlParts = new string[sections.Count];
            var cachedHtml = reuseMap != null ? _incrementalConverter.GetCachedHtml() : null;

            // 已有内容时复用当前 DOM，不重建 shell，避免预览闪烁
            bool hasExistingContent = _renderedSectionIndices.Count > 0;
            _renderedSectionIndices.Clear();

            // 尽早启动段转换，并行执行
            List<Task<(int idx, string html)>> tasks = [.. sections.Select((s, i) => ConvertOrReuseAsync(i, s))];

            Task? renderShellTask = null;
            if (!hasExistingContent)
                renderShellTask = RenderShellAsync(sections.Count, ct, workDir);
            else
                await AdjustSectionCountAsync(sections.Count);

            bool shellReady = hasExistingContent;
            await foreach (var task in Task.WhenEach(tasks))
            {
                var (idx, html) = await task;
                htmlParts[idx] = html;
                if (!shellReady)
                {
                    await renderShellTask!;
                    shellReady = true;
                }
                await FillPreviewSectionAsync(idx, html, workDir);
            }
            if (!shellReady)
                await renderShellTask!;

            _incrementalConverter.CacheFullResult(sections, [.. htmlParts]);

            if (_navigateToCursorAfterRender)
            {
                _navigateToCursorAfterRender = false;
                await NavigatePreviewToCursorAsync();
            }
            else if (restoreScrollY.HasValue)
            {
                await webView2.CoreWebView2.ExecuteScriptAsync($"window.scrollTo(0, {restoreScrollY.Value});");
            }
            _scrollToTop = false;

            async Task<(int Index, string Html)> ConvertOrReuseAsync(int idx, string section)
            {
                if (reuseMap != null && reuseMap[idx] >= 0)
                    return (idx, cachedHtml![reuseMap[idx]]);
                var html = await Task.Run(() => PandocConverter.ConvertToHtmlAsync(section, ct), ct).ConfigureAwait(false);
                return (idx, html);
            }
        }

        private async Task IncrementalConvertAsync(List<string> sections, List<int> changedIndices, CancellationToken ct)
        {
            var workDir = _workDir;

            var cachedHtml = _incrementalConverter.GetCachedHtml();
            if (cachedHtml == null || cachedHtml.Count != sections.Count)
            {
                _incrementalConverter.Clear();
                await FullConvertAsync(sections, ct);
                return;
            }

            _renderedSectionIndices.Clear();
            for (int i = 0; i < sections.Count; i++)
                _renderedSectionIndices.Add(i);
            foreach (var idx in changedIndices)
                _renderedSectionIndices.Remove(idx);

            List<Task<(int idx, string html)>> tasks = [.. changedIndices.Select(idx => ConvertAndCacheAsync(idx, sections[idx]))];

            await foreach (var task in Task.WhenEach(tasks))
            {
                var (idx, html) = await task;
                await FillPreviewSectionAsync(idx, html, workDir);
            }

            if (_navigateToCursorAfterRender)
            {
                _navigateToCursorAfterRender = false;
                _scrollToTop = false;
                await NavigatePreviewToCursorAsync();
            }
            _scrollToTop = false;

            async Task<(int Index, string Html)> ConvertAndCacheAsync(int idx, string section)
            {
                var html = await Task.Run(() => PandocConverter.ConvertToHtmlAsync(section, ct), ct).ConfigureAwait(false);
                cachedHtml[idx] = html;
                return (idx, html);
            }
        }



        private ulong _lastShellNavigationId;

        private async Task RenderShellAsync(int sectionCount, CancellationToken ct, string? workDir = null)
        {
            _renderedSectionIndices.Clear();
            workDir ??= _workDir;
            var darkTheme = _darkTheme;
            var shellHtml = await Task.Run(() => HtmlGenerator.GenerateShellHtml(sectionCount, darkTheme), ct);

            if (!string.IsNullOrEmpty(workDir))
                shellHtml = await Task.Run(() => ImagePathProcessor.ProcessHtmlImagePaths(shellHtml, workDir), ct);

            var tempFile = await FileService.WriteTempHtmlAsync(shellHtml);

            await Dispatcher.InvokeAsync(() =>
            {
                if (webView2.CoreWebView2 is not { } navWv) return;

                if (!string.IsNullOrEmpty(workDir))
                {
                    navWv.ClearVirtualHostNameToFolderMapping("markdown.local");
                    navWv.SetVirtualHostNameToFolderMapping("markdown.local", workDir, CoreWebView2HostResourceAccessKind.Allow);
                }

                void OnNavigating(object? s, CoreWebView2NavigationStartingEventArgs e)
                {
                    _lastShellNavigationId = e.NavigationId;
                }

                void OnNavigated(object? s, CoreWebView2NavigationCompletedEventArgs e)
                {
                    navWv.NavigationCompleted -= OnNavigated;
                    navWv.NavigationStarting -= OnNavigating;
                    // 仅当导航 ID 匹配时才清理，避免错误删除新导航的临时文件
                    if (e.NavigationId == _lastShellNavigationId)
                        FileService.TryDeleteFile(tempFile);
                }

                navWv.NavigationStarting += OnNavigating;
                navWv.NavigationCompleted += OnNavigated;

                navWv.Navigate(new Uri(tempFile).AbsoluteUri);
            });
        }

        private async Task FillPreviewSectionAsync(int idx, string htmlContent, string? workDir = null)
        {
            workDir ??= _workDir;
            if (webView2.CoreWebView2 is not { } wv) return;

            htmlContent = await Task.Run(() =>
            {
                var html = !string.IsNullOrEmpty(workDir)
                    ? ImagePathProcessor.ProcessHtmlImagePaths(htmlContent, workDir)
                    : htmlContent;
                html = HtmlGenerator.WrapTables(html);
                html = HtmlGenerator.WrapLineSpans(html);
                html = HtmlGenerator.AddBlockOrdinals(html);
                return html;
            });

            var escaped = await Task.Run(() => System.Text.Json.JsonSerializer.Serialize(htmlContent, AppJsonContext.Default.String));
            await wv.ExecuteScriptAsync(
                $"document.querySelector('[data-section-idx=\"{idx}\"]').innerHTML={escaped};");

            _renderedSectionIndices.Add(idx);
        }

        /// <summary>在已有预览 DOM 中调整段容器数量（段数因标题增减而变化时）</summary>
        private async Task AdjustSectionCountAsync(int targetCount)
        {
            if (webView2.CoreWebView2 is not { } wv) return;
            await wv.ExecuteScriptAsync($@"
(function(count){{
    var existing=document.querySelectorAll('[data-section-idx]');
    while(existing.length>count) existing[existing.length-1].remove(),existing=document.querySelectorAll('[data-section-idx]');
    for(var i=existing.length;i<count;i++){{
        var d=document.createElement('div');
        d.className='md-section';
        d.setAttribute('data-section-idx',i);
        document.body.appendChild(d);
    }}
}})({targetCount});");
        }

        private async Task<double?> SaveScrollPositionAsync()
        {
            if (webView2.CoreWebView2 is not { } wv) return null;
            var result = await wv.ExecuteScriptAsync("window.scrollY");
            return double.TryParse(result, out var y) && y > 0 ? y : null;
        }

        private async Task DisplayHtmlAsync(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                var emptyBg = ThemeManager.GetBackgroundColor(_darkTheme);
                webView2.CoreWebView2.NavigateToString($"<html><body style='background-color:{emptyBg}'></body></html>");
                return;
            }
            if (webView2.CoreWebView2 == null) return;

            var tempFile = await FileService.WriteTempHtmlAsync(html);

            void OnNavigated(object? s, CoreWebView2NavigationCompletedEventArgs e)
            {
                webView2.CoreWebView2.NavigationCompleted -= OnNavigated;
                if (e.NavigationId == _lastShellNavigationId)
                    FileService.TryDeleteFile(tempFile);
            }

            webView2.CoreWebView2.NavigationCompleted += OnNavigated;
            webView2.CoreWebView2.Navigate(new Uri(tempFile).AbsoluteUri);
        }
    }
}
