using System.Diagnostics;
using System.IO;
using System.Text;

namespace MarkdownConverter
{
    public static class PandocConverter
    {
        private const string PandocNotAvailableMessage =
            "未找到 pandoc，请安装 pandoc 并将其添加到 PATH 环境变量。\n下载地址: https://pandoc.org/installing.html";
        private static readonly SemaphoreSlim _semaphore = new(Math.Max(2, Environment.ProcessorCount / 2), Math.Max(2, Environment.ProcessorCount / 2));
        private const int TimeoutMs = 30000;

        /// <summary>Markdown → HTML 的 pandoc 参数</summary>
        private static readonly string[] _mdToHtmlArgs =
        [
            "--mathml",
            "-f", "markdown+tex_math_dollars+tex_math_single_backslash+table_captions+pipe_tables+grid_tables+raw_html+hard_line_breaks-subscript-raw_tex",
            "-t", "html5",
        ];

        private static readonly Lock _pandocCheckLock = new();
        private static bool? _pandocAvailable;
        private static int _pandocCheckFailCount;
        private static DateTime _lastPandocCheckFailTime = DateTime.MinValue;

        /// <summary>检查 pandoc 是否可用，若不可用则抛出清晰的引导信息</summary>
        private static void EnsurePandocAvailable()
        {
            // 快速路径：无锁读取已缓存的有效状态（lock 外读取，可能读到旧值但不影响正确性）
            if (_pandocAvailable == true)
                return;

            lock (_pandocCheckLock)
            {
                // 双重检查锁定
                if (_pandocAvailable == true)
                    return;

                if (_pandocAvailable == false)
                {
                    // 连续失败 3 次后，给用户 30 秒缓冲期重新检测
                    if (_pandocCheckFailCount >= 3 && (DateTime.UtcNow - _lastPandocCheckFailTime).TotalSeconds < 30)
                        throw new InvalidOperationException(PandocNotAvailableMessage);
                    // 超过 30 秒或未达上限，重置并重新检测
                    _pandocCheckFailCount = 0;
                    _pandocAvailable = null;
                }

                try
                {
                    using var proc = new Process
                    {
                        StartInfo = new ProcessStartInfo
                        {
                            FileName = "pandoc",
                            Arguments = "--version",
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        }
                    };
                    proc.Start();
                    proc.WaitForExit(5000);
                    _pandocAvailable = proc.ExitCode == 0;
                }
                catch
                {
                    _pandocAvailable = false;
                }

                if (_pandocAvailable == false)
                {
                    _pandocCheckFailCount++;
                    _lastPandocCheckFailTime = DateTime.UtcNow;
                    throw new InvalidOperationException(PandocNotAvailableMessage);
                }

                _pandocCheckFailCount = 0;
            }
        }

        public static Task<string> ConvertToHtmlAsync(string markdown) =>
            ConvertToHtmlAsync(markdown, CancellationToken.None);

        public static async Task<string> ConvertToHtmlAsync(string markdown, CancellationToken cancellationToken)
        {
            EnsurePandocAvailable();
            await _semaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                using var process = CreateAndStartProcess(_mdToHtmlArgs);

                var sanitized = SanitizeMathBlocks(markdown);
                await process.StandardInput.WriteAsync(sanitized.AsMemory(), cancellationToken).ConfigureAwait(false);
                await process.StandardInput.FlushAsync(cancellationToken).ConfigureAwait(false);
                process.StandardInput.Close();

                var outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
                var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);

                try
                {
                    using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                    cts.CancelAfter(TimeoutMs);
                    await process.WaitForExitAsync(cts.Token).ConfigureAwait(false);
                }
                catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
                {
                    try { process.Kill(); }
                    catch (InvalidOperationException) { /* 进程已退出 */ }
                    throw new TimeoutException("Pandoc转换超时（30秒）");
                }

                if (process.ExitCode != 0)
                    throw new InvalidOperationException($"Pandoc转换失败: {await errorTask.ConfigureAwait(false)}");

                // 确保 stderr 被读取完毕，避免资源泄漏
                _ = await errorTask.ConfigureAwait(false);

                return await outputTask.ConfigureAwait(false);
            }
            finally
            {
                _semaphore.Release();
            }
        }

        public static async Task<bool> ConvertFileAsync(string inputFile, string outputFile, string[] extraArgs, CancellationToken cancellationToken = default)
        {
            if (string.IsNullOrWhiteSpace(inputFile)) throw new ArgumentException("输入文件路径不能为空", nameof(inputFile));
            if (string.IsNullOrWhiteSpace(outputFile)) throw new ArgumentException("输出文件路径不能为空", nameof(outputFile));
            if (!File.Exists(inputFile)) throw new FileNotFoundException("输入文件不存在", inputFile);
            EnsurePandocAvailable();
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "pandoc",
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardErrorEncoding = Encoding.UTF8,
                    StandardOutputEncoding = Encoding.UTF8
                }
            };
            process.StartInfo.ArgumentList.Add(inputFile);
            process.StartInfo.ArgumentList.Add("-o");
            process.StartInfo.ArgumentList.Add(outputFile);
            foreach (var arg in extraArgs)
                process.StartInfo.ArgumentList.Add(arg);
            process.Start();

            var errorTask = process.StandardError.ReadToEndAsync(cancellationToken);

            try
            {
                using var cts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
                cts.CancelAfter(TimeoutMs);
                await process.WaitForExitAsync(cts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                try { process.Kill(); }
                catch (InvalidOperationException) { /* 进程已退出 */ }
                return false;
            }

            // 确保 stderr 被读取完毕，避免资源泄漏
            if (process.ExitCode != 0)
                _ = await errorTask.ConfigureAwait(false);

            return process.ExitCode == 0;
        }

        private static Process CreateProcess(IEnumerable<string> arguments)
        {
            var proc = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "pandoc",
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    StandardInputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                    StandardOutputEncoding = Encoding.UTF8
                }
            };
            foreach (var arg in arguments)
                proc.StartInfo.ArgumentList.Add(arg);
            return proc;
        }

        /// <summary>使用 using 包装创建和启动 Process，确保释放</summary>
        private static Process CreateAndStartProcess(IEnumerable<string> arguments)
        {
            var proc = CreateProcess(arguments);
            proc.Start();
            return proc;
        }

        /// <summary>
        /// 预处理 markdown 中的 $$...$$ 数学公式块：若块内存在空行（pandoc 不允许），
        /// 对分隔行和内容做字符转义，使 pandoc 将 LaTeX 源码原样输出。
        /// - $$ 纯分隔行 → &#36;&#36;（阻止 pandoc 数学模式解析）
        /// - 内容中的 _ → \_，^ → \^（阻止 pandoc 强调/上标解析）
        /// 无空行的有效公式块不受影响。
        /// </summary>
        private static string SanitizeMathBlocks(string markdown)
        {
            // 收集所有纯 $$ 行的行起始偏移
            var dollarOffsets = new List<int>();
            var span = markdown.AsSpan();
            int pos = 0;

            while (pos < span.Length)
            {
                int lineStart = pos;
                int nl = span[pos..].IndexOf('\n');
                int lineEnd = nl >= 0 ? pos + nl : span.Length;

                int p = lineStart;
                while (p < lineEnd && span[p] == ' ') p++;
                if (p + 1 < lineEnd && span[p] == '$' && span[p + 1] == '$')
                {
                    int k = p + 2;
                    while (k < lineEnd && span[k] == ' ') k++;
                    if (k >= lineEnd)
                        dollarOffsets.Add(lineStart);
                }

                pos = nl >= 0 ? lineEnd + 1 : span.Length;
            }

            if (dollarOffsets.Count < 2)
                return markdown;

            // 检查哪些 $$ 对内部有空行，记录这些块的起止范围
            var blocksToClean = new List<(int openOff, int closeOff)>();
            for (int p = 0; p + 1 < dollarOffsets.Count; p += 2)
            {
                int openOff = dollarOffsets[p];
                int closeOff = dollarOffsets[p + 1];

                bool hasEmpty = false;
                int scanPos = openOff;
                while (scanPos < closeOff)
                {
                    int nl = span[scanPos..].IndexOf('\n');
                    int lineEnd = nl >= 0 ? scanPos + nl : closeOff;
                    int j = scanPos;
                    while (j < lineEnd && span[j] == ' ') j++;
                    if (j >= lineEnd && scanPos != openOff)
                    {
                        hasEmpty = true;
                        break;
                    }
                    scanPos = nl >= 0 ? lineEnd + 1 : closeOff;
                }

                if (hasEmpty)
                    blocksToClean.Add((openOff, closeOff));
            }

            if (blocksToClean.Count == 0)
                return markdown;

            // 重建字符串：逐行替换
            var sb = new StringBuilder(markdown.Length + blocksToClean.Count * 50);
            pos = 0;
            foreach (var (openOff, closeLineStart) in blocksToClean)
            {
                sb.Append(span[pos..openOff]);

                int nl = span[closeLineStart..].IndexOf('\n');
                int closeLineEnd = nl >= 0 ? closeLineStart + nl : span.Length;

                int linePos = openOff;
                while (linePos < closeLineEnd)
                {
                    int nextNl = span[linePos..].IndexOf('\n');
                    int lineEnd = nextNl >= 0 ? linePos + nextNl : closeLineEnd;

                    // 判断是否为纯 $$ 分隔行
                    int j = linePos;
                    while (j < lineEnd && span[j] == ' ') j++;
                    bool isDollarLine = j + 1 < lineEnd && span[j] == '$' && span[j + 1] == '$';
                    if (isDollarLine)
                    {
                        int k = j + 2;
                        while (k < lineEnd && span[k] == ' ') k++;
                        isDollarLine = k >= lineEnd;
                    }

                    if (isDollarLine)
                    {
                        // 保留前导空白，$$ 替换为 HTML 实体
                        int ws = 0;
                        while (ws < lineEnd - linePos && span[linePos + ws] == ' ') ws++;
                        if (ws > 0)
                            sb.Append(span[linePos..(linePos + ws)]);
                        sb.Append("&#36;&#36;");
                    }
                    else
                    {
                        // 内容行：转义 _ ^ \（阻止 pandoc 强调/上标/特殊字符解析）
                        for (int i = linePos; i < lineEnd; i++)
                        {
                            char c = span[i];
                            if (c is '_' or '^' or '\\')
                            {
                                sb.Append('\\');
                                sb.Append(c);
                            }
                            else
                            {
                                sb.Append(c);
                            }
                        }
                    }

                    if (lineEnd < closeLineEnd)
                        sb.Append('\n');
                    linePos = lineEnd + 1;
                }

                pos = closeLineEnd;
            }

            sb.Append(span[pos..]);
            return sb.ToString();
        }
    }
}
