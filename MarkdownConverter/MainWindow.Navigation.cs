using System.Windows.Threading;
using Microsoft.Web.WebView2.Core;

namespace MarkdownConverter
{
    /// <summary>MainWindow 的双向导航逻辑：编辑器↔预览</summary>
    /// <remarks>
    /// 核心思路：左右两侧由 pandoc 转换得到，结构完全一致。
    /// 不再使用文本模糊匹配，改用纯结构路径定位：
    ///   (sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal)
    /// 其中：
    ///   sectionIdx    — 按 ## 分段的索引
    ///   headingOrdinal — 该段内第几个标题（-1 表示 preamble 区域）
    ///   blockOrdinal   — 标题后第几个块级元素（p/ul/ol/pre/table/blockquote 等）
    ///   lineOrdinal    — 块内第几个 br 分隔行（0 表示块内首行）
    /// </remarks>
    public partial class MainWindow
    {
        /// <summary>段级 SectionLineMap 缓存：sectionIdx → map</summary>
        private readonly Dictionary<int, SectionLineMap> _lineMapCache = [];

        /// <summary>编辑器 → 预览</summary>
        private async Task NavigatePreviewToCursorAsync(bool selectInEditor = true)
        {
            if (webView2?.CoreWebView2 == null) return;
            _navigateToCursorAfterRender = false;

            SetStatus("正在定位到预览...", ColorConstants.StatusBlue);

            var markdown = txtMarkdown.Text;
            var (sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal) = GetStructuralPath(txtMarkdown.CaretOffset, markdown);
            if (sectionIdx < 0)
            {
                SetStatus("未找到匹配内容", ColorConstants.StatusAmber);
                return;
            }

            var result = await webView2.CoreWebView2.ExecuteScriptAsync(
                $"nav_byPath({sectionIdx},{headingOrdinal},{blockOrdinal},{lineOrdinal})");

            // 左侧选中当前行（双击时 e.Handled=true 阻止了默认选中）
            if (selectInEditor)
            {
                Dispatcher.Invoke(() => SelectAndScrollTo(txtMarkdown.CaretOffset));
            }

            if (result == "true")
            {
                SetStatus("已定位到预览", ColorConstants.StatusGreen);
            }
            else
            {
                SetStatus("未找到匹配内容", ColorConstants.StatusAmber);
            }
        }

        /// <summary>预览 → 编辑器（接收 WebView2 回传消息）</summary>
        private void OnPreviewWebMessageReceived(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
        {
            var raw = e.TryGetWebMessageAsString();
            if (string.IsNullOrEmpty(raw)) return;

            try
            {
                using var doc = System.Text.Json.JsonDocument.Parse(raw);
                var sectionIdx = doc.RootElement.GetProperty("sectionIdx").GetInt32();
                var headingOrdinal = doc.RootElement.GetProperty("headingOrdinal").GetInt32();
                var blockOrdinal = doc.RootElement.GetProperty("blockOrdinal").GetInt32();
                var lineOrdinal = doc.RootElement.GetProperty("lineOrdinal").GetInt32();

                if (_isConverting)
                {
                    var ranges = _incrementalConverter.GetSectionCharRanges(txtMarkdown.Text);
                    if (sectionIdx < 0 || sectionIdx >= ranges.Count || !_renderedSectionIndices.Contains(sectionIdx))
                    {
                        SetStatus("请等待转换完成后，再点击定位...", ColorConstants.StatusAmber);
                        return;
                    }
                }

                SetStatus($"点击：s={sectionIdx} h={headingOrdinal} b={blockOrdinal} l={lineOrdinal}", ColorConstants.StatusBlue);
                NavigateEditorByPath(sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal, txtMarkdown.Text);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Web消息解析失败: {ex.Message}");
            }
        }

        /// <summary>计算光标位置的结构路径</summary>
        private (int sectionIdx, int headingOrdinal, int blockOrdinal, int lineOrdinal)
            GetStructuralPath(int cursorOffset, string markdown)
        {
            var sectionRanges = _incrementalConverter.GetSectionCharRanges(markdown);
            int sectionIdx = GetSectionIndexAtCaret();
            if (sectionIdx < 0 || sectionIdx >= sectionRanges.Count)
                return (-1, -1, -1, -1);

            var (secStart, secEnd) = sectionRanges[sectionIdx];
            int relOffset = cursorOffset - secStart;
            if (relOffset < 0) relOffset = 0;
            if (relOffset > secEnd - secStart) relOffset = secEnd - secStart;

            var sectionText = markdown.AsSpan(secStart, secEnd - secStart);
            var map = GetOrBuildLineMap(sectionIdx, sectionText);

            // 检查光标是否在标题行上
            for (int i = 0; i < map.Headings.Count; i++)
            {
                int hStart = map.Headings[i].offset;
                int hEnd = map.Headings[i].end;
                if (relOffset >= hStart && relOffset <= hEnd)
                    return (sectionIdx, i, -1, 0);
            }

            // 确定光标前的 headingOrdinal
            int headingOrdinal = -1;
            for (int i = 0; i < map.Headings.Count; i++)
            {
                if (map.Headings[i].offset < relOffset)
                    headingOrdinal = i;
                else
                    break;
            }

            // 在内容行映射中查找光标位置所属的行
            int blockOrdinal = 0;
            int lineOrdinal = 0;
            for (int i = 0; i < map.ContentOffsets.Count; i++)
            {
                int lineOff = map.ContentOffsets[i];
                if (lineOff <= relOffset)
                {
                    blockOrdinal = map.ContentBlockOrdinals[i];
                    lineOrdinal = map.ContentLineOrdinals[i];
                }
                else break;
            }

            return (sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal);
        }

        /// <summary>按结构路径定位编辑器中的位置</summary>
        private void NavigateEditorByPath(int sectionIdx, int headingOrdinal, int blockOrdinal,
            int lineOrdinal, string markdown)
        {
            var sectionRanges = _incrementalConverter.GetSectionCharRanges(markdown);
            if (sectionIdx < 0 || sectionIdx >= sectionRanges.Count)
            {
                SetStatus("未找到匹配内容", ColorConstants.StatusAmber);
                return;
            }

            var (secStart, secEnd) = sectionRanges[sectionIdx];
            var sectionText = markdown.AsSpan(secStart, secEnd - secStart);
            var map = GetOrBuildLineMap(sectionIdx, sectionText);

            int targetOffset = -1;

            if (blockOrdinal == -1)
            {
                if (headingOrdinal >= 0 && headingOrdinal < map.Headings.Count)
                    targetOffset = secStart + map.Headings[headingOrdinal].offset;
            }
            else
            {
                // 在内容行中查找匹配的 (headingOrdinal, blockOrdinal, lineOrdinal)
                for (int i = 0; i < map.ContentOffsets.Count; i++)
                {
                    if (map.ContentHeadingOrdinals[i] == headingOrdinal &&
                        map.ContentBlockOrdinals[i] == blockOrdinal &&
                        map.ContentLineOrdinals[i] == lineOrdinal)
                    {
                        targetOffset = secStart + map.ContentOffsets[i];
                        break;
                    }
                }
            }

            if (targetOffset >= 0)
            {
                Dispatcher.Invoke(() =>
                {
                    txtMarkdown.Focus();
                    SelectAndScrollTo(targetOffset);
                });
                SetStatus("已定位到编辑器", ColorConstants.StatusGreen);
            }
            else
            {
                SetStatus("未找到匹配内容", ColorConstants.StatusAmber);
            }
        }

        /// <summary>获取或构建段级行映射（带缓存）</summary>
        private SectionLineMap GetOrBuildLineMap(int sectionIdx, ReadOnlySpan<char> sectionText)
        {
            if (_lineMapCache.TryGetValue(sectionIdx, out var cached))
                return cached;

            var map = BuildSectionLineMap(sectionText);
            _lineMapCache[sectionIdx] = map;
            return map;
        }

        /// <summary>单次遍历构建段内行映射：标题行 + 内容行结构</summary>
        private static SectionLineMap BuildSectionLineMap(ReadOnlySpan<char> text)
        {
            var map = new SectionLineMap();

            int pos = 0;
            bool inCodeFence = false;
            bool inBlock = false;
            bool blockIsListItem = false;
            string blockListItemStyle = "";
            bool inTable = false;
            int currentBlock = -1;
            int currentLine = 0;

            while (pos < text.Length)
            {
                int lineStart = pos;
                int nl = text[pos..].IndexOf('\n');
                int lineEnd = nl >= 0 ? pos + nl : text.Length;
                var trimmed = text[lineStart..lineEnd].Trim();

                // 代码围栏
                if (trimmed.StartsWith("```"))
                {
                    if (!inCodeFence) currentBlock++;
                    inCodeFence = !inCodeFence;
                    if (!inCodeFence) { inBlock = false; currentLine = 0; blockIsListItem = false; blockListItemStyle = "";
                    }
                    pos = nl >= 0 ? lineEnd + 1 : text.Length;
                    continue;
                }

                if (inCodeFence)
                {
                    map.AddContent(currentBlock, currentLine, lineStart, map.Headings.Count - 1);
                    currentLine++;
                    pos = nl >= 0 ? lineEnd + 1 : text.Length;
                    continue;
                }

                // 标题行 → 块边界，退出表格
                if (trimmed.Length > 0 && trimmed[0] == '#')
                {
                    map.Headings.Add((lineStart, lineEnd));
                    inBlock = false;
                    currentBlock = -1;
                    currentLine = 0;
                    blockIsListItem = false;
                    blockListItemStyle = "";
                    inTable = false;
                    pos = nl >= 0 ? lineEnd + 1 : text.Length;
                    continue;
                }

                // 空行 → 块边界（同样式列表项间的空行不分割），同时退出表格
                if (trimmed.Length == 0)
                {
                    if (inBlock)
                    {
                        bool listItemBreak = false;
                        if (blockIsListItem)
                        {
                            var nextStyle = GetNextLineListItemStyle(text, nl, lineEnd);
                            // 空行后如果是同样式的列表项，不分割 block；否则分割
                            listItemBreak = !string.IsNullOrEmpty(nextStyle) && nextStyle == blockListItemStyle;
                        }
                        if (!listItemBreak)
                        {
                            inBlock = false;
                            currentLine = 0;
                            blockIsListItem = false;
                            blockListItemStyle = "";
                        }
                    }
                    inTable = false;
                    pos = nl >= 0 ? lineEnd + 1 : text.Length;
                    continue;
                }

                // 管道表格分隔行 → 跳过，标记在表格中
                if (IsTableSeparatorRow(trimmed))
                {
                    inTable = true;
                    pos = nl >= 0 ? lineEnd + 1 : text.Length;
                    continue;
                }

                // ═══════════════════════════════════════════════════════
                //  管道表格表头前瞻检测
                //  若当前行以 | 开头且下一行为表格分隔行，说明是表头行。
                //  此时若当前块已有非表格内容（不以 | 开头），则结束当前块，
                //  使表格独占一个新块，与 pandoc HTML 输出的块结构对齐。
                // ═══════════════════════════════════════════════════════
                if (inBlock && trimmed[0] == '|')
                {
                    // 前瞻检查下一行是否为表格分隔行
                    int nextLineStart = nl >= 0 ? lineEnd + 1 : text.Length;
                    if (nextLineStart < text.Length)
                    {
                        int snl = text[nextLineStart..].IndexOf('\n');
                        int snlEnd = snl >= 0 ? nextLineStart + snl : text.Length;
                        var nextTrimmed = text[nextLineStart..snlEnd].Trim();
                        if (IsTableSeparatorRow(nextTrimmed))
                        {
                            // 当前行是表头行：检查当前块首行是否以 | 开头
                            bool blockHasNonTableContent = false;
                            for (int ci = map.ContentOffsets.Count - 1; ci >= 0; ci--)
                            {
                                if (map.ContentBlockOrdinals[ci] == currentBlock)
                                {
                                    if (text[map.ContentOffsets[ci]] != '|')
                                        blockHasNonTableContent = true;
                                    break;
                                }
                            }
                            if (blockHasNonTableContent)
                            {
                                // 结束当前块，表格将开启新块
                                inBlock = false;
                                currentLine = 0;
                                blockIsListItem = false;
                                blockListItemStyle = "";
                            }
                        }
                    }
                }

                // ═══════════════════════════════════════════════════════
                //  表格退出检测：表格结束后（已见过分隔行）遇到非 | 开头的行，
                //  说明表格已结束，该行应属于新块。
                // ═══════════════════════════════════════════════════════
                if (inTable && trimmed[0] != '|')
                {
                    inTable = false;
                    inBlock = false;
                    currentLine = 0;
                    blockIsListItem = false;
                    blockListItemStyle = "";
                }

                // 内容行
                bool isThisLineLi = IsListItem(trimmed);
                _ = GetIndent(text, lineStart, lineEnd); // 计算前导空格数

                if (!inBlock)
                {
                    currentBlock++;
                    currentLine = 0;
                    inBlock = true;
                    blockIsListItem = isThisLineLi;
                    blockListItemStyle = isThisLineLi ? GetListItemStyle(trimmed) : "";
                }
                else if (isThisLineLi)
                {
                    // 每个列表项行都开启独立 block（与 HTML 端每个 <li> 独立 data-block 对齐）
                    currentBlock++;
                    currentLine = 0;
                    blockIsListItem = true;
                    blockListItemStyle = GetListItemStyle(trimmed);
                }

                map.AddContent(currentBlock, currentLine, lineStart, map.Headings.Count - 1);
                currentLine++;
                pos = nl >= 0 ? lineEnd + 1 : text.Length;
            }

            return map;
        }

        private sealed class SectionLineMap
        {
            /// <summary>标题行信息：(offset, end) 相对段起始</summary>
            public List<(int offset, int end)> Headings { get; } = [];
            /// <summary>内容行偏移（相对段起始）</summary>
            public List<int> ContentOffsets { get; } = [];
            /// <summary>内容行所属 blockOrdinal（在每个 heading 后重置为 0）</summary>
            public List<int> ContentBlockOrdinals { get; } = [];
            /// <summary>内容行在块内的 lineOrdinal</summary>
            public List<int> ContentLineOrdinals { get; } = [];
            /// <summary>内容行所属的 headingOrdinal（-1 = preamble）</summary>
            public List<int> ContentHeadingOrdinals { get; } = [];

            public void AddContent(int block, int line, int offset, int headingOrdinal)
            {
                ContentBlockOrdinals.Add(block);
                ContentLineOrdinals.Add(line);
                ContentOffsets.Add(offset);
                ContentHeadingOrdinals.Add(headingOrdinal);
            }
        }

        // ═══════════════════════════════════════════════════════════
        //  行类型判定辅助（纯静态，不捕获上下文）
        // ═══════════════════════════════════════════════════════════

        /// <summary>获取 offset 所在行的行首偏移量（0-based）。offset 可以是行内任意位置或行尾。</summary>
        private static int GetLineStart(string text, int offset)
        {
            if (string.IsNullOrEmpty(text) || offset <= 0) return 0;
            int clamped = Math.Min(offset, text.Length);
            int nl = text.LastIndexOf('\n', clamped - 1);
            return nl + 1; // LastIndexOf 返回 -1 时表示第 0 行
        }

        /// <summary>检查 [start, end) 范围内是否全为空白字符</summary>
        private static bool IsAllSpaces(string text, int start, int end)
        {
            for (int i = start; i < end; i++)
                if (!char.IsWhiteSpace(text[i])) return false;
            return true;
        }

        /// <summary>检查行首是否为列表项（无序或有序），返回非负表示是列表项</summary>
        private static bool CheckIsListItem(string text, int lineStart, int lineEnd)
        {
            int firstChar = lineStart;
            while (firstChar < lineEnd && text[firstChar] == ' ') firstChar++;
            if (firstChar >= lineEnd) return false;

            char c = text[firstChar];
            if (c == '-' || c == '*' || c == '+')
                return firstChar + 1 < lineEnd && text[firstChar + 1] == ' ';
            if (char.IsDigit(c))
            {
                int d = firstChar;
                while (d < lineEnd && char.IsDigit(text[d])) d++;
                if (d < lineEnd && (text[d] == '.' || text[d] == ')') && d + 1 < lineEnd && text[d + 1] == ' ')
                    return true;
            }
            return false;
        }

        /// <summary>判断是否为 Markdown 列表项行首（与 pandoc 解析规则对齐）</summary>
        private static bool IsListItem(ReadOnlySpan<char> trimmed)
        {
            if (trimmed.Length == 0) return false;

            // 括号有序列表：(1) text 或 (1)text（pandoc 原生支持）
            if (trimmed[0] == '(')
            {
                int j = 1;
                while (j < trimmed.Length && char.IsDigit(trimmed[j])) j++;
                if (j > 1 && j < trimmed.Length && trimmed[j] == ')')
                {
                    // ) 后是空格或结束，都视为列表项
                    return j + 1 >= trimmed.Length || trimmed[j + 1] == ' ';
                }
            }

            // 有序列表：数字 + '.' 或 ')' 后必须紧跟空白（排除 "1.1 版本号" 等误匹配）
            int i = 0;
            while (i < trimmed.Length && char.IsDigit(trimmed[i])) i++;
            if (i > 0 && i < trimmed.Length && (trimmed[i] == '.' || trimmed[i] == ')'))
                return i + 1 < trimmed.Length && trimmed[i + 1] == ' ';

            // 无序列表项：*, -, +（排除 **bold** 等非列表用法）
            if (trimmed[0] == '*' && (trimmed.Length == 1 || trimmed[1] != '*')) return true;
            if (trimmed[0] == '-' && (trimmed.Length == 1 || trimmed[1] != '-')) return true;
            if (trimmed[0] == '+' && (trimmed.Length == 1 || trimmed[1] != '+')) return true;
            return false;
        }

        /// <summary>获取列表项样式标识。同一样式的列表项在空行后合并 block，不同样式则分段。</summary>
        /// <returns>"digit"（数字 + 点/括号）、"paren"（括号数字）、"bullet"（无序）、""（非列表项）</returns>
        private static string GetListItemStyle(ReadOnlySpan<char> trimmed)
        {
            if (trimmed.Length == 0) return "";

            // 括号有序列表：(1) 样式
            if (trimmed[0] == '(')
            {
                int j = 1;
                while (j < trimmed.Length && char.IsDigit(trimmed[j])) j++;
                if (j > 1 && j < trimmed.Length && trimmed[j] == ')')
                {
                    // ) 后是空格或结束，都视为 paren 样式
                    return j + 1 >= trimmed.Length || trimmed[j + 1] == ' ' ? "paren" : "";
                }
            }

            // 有序列表：数字 + '.' 或 ')' 
            int i = 0;
            while (i < trimmed.Length && char.IsDigit(trimmed[i])) i++;
            if (i > 0 && i < trimmed.Length && (trimmed[i] == '.' || trimmed[i] == ')'))
                return i + 1 < trimmed.Length && trimmed[i + 1] == ' ' ? "digit" : "";

            // 无序列表项：*, -, +
            if (trimmed[0] == '*' && (trimmed.Length == 1 || trimmed[1] != '*')) return "bullet";
            if (trimmed[0] == '-' && (trimmed.Length == 1 || trimmed[1] != '-')) return "bullet";
            if (trimmed[0] == '+' && (trimmed.Length == 1 || trimmed[1] != '+')) return "bullet";

            return "";
        }

        /// <summary>判断是否为管道表格分隔行（如 |---|---|）</summary>
        private static bool IsTableSeparatorRow(ReadOnlySpan<char> trimmed)
        {
            if (trimmed.Length == 0 || trimmed[0] != '|') return false;
            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (c != '|' && c != '-' && c != ':' && c != ' ') return false;
            }
            return true;
        }

        /// <summary>判断空行后的第一个非空行是否为列表项</summary>
        private static bool IsNextLineListItem(ReadOnlySpan<char> text, int currentNl, int currentLineEnd)
        {
            int scanPos = currentNl >= 0 ? currentLineEnd + 1 : text.Length;
            while (scanPos < text.Length)
            {
                int snl = text[scanPos..].IndexOf('\n');
                int sEnd = snl >= 0 ? scanPos + snl : text.Length;
                var sTrimmed = text[scanPos..sEnd].Trim();
                if (sTrimmed.Length > 0)
                    return IsListItem(sTrimmed);
                scanPos = snl >= 0 ? sEnd + 1 : text.Length;
            }
            return false;
        }

        /// <summary>获取空行后的第一个非空行的列表项样式</summary>
        private static string GetNextLineListItemStyle(ReadOnlySpan<char> text, int currentNl, int currentLineEnd)
        {
            int scanPos = currentNl >= 0 ? currentLineEnd + 1 : text.Length;
            while (scanPos < text.Length)
            {
                int snl = text[scanPos..].IndexOf('\n');
                int sEnd = snl >= 0 ? scanPos + snl : text.Length;
                var sTrimmed = text[scanPos..sEnd].Trim();
                if (sTrimmed.Length > 0)
                    return GetListItemStyle(sTrimmed);
                scanPos = snl >= 0 ? sEnd + 1 : text.Length;
            }
            return "";
        }

        /// <summary>计算行的前导空格数（用于识别嵌套列表）</summary>
        private static int GetIndent(ReadOnlySpan<char> text, int lineStart, int lineEnd)
        {
            int indent = 0;
            for (int i = lineStart; i < lineEnd && i < text.Length; i++)
            {
                if (text[i] == ' ') indent++;
                else break;
            }
            return indent;
        }

        /// <summary>查找同行/跨行公式（$$...$$ 或 $...$）的行范围。光标严格在分隔符之间才匹配。空行是公式界限。</summary>
        /// <returns>(打开行行首, 闭合行行尾)，不含换行符。非跨行公式内返回 (-1, -1)。</returns>
        private static (int lineStart, int lineEnd) FindMultiLineFormulaSpan(string markdown, int caretOffset)
        {
            if (string.IsNullOrEmpty(markdown) || caretOffset < 0 || caretOffset > markdown.Length)
                return (-1, -1);

            // ── 表格行保护 ──
            // 表格行内的 $$/ $ 公式都是自包含的内联公式，不可能是跨行公式块。
            int tableLineStart = GetLineStart(markdown, caretOffset);
            if (tableLineStart < markdown.Length)
            {
                int firstNonSpace = tableLineStart;
                while (firstNonSpace < markdown.Length && markdown[firstNonSpace] == ' ') firstNonSpace++;
                if (firstNonSpace < markdown.Length && markdown[firstNonSpace] == '|')
                    return (-1, -1);
            }

            // ── $$...$$（同行或跨行） ──
            var result = TryFindDollarDollarBlock(markdown, caretOffset);
            if (result.lineStart >= 0) return result;

            // ── $...$ 跨行 ──
            return TryFindSingleDollarCrossLine(markdown, caretOffset);
        }

        /// <summary>从全文扫描行首的 $$...$$ 块，返回光标所在块的范围。</summary>
        private static (int lineStart, int lineEnd) TryFindDollarDollarBlock(string markdown, int caretOffset)
        {
            int i = 0;
            while (i + 1 < markdown.Length)
            {
                if (markdown[i] != '$' || markdown[i + 1] != '$')
                {
                    i++;
                    continue;
                }

                // 检查 $$ 是否在行首（允许前导空白）
                int lineStart = i;
                while (lineStart > 0 && markdown[lineStart - 1] != '\n') lineStart--;
                bool atLineStart = true;
                for (int k = lineStart; k < i; k++)
                {
                    if (markdown[k] != ' ') { atLineStart = false; break; }
                }

                if (!atLineStart)
                {
                    i++;
                    continue;
                }

                int openLineStart = lineStart;
                int pos = i + 2;
                bool hasClose = false;
                int closeLineEnd = -1;

                while (pos < markdown.Length)
                {
                    int nl = markdown.IndexOf('\n', pos);
                    int curLineEnd = nl >= 0 ? nl : markdown.Length;

                    // 空行 → 不构成完整公式块（仅在非打开行时检查）
                    if (pos > i + 2)
                    {
                        if (IsAllSpaces(markdown, pos, curLineEnd)) break;
                    }

                    // 查找行尾的 $$
                    int scan = pos;
                    while (scan < curLineEnd)
                    {
                        if (scan + 1 < curLineEnd && markdown[scan] == '$' && markdown[scan + 1] == '$')
                        {
                            int k = scan + 2;
                            while (k < curLineEnd && markdown[k] == ' ') k++;
                            if (k >= curLineEnd)
                            {
                                hasClose = true;
                                closeLineEnd = curLineEnd;
                                goto endDDSearch;
                            }
                            scan = k;
                        }
                        else scan++;
                    }

                    pos = nl >= 0 ? curLineEnd + 1 : markdown.Length;
                }

                endDDSearch:
                if (hasClose)
                {
                    if (caretOffset >= openLineStart && caretOffset < closeLineEnd)
                        return (openLineStart, closeLineEnd);
                    i = closeLineEnd;
                    continue;
                }
                i += 2;
            }

            return (-1, -1);
        }

        /// <summary>从光标位置向后扫描跨行 $...$ 公式。</summary>
        private static (int lineStart, int lineEnd) TryFindSingleDollarCrossLine(string markdown, int caretOffset)
        {
            // 从光标所在行末尾向后扫描，确保同一行上光标前后的 $ 都不会漏掉
            int lineEnd = markdown.IndexOf('\n', caretOffset);
            if (lineEnd < 0) lineEnd = markdown.Length;
            int scanStart = Math.Min(Math.Max(caretOffset, lineEnd - 1), markdown.Length - 1);
            if (scanStart < 0) scanStart = 0;

            // 计算向后扫描的边界：不越过列表项行首（跨列表项的 $ 不应配对）
            int scanLimit = 0;
            int caretLineStart = GetLineStart(markdown, caretOffset);
            {
                // 从光标行首向上查找，遇到列表项行首就停止
                int p = caretLineStart;
                while (p > 0)
                {
                    // 找上一行的起始位置
                    int prevNl = markdown.LastIndexOf('\n', p - 1);
                    int prevLineStart = prevNl >= 0 ? prevNl + 1 : 0;
                    // 空行或列表项行首都是边界
                    if (IsAllSpaces(markdown, prevLineStart, p) || CheckIsListItem(markdown, prevLineStart, p))
                    {
                        scanLimit = prevLineStart;
                        break;
                    }
                    p = prevLineStart;
                }
            }

            for (int candidate = scanStart; candidate >= scanLimit; candidate--)
            {
                if (markdown[candidate] != '$')
                    continue;

                // 排除转义 \$：奇数个前导 \ 为转义
                int bslashes = 0, bi = candidate - 1;
                while (bi >= 0 && markdown[bi] == '\\') { bslashes++; bi--; }
                if (bslashes % 2 == 1)
                    continue;

                // 排除 $$ 中的 $（前有 $ 表示是 $$ 的第二个 $，跳过但不停止）
                if (candidate > 0 && markdown[candidate - 1] == '$')
                    continue;

                // 遇到 $$ 的第一个 $ 时停止向后扫描（不越过 $$...$$ 块边界）
                if (candidate + 1 < markdown.Length && markdown[candidate + 1] == '$')
                    break;

                // 排除前有字母数字（如 text$ 不可能是公式开头）
                if (candidate > 0 && char.IsLetterOrDigit(markdown[candidate - 1]))
                    continue;

                // 排除前有闭合括号 } ) ]（这些位置更可能是 LaTeX 闭合 $ 而非公式开头）
                if (candidate > 0)
                {
                    char prev = markdown[candidate - 1];
                    if (prev == '}' || prev == ')' || prev == ']')
                        continue;
                }

                // 排除后跟数字（货币 $5.99）
                if (candidate + 1 < markdown.Length && char.IsDigit(markdown[candidate + 1]))
                    continue;

                // 候选 $ 通过校验，尝试向前找匹配的闭合 $
                // 空行或列表项行首是界限：先找到当前段的结束位置
                int openPos = candidate;
                int paraEnd = markdown.Length;
                for (int p = openPos + 1; p < markdown.Length; p++)
                {
                    if (markdown[p] == '\n')
                    {
                        int nl2 = markdown.IndexOf('\n', p + 1);
                        if (nl2 < 0) nl2 = markdown.Length;
                        // 空行或列表项行首都是段落边界
                        if (IsAllSpaces(markdown, p + 1, nl2) || CheckIsListItem(markdown, p + 1, nl2))
                        {
                            paraEnd = p;
                            break;
                        }
                    }
                }
                for (int j = openPos + 1; j < paraEnd; j++)
                {
                    if (markdown[j] != '$')
                        continue;

                    // 排除转义 \$
                    bslashes = 0;
                    bi = j - 1;
                    while (bi >= 0 && markdown[bi] == '\\') { bslashes++; bi--; }
                    if (bslashes % 2 == 1)
                        continue;

                    // 排除 $$ 中的 $（前接或后接 $ 都是 $$ 的一部分）
                    if ((j > 0 && markdown[j - 1] == '$') ||
                        (j + 1 < markdown.Length && markdown[j + 1] == '$'))
                        continue;

                    // 排除后跟数字（货币）
                    if (j + 1 < markdown.Length && char.IsDigit(markdown[j + 1]))
                        continue;

                    // 注意：闭合 $ 不检查前有字母数字（\infty$ 等 LaTeX 用法合法）

                    int closePos = j;
                    int openLine = GetLineStart(markdown, openPos);
                    int closeLine = GetLineStart(markdown, closePos);

                    if (openLine != closeLine) // 不同行 → 跨行公式
                    {
                        int caretLine = GetLineStart(markdown, caretOffset);
                        if (caretLine >= openLine && caretLine <= closeLine)
                        {
                            int closeLineEnd = markdown.IndexOf('\n', closePos);
                            if (closeLineEnd < 0) closeLineEnd = markdown.Length;
                            return (openLine, closeLineEnd);
                        }
                    }
                    break; // 无论是否跨行，配对完成
                }
            }

            return (-1, -1);
        }

        /// <summary>在编辑器中选中光标所在行（优先匹配跨行公式），并滚动到可见位置</summary>
        private void SelectAndScrollTo(int offset)
        {
            var markdown = txtMarkdown.Text;
            var (lineStart, lineEnd) = FindMultiLineFormulaSpan(markdown, offset);

            // 如果公式匹配成功，检查范围是否合理（不超过 50 行）
            bool useFormulaRange = lineStart >= 0;
            if (useFormulaRange)
            {
                // 用 IndexOf 跳跃式统计行数，快于逐字符扫描
                int newlineCount = 0;
                int searchPos = lineStart;
                while ((searchPos = markdown.IndexOf('\n', searchPos)) >= 0 && searchPos < lineEnd && searchPos < markdown.Length)
                {
                    newlineCount++;
                    searchPos++;
                }
                if (newlineCount > 50)
                {
                    System.Diagnostics.Debug.WriteLine($"[SelectAndScrollTo] 公式范围过大({newlineCount}行)，回退到单行选择。offset={offset} lineStart={lineStart} lineEnd={lineEnd}");
                    useFormulaRange = false;
                }
            }

            if (useFormulaRange)
            {
                txtMarkdown.Select(lineStart, lineEnd - lineStart);
                txtMarkdown.ScrollToLine(txtMarkdown.Document.GetLineByOffset(lineStart).LineNumber);
            }
            else
            {
                int ls = GetLineStart(markdown, offset);
                int nextNl = markdown.IndexOf('\n', ls);
                int le = nextNl >= 0 ? nextNl : markdown.Length;
                txtMarkdown.Select(ls, le - ls);
                txtMarkdown.ScrollToLine(txtMarkdown.Document.GetLineByOffset(ls).LineNumber);
            }
        }

        private int GetSectionIndexAtCaret()
        {
            var markdown = txtMarkdown.Text;
            var ranges = _incrementalConverter.GetSectionCharRanges(markdown);
            int caretPos = txtMarkdown.CaretOffset;
            if (ranges.Count == 0) return -1;

            // 二分查找：找到 Start <= caretPos 的最后一个段
            int lo = 0, hi = ranges.Count - 1, result = -1;
            while (lo <= hi)
            {
                int mid = lo + (hi - lo) / 2;
                if (ranges[mid].Start <= caretPos)
                {
                    result = mid;
                    lo = mid + 1;
                }
                else
                {
                    hi = mid - 1;
                }
            }

            if (result >= 0 && caretPos < ranges[result].End)
                return result;
            if (result >= 0 && caretPos >= ranges[^1].Start)
                return ranges.Count - 1;
            return -1;
        }
    }
}