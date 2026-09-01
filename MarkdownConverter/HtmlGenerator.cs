using System.Text;
using System.Text.RegularExpressions;

namespace MarkdownConverter
{
    public static partial class HtmlGenerator
    {
        /// <summary>使用 CSS 变量 + class 切换的主题方案，替代 JS 注入样式</summary>
        private static string ThemeCss { get; } = BuildThemeCss();

        private static string BuildThemeCss()
        {
            var lc = ThemeManager.GetColors(false);
            var dc = ThemeManager.GetColors(true);
            var sb = new StringBuilder(3000);

            // CSS 变量定义：浅色（默认）+ 深色（.dark-theme）
            sb.Append(":root{");
            sb.Append($"--text:{lc.Text};--heading:{lc.Heading};--border:{lc.Border};--link:{lc.Link};");
            sb.Append($"--blockquote-bg:{lc.BlockquoteBg};--blockquote-text:{lc.BlockquoteText};--blockquote-border:{lc.BlockquoteBorder};");
            sb.Append($"--code-bg:{lc.CodeBg};--code-text:{lc.CodeText};--code-border:{lc.CodeBorder};");
            sb.Append($"--table-border:{lc.TableBorder};--body-bg:{lc.BodyBg};");
            sb.Append('}');
            sb.Append(".dark-theme{");
            sb.Append($"--text:{dc.Text};--heading:{dc.Heading};--border:{dc.Border};--link:{dc.Link};");
            sb.Append($"--blockquote-bg:{dc.BlockquoteBg};--blockquote-text:{dc.BlockquoteText};--blockquote-border:{dc.BlockquoteBorder};");
            sb.Append($"--code-bg:{dc.CodeBg};--code-text:{dc.CodeText};--code-border:{dc.CodeBorder};");
            sb.Append($"--table-border:{dc.TableBorder};--body-bg:{dc.BodyBg};");
            sb.Append('}');

            // 使用 CSS 变量的样式规则
            sb.Append("html{color:#1a1a1a;background-color:transparent}");
            sb.Append("body{direction:ltr;unicode-bidi:plaintext;font-family:'Segoe UI',sans-serif;line-height:1.6;padding:20px;max-width:850px;margin:0 auto;overflow-x:hidden;color:var(--text);background-color:var(--body-bg)}");
            sb.Append("h1,h2,h3,h4,h5,h6{color:var(--heading);border-bottom-color:var(--border)}");
            sb.Append("h1{border-bottom:2px solid var(--border);padding-bottom:10px}");
            sb.Append("a{color:var(--link);text-decoration:none}");
            sb.Append("a:hover{text-decoration:underline}");
            sb.Append("img{max-width:100%;height:auto;display:block;margin:10px 0;border-radius:4px}");
            sb.Append("figure{margin:10px 0;text-align:center}");
            sb.Append("figure img{margin:0 auto}");
            sb.Append("figcaption{display:block;margin-top:6px;font-size:0.9em;color:inherit;text-align:center}");
            sb.Append("blockquote{border-left:4px solid var(--blockquote-border);padding-left:16px;margin-left:0;color:var(--blockquote-text);background:var(--blockquote-bg)}");
            sb.Append("table{border-collapse:collapse;width:100%;margin:16px 0}");
            sb.Append("th,td{border:1px solid var(--table-border);padding:10px;text-align:left;background:transparent}");
            sb.Append(".table-wrapper{overflow-x:auto;max-width:100%;margin:16px 0}.table-wrapper table{margin:0}");
            sb.Append("ul,ol{padding-left:24px}");
            sb.Append("pre,code{background:var(--code-bg);color:var(--code-text);border-color:var(--code-border)}");
            sb.Append("pre{white-space:pre-wrap;border:1px solid var(--code-border);padding:12px;border-radius:5px}");
            sb.Append("code{font-family:'Consolas',monospace;padding:2px 4px;border-radius:3px}");
            sb.Append(".task-list{margin:4px 0;padding-left:20px}.task-list li{margin:2px 0}.task-list input[type='checkbox']{margin-right:6px;width:14px;height:14px;cursor:pointer}.task-list li.completed{text-decoration:line-through;opacity:.7}");
            sb.Append("math{font-size:1.1em;font-family:math}");
            sb.Append(".md-section{content-visibility:auto;contain-intrinsic-size:500px}");
            sb.Append("::selection{background:Highlight;color:HighlightText}");
            return sb.ToString();
        }

        public static string GenerateHtmlContent(string content, bool darkTheme)
        {
            content = ProcessImageCaptions(content);
            content = WrapTables(content);
            var sb = CreateHtmlBuilder(darkTheme, content.Length);
            sb.Append(content);
            return FinalizeHtml(sb);
        }

        public static string GenerateShellHtml(int sectionCount, bool darkTheme)
        {
            var sb = CreateHtmlBuilder(darkTheme, sectionCount * 80);
            for (int i = 0; i < sectionCount; i++)
                sb.Append("<div class=\"md-section\" data-section-idx=\"").Append(i).Append("\"></div>");
            return FinalizeHtml(sb);
        }

        private static StringBuilder CreateHtmlBuilder(bool darkTheme, int extraCapacity)
        {
            var htmlClass = darkTheme ? " class=\"dark-theme\"" : "";
            var bodyStyle = " style=\"margin:0\"";
            var sb = new StringBuilder(ThemeCss.Length + extraCapacity + 300);
            sb.Append("<!DOCTYPE html><html xmlns=\"http://www.w3.org/1999/xhtml\"").Append(htmlClass).Append("><head><meta charset=\"UTF-8\">");
            sb.Append("<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; img-src http://markdown.local https: data:; font-src http://markdown.local https: data:;\">");
            sb.Append("<style>");
            sb.Append(ThemeCss);
            sb.Append("</style></head><body").Append(bodyStyle).Append('>');
            return sb;
        }

        private static string FinalizeHtml(StringBuilder sb)
        {
            sb.Append("</body></html>");
            return sb.ToString();
        }

        /// <summary>将 HTML 中的 &lt;table&gt; 包裹在可横向滚动的 &lt;div class="table-wrapper"&gt; 中</summary>
        public static string WrapTables(string html)
        {
            if (html.AsSpan().IndexOf("<table") < 0) return html;
            return TableRegex().Replace(html, m =>
                $"<div class=\"table-wrapper\">{m.Value}</div>");
        }

        [GeneratedRegex(@"<table\b[^>]*>.*?</table>", RegexOptions.Singleline | RegexOptions.IgnoreCase)]
        private static partial Regex TableRegex();

        [GeneratedRegex(@"<img\s+[^>]*?alt\s*=\s*""([^""]*)""[^>]*>", RegexOptions.IgnoreCase)]
        private static partial Regex ImgWithAltRegex();

        /// <summary>
        /// 给所有顶级块级元素添加 data-block="N" 和 data-heading="M" 属性，
        /// 使 JS 端可直接读取而无需自行计算，彻底消除 C#/JS 两侧计数不一致。
        /// </summary>
        public static string AddBlockOrdinals(string html)
        {
            var sb = new StringBuilder(html.Length + 400);
            int pos = 0;
            int blockOrdinal = -1;  // 从 -1 开始，每次遇到标签递增；遇到标题重置
            int headingOrdinal = -1;
            // 栈跟踪 <li> 内是否应跳过首 <p>（Pandoc 将列表项文本包裹在 <p> 中）
            // 每层 <li> 入栈 true，遇到首 <p> 跳过后或遇到其他标签后设为 false
            List<bool> skipPInLi = new List<bool>();

            while (pos < html.Length)
            {
                int lt = html.IndexOf('<', pos);
                if (lt < 0) { sb.Append(html, pos, html.Length - pos); break; }

                // 注释
                if (lt + 3 < html.Length && html[lt + 1] == '!' && html[lt + 2] == '-' && html[lt + 3] == '-')
                {
                    int end = html.IndexOf("-->", lt + 4);
                    if (end < 0) { sb.Append(html, pos, html.Length - pos); break; }
                    sb.Append(html, pos, end - pos + 3);
                    pos = end + 3;
                    continue;
                }

                // DOCTYPE
                if (lt + 8 < html.Length &&
                    (html[lt + 1] == 'D' || html[lt + 1] == 'd') &&
                    (html[lt + 2] == 'O' || html[lt + 2] == 'o') &&
                    char.ToLowerInvariant(html[lt + 3]) == 'c' &&
                    char.ToLowerInvariant(html[lt + 4]) == 't')
                {
                    int gt = html.IndexOf('>', lt);
                    if (gt < 0) { sb.Append(html, pos, html.Length - pos); break; }
                    sb.Append(html, pos, gt - pos + 1);
                    pos = gt + 1;
                    continue;
                }

                int tagClose = html.IndexOf('>', lt);
                if (tagClose < 0) { sb.Append(html, pos, html.Length - pos); break; }

                // 关闭标签 </xxx>
                if (lt + 1 < html.Length && html[lt + 1] == '/')
                {
                    // 解析关闭标签名，维护 skipPInLi 栈
                    int cnStart = lt + 2;
                    int cnEnd = cnStart;
                    while (cnEnd < html.Length && !char.IsWhiteSpace(html[cnEnd]) && html[cnEnd] != '>')
                        cnEnd++;
                    if (cnEnd > cnStart)
                    {
                        var closeTagSpan = html.AsSpan(cnStart, cnEnd - cnStart);
                        if (closeTagSpan.Equals("li", StringComparison.OrdinalIgnoreCase))
                        {
                            if (skipPInLi.Count > 0)
                                skipPInLi.RemoveAt(skipPInLi.Count - 1); // 退出 <li>
                        }
                    }

                    sb.Append(html, pos, tagClose - pos + 1);
                    pos = tagClose + 1;
                    continue;
                }

                // 找标签名
                int nameStart = lt + 1;
                int nameEnd = nameStart;
                while (nameEnd < html.Length && !char.IsWhiteSpace(html[nameEnd]) && html[nameEnd] != '>')
                    nameEnd++;
                if (nameEnd <= nameStart || nameEnd > html.Length)
                {
                    sb.Append(html, pos, tagClose - pos + 1);
                    pos = tagClose + 1;
                    continue;
                }

                var tagSpan = html.AsSpan(nameStart, nameEnd - nameStart);

                // <li> 标签：在任何深度都分配 data-block（嵌套列表独立 block）
                bool isListItem = tagSpan.Equals("li", StringComparison.OrdinalIgnoreCase);
                if (isListItem) skipPInLi.Add(true); // 进入 <li>，期待首 <p> 是列表项文本

                // 跳过 <li> 内自动生成的首个 <p>（pandoc 将列表项文本包裹在 <p> 中）
                // 后续的 <p>（如子列表后的普通段落）则正常分配 data-block
                bool isP = tagSpan.Length == 1 && (tagSpan[0] == 'p' || tagSpan[0] == 'P');
                bool maySkipP = isP && skipPInLi.Count > 0 && skipPInLi[skipPInLi.Count - 1];
                bool skipPInsideLi = maySkipP;
                // 遇到 <p> 且跳过后，或遇到非 <li> 非 <p> 的标签时，将栈顶设为 false
                // 注意：不能在 <li> 自身时设 false，否则刚 push 的 true 立即失效
                if (!isListItem && skipPInLi.Count > 0 && skipPInLi[skipPInLi.Count - 1])
                {
                    skipPInLi[skipPInLi.Count - 1] = false;
                }

                // 块级标签（不含 ul/ol）在任何深度都分配 data-block，
                // 使 <li> 内的 <table>/<figure> 等也可独立定位；但 <li> 内的 <p> 跳过
                if ((CheckIsBlockTag(tagSpan) && !skipPInsideLi) || isListItem)
                {
                    // h1-h6 是标题，递增 headingOrdinal 并重置 blockOrdinal
                    bool isHeading = tagSpan.Length == 2 && (tagSpan[0] == 'h' || tagSpan[0] == 'H') &&
                                     tagSpan[1] >= '1' && tagSpan[1] <= '6';
                    if (isHeading)
                    {
                        headingOrdinal++;
                        blockOrdinal = -1;
                        // 标题不计入 block 计数（C# 端 blockOrdinal 只对内容块编号，标题用 -1）
                        sb.Append(html, pos, tagClose - pos);
                        sb.Append(" data-block=\"-1\" data-heading=\"").Append(headingOrdinal).Append('"');
                        sb.Append('>');
                    }
                    else
                    {
                        blockOrdinal++;
                        sb.Append(html, pos, tagClose - pos);
                        sb.Append(" data-block=\"").Append(blockOrdinal).Append("\" data-heading=\"").Append(headingOrdinal).Append('"');
                        sb.Append('>');
                    }
                }
                else
                {
                    sb.Append(html, pos, tagClose - pos + 1);
                }

                pos = tagClose + 1;
            }

            return sb.ToString();
        }

        /// <summary>
        /// 将 &lt;p&gt; 内 &lt;br&gt; 分割的每一行包裹在 &lt;span data-line="N"&gt; 中，
        /// 使导航左→右时可以精确定位到单行高亮。
        /// </summary>
        public static string WrapLineSpans(string html)
        {
            if (html.AsSpan().IndexOf("<br", StringComparison.OrdinalIgnoreCase) < 0)
                return html;

            var sb = new StringBuilder(html.Length + 200);
            int i = 0;

            while (i < html.Length)
            {
                // 找下一个 <p……> 标签
                int pTagStart = html.IndexOf("<p", i, StringComparison.OrdinalIgnoreCase);
                if (pTagStart < 0) { sb.Append(html, i, html.Length - i); break; }

                int pTagEnd = html.IndexOf('>', pTagStart);
                if (pTagEnd < 0) { sb.Append(html, i, html.Length - i); break; }

                // 检查是否可能包含 <br（快速过掉无 br 的 <p>）
                int pCloseTag = html.IndexOf("</p>", pTagEnd, StringComparison.OrdinalIgnoreCase);
                if (pCloseTag < 0) { sb.Append(html, i, html.Length - i); break; }

                int contentStart = pTagEnd + 1;
                int contentLen = pCloseTag - contentStart;
                if (contentLen <= 0)
                {
                    sb.Append(html, i, pCloseTag - i + 4);
                    i = pCloseTag + 4;
                    continue;
                }

                // 用 span 检查 inner html 是否包含 <br（避免 Substring 分配）
                var contentSpan = html.AsSpan(contentStart, contentLen);
                if (contentSpan.IndexOf("<br".AsSpan(), StringComparison.OrdinalIgnoreCase) < 0)
                {
                    sb.Append(html, i, pCloseTag - i + 4);
                    i = pCloseTag + 4;
                    continue;
                }

                // 有这个 <p> 有 <br>，处理它
                sb.Append(html, i, pTagEnd - i + 1); // 从上次位置到 <p> 结束

                // 去掉 contentSpan 首尾空白（直接在 span 上操作，不分配）
                int innerStart = 0;
                while (innerStart < contentSpan.Length && (contentSpan[innerStart] == '\n' || contentSpan[innerStart] == '\r' || contentSpan[innerStart] == ' '))
                    innerStart++;
                int innerEnd = contentSpan.Length;
                while (innerEnd > innerStart && (contentSpan[innerEnd - 1] == '\n' || contentSpan[innerEnd - 1] == '\r' || contentSpan[innerEnd - 1] == ' '))
                    innerEnd--;
                if (innerStart >= innerEnd) { sb.Append("</p>"); i = pCloseTag + 4; continue; }
                var innerSpan = contentSpan[innerStart..innerEnd];

                int lineIdx = 0;
                int pos = 0;
                while (pos < innerSpan.Length)
                {
                    // 用 IndexOf 替代逐字符扫描 <br
                    int brPos = innerSpan.Slice(pos).IndexOf("<br".AsSpan(), StringComparison.OrdinalIgnoreCase);
                    if (brPos >= 0) brPos += pos; else brPos = -1;

                    if (brPos < 0)
                    {
                        // 最后一段
                        sb.Append("<span data-line=\"").Append(lineIdx).Append("\">");
                        sb.Append(innerSpan[pos..]);
                        sb.Append("</span>");
                        break;
                    }

                    // 这段文本
                    sb.Append("<span data-line=\"").Append(lineIdx).Append("\">");
                    sb.Append(innerSpan.Slice(pos, brPos - pos));
                    sb.Append("</span>");

                    // 保留原来的 <br.../> 不变（在原始 html 中定位）
                    int absBrPos = contentStart + innerStart + brPos;
                    int brEnd = html.IndexOf('>', absBrPos);
                    if (brEnd < 0) { sb.Append(html, absBrPos, html.Length - absBrPos); break; }
                    sb.Append(html, absBrPos, brEnd - absBrPos + 1);
                    // 跳过 <br> 后的空白（\n, \r, 空格等）
                    int afterBr = brEnd + 1 - contentStart - innerStart;
                    pos = afterBr;
                    while (pos < innerSpan.Length && (innerSpan[pos] == '\n' || innerSpan[pos] == '\r' || innerSpan[pos] == ' '))
                        pos++;
                    lineIdx++;
                }

                sb.Append("</p>");
                i = pCloseTag + 4;
            }

            return sb.ToString();
        }

        private static string ProcessImageCaptions(string html)
        {
            if (html.AsSpan().IndexOf("<img") < 0) return html;
            return ImgWithAltRegex().Replace(html, m =>
            {
                var alt = m.Groups[1].Value;
                if (string.IsNullOrWhiteSpace(alt)) return m.Value;

                var textBefore = html.AsSpan(0, m.Index);
                var lastFigureOpen = textBefore.LastIndexOf("<figure".AsSpan());
                var lastFigureClose = textBefore.LastIndexOf("</figure>".AsSpan());
                if (lastFigureOpen > lastFigureClose)
                    return m.Value;

                return string.Concat("<figure>", m.Value, "<figcaption>", EscapeHtml(alt), "</figcaption></figure>");
            });
        }

        private static bool CheckIsBlockTag(ReadOnlySpan<char> tag)
        {
            if (tag.Length == 0 || tag.Length > 10) return false;
            // 快速路径：小写标签直接比较（避免分配）
            switch (tag.Length)
            {
                case 1: return tag[0] is 'p' or 'P';
                case 2:
                    char f = char.ToLowerInvariant(tag[0]), s = char.ToLowerInvariant(tag[1]);
                    return (f == 'h' && s is >= '1' and <= '6') ||
                           (f == 'h' && s == 'r') || (f == 'l' && s == 'i');
                case 3:
                    return tag.Equals("pre", StringComparison.OrdinalIgnoreCase) ||
                           tag.Equals("div", StringComparison.OrdinalIgnoreCase);
                case 6:
                    return tag.Equals("figure", StringComparison.OrdinalIgnoreCase);
                case 10:
                    return tag.Equals("blockquote", StringComparison.OrdinalIgnoreCase);
                default: return false;
            }
        }

        /// <summary>转义 HTML 特殊字符，防止 XSS</summary>
        private static string EscapeHtml(string text)
        {
            if (text.IndexOfAny(['&', '<', '>', '"', '\'']) < 0) return text;
            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;")
                       .Replace("\"", "&quot;").Replace("'", "&#39;");
        }
    }
}
