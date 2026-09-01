using System.Text.RegularExpressions;

namespace MarkdownConverter
{
    public static partial class ImagePathProcessor
    {
        [GeneratedRegex(@"<img\s+[^>]*?src\s*=""([^""]*)""[^>]*>", RegexOptions.IgnoreCase)]
        private static partial Regex TagRegex();

        public static string ProcessHtmlImagePaths(string html, string workDir)
        {
            if (string.IsNullOrEmpty(workDir) || string.IsNullOrEmpty(html)) return html;
            if (html.AsSpan().IndexOf("<img") < 0) return html;
            return TagRegex().Replace(html, m =>
            {
                var src = m.Groups[1].Value;
                if (string.IsNullOrEmpty(src)) return m.Value;
                // 允许的协议前缀
                if (src.StartsWith("https://", StringComparison.OrdinalIgnoreCase) ||
                    src.StartsWith("http://markdown.local/", StringComparison.OrdinalIgnoreCase))
                    return m.Value;
                // http:// 非本地 → 保留原样
                if (src.StartsWith("http://", StringComparison.OrdinalIgnoreCase)) return m.Value;
                // data: URI → 保留原样（base64 内嵌图片）
                if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return m.Value;
                // 相对路径：规范化并检查路径遍历
                var normalized = src.Replace('\\', '/').TrimStart('/');
                if (IsPathTraversal(normalized))
                    return m.Value;
                return m.Value.Replace(src, $"http://markdown.local/{normalized}");
            });
        }

        /// <summary>检测路径遍历攻击，包括 URL 编码变体</summary>
        private static bool IsPathTraversal(string path)
        {
            // 迭代解码，防止双重/多重编码绕过
            var decoded = path;
            for (int i = 0; i < 3; i++)
            {
                var next = Uri.UnescapeDataString(decoded);
                if (next == decoded) break;
                decoded = next;
            }
            return ContainsTraversalPattern(path) || ContainsTraversalPattern(decoded);
        }

        private static bool ContainsTraversalPattern(string path)
        {
            // 标准化斜杠方向后检查
            var normalized = path.Replace('\\', '/');
            return normalized.Contains("/../") || normalized.StartsWith("../") || normalized.EndsWith("/..") || normalized == "..";
        }
    }
}
