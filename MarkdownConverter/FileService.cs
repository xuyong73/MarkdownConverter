using System.IO;
using System.Text;

namespace MarkdownConverter
{
    internal static class FileService
    {
        private const string PreviewPrefix = "MDConv_preview_";

        /// <summary>判断文件是否为 Markdown 文件</summary>
        public static bool IsMarkdownFile(string path) =>
            Path.GetExtension(path) is ".md" or ".markdown";

        public static async Task<(string text, string workDir, string fileName)> LoadFileAsync(string filePath)
        {
            var text = await File.ReadAllTextAsync(filePath, Encoding.UTF8);
            var workDir = Path.GetDirectoryName(filePath) ?? "";
            var fileName = Path.GetFileName(filePath);
            return (text, workDir, fileName);
        }

        public static async Task SaveMarkdownAsync(string filePath, string content)
        {
            await File.WriteAllTextAsync(filePath, content, Encoding.UTF8);
        }

        /// <summary>创建临时 HTML 文件并返回路径，供 WebView2 导航使用</summary>
        public static async Task<string> WriteTempHtmlAsync(string html)
        {
            var tempFile = Path.Combine(Path.GetTempPath(), $"{PreviewPrefix}{Guid.NewGuid():n}.html");
            await File.WriteAllTextAsync(tempFile, html, Encoding.UTF8);
            return tempFile;
        }

        /// <summary>删除临时文件，忽略失败</summary>
        public static void TryDeleteFile(string path)
        {
            try { File.Delete(path); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"清理临时文件失败: {ex.Message}"); }
        }

        public static async Task<bool> ExportWordAsync(string filePath, string markdown, string? workDir, bool darkTheme)
        {
            var htmlContent = await PandocConverter.ConvertToHtmlAsync(markdown);
            var fullHtml = HtmlGenerator.GenerateHtmlContent(htmlContent, darkTheme);

            var tempHtmlFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.html");
            try
            {
                await File.WriteAllTextAsync(tempHtmlFile, fullHtml, Encoding.UTF8);
                return await PandocConverter.ConvertFileAsync(tempHtmlFile, filePath, ["-f", "html", "-t", "docx"]);
            }
            finally
            {
                TryDeleteFile(tempHtmlFile);
            }
        }

        public static async Task SaveHtmlAsync(string filePath, string markdown, bool darkTheme)
        {
            var htmlContent = await PandocConverter.ConvertToHtmlAsync(markdown);
            var finalHtml = HtmlGenerator.GenerateHtmlContent(htmlContent, darkTheme);
            await File.WriteAllTextAsync(filePath, finalHtml, Encoding.UTF8);
        }

        public static void CleanupTempFiles()
        {
            try
            {
                foreach (var f in Directory.EnumerateFiles(Path.GetTempPath(), $"{PreviewPrefix}*.html"))
                    TryDeleteFile(f);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"CleanupTempFiles 异常: {ex.Message}"); }
        }
    }
}
