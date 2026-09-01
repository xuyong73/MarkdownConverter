using System.Collections.Concurrent;
using System.Windows.Media;

namespace MarkdownConverter
{
    /// <summary>集中管理所有颜色值，消除硬编码</summary>
    public static class ColorConstants
    {
        // ===== 状态栏颜色 =====
        public const string StatusBlue = "#3B82F6";      // 信息/进行中
        public const string StatusGreen = "#10B981";     // 成功
        public const string StatusRed = "#EF4444";       // 错误
        public const string StatusAmber = "#F59E0B";     // 警告/过程

        // ===== 主题色 — 浅色 =====
        public const string LightBodyBg = "#FFFFFF";
        public const string LightText = "#1E1E1E";
        public const string LightHeading = "#2563EB";
        public const string LightLink = "#2563EB";
        public const string LightBorder = "#eee";
        public const string LightCodeBorder = "#E9ECEF";
        public const string LightTableBorder = "#E0E0E0";
        public const string LightSplitter = "#E0E0E0";
        public const string LightEditor = "#FFFFFF";
        public const string LightBlockquoteBg = "rgba(37,99,235,0.05)";
        public const string LightBlockquoteText = "#555";
        public const string LightCodeBg = "#F8F9FA";
        public const string LightCodeText = "#C7254E";

        // ===== 主题色 — 深色 =====
        public const string DarkText = "#D4D4D4";
        public const string DarkHeading = "#93C5FD";
        public const string DarkBorder = "#444";
        public const string DarkLink = "#60A5FA";
        public const string DarkBodyBg = "#000000";
        public const string DarkMainBg = "#000000";
        public const string DarkEditor = "#000000";
        public const string DarkSplitter = "#3A3A4A";
        public const string DarkBlockquoteBg = "#2D3748";
        public const string DarkBlockquoteText = "#A0AEC0";
        public const string DarkCodeBg = "#2D2D30";
        public const string DarkCodeText = "#E2E8F0";
        public const string DarkCodeBorder = "#4A5568";
    }

    public static class ThemeManager
    {
        private static readonly ConcurrentDictionary<string, SolidColorBrush> _brushCache = new();
        private static readonly ThemeColors _light = GetLightThemeColors();
        private static readonly ThemeColors _dark = GetDarkThemeColors();

        public static ThemeColors GetColors(bool darkTheme) => darkTheme ? _dark : _light;

        public static string GetThemeJsCss(bool darkTheme)
        {
            return darkTheme
                ? "document.documentElement.classList.add('dark-theme');"
                : "document.documentElement.classList.remove('dark-theme');";
        }

        private static ThemeColors GetDarkThemeColors()
        {
            return new ThemeColors
            {
                Text = ColorConstants.DarkText,
                Heading = ColorConstants.DarkHeading,
                Border = ColorConstants.DarkBorder,
                Link = ColorConstants.DarkLink,
                BlockquoteBg = ColorConstants.DarkBlockquoteBg,
                BlockquoteText = ColorConstants.DarkBlockquoteText,
                BlockquoteBorder = ColorConstants.DarkLink,       // == Link
                CodeBg = ColorConstants.DarkCodeBg,
                CodeText = ColorConstants.DarkCodeText,
                CodeBorder = ColorConstants.DarkCodeBorder,
                TableBorder = ColorConstants.DarkBorder,          // == Border
                BodyBg = ColorConstants.DarkBodyBg,
                Main = ColorConstants.DarkMainBg,
                Editor = ColorConstants.DarkEditor,
                Splitter = ColorConstants.DarkSplitter
            };
        }

        private static ThemeColors GetLightThemeColors()
        {
            return new ThemeColors
            {
                Text = ColorConstants.LightText,
                Heading = ColorConstants.LightHeading,
                Border = ColorConstants.LightBorder,
                Link = ColorConstants.LightLink,                 // == Heading
                BlockquoteBg = ColorConstants.LightBlockquoteBg,
                BlockquoteText = ColorConstants.LightBlockquoteText,
                BlockquoteBorder = ColorConstants.LightLink,     // == Heading
                CodeBg = ColorConstants.LightCodeBg,
                CodeText = ColorConstants.LightCodeText,
                CodeBorder = ColorConstants.LightCodeBorder,
                TableBorder = ColorConstants.LightTableBorder,
                BodyBg = ColorConstants.LightBodyBg,
                Main = ColorConstants.LightBodyBg,               // == BodyBg
                Editor = ColorConstants.LightEditor,
                Splitter = ColorConstants.LightSplitter           // == TableBorder
            };
        }

        public static string GetBackgroundColor(bool darkTheme) => GetColors(darkTheme).BodyBg;

        public static SolidColorBrush GetBrush(string color) =>
            _brushCache.GetOrAdd(color, static c => new SolidColorBrush((Color)ColorConverter.ConvertFromString(c)));
    }

    public readonly record struct ThemeColors
    {
        public string Text { get; init; }
        public string Heading { get; init; }
        public string Border { get; init; }
        public string Link { get; init; }
        public string BlockquoteBg { get; init; }
        public string BlockquoteText { get; init; }
        public string BlockquoteBorder { get; init; }
        public string CodeBg { get; init; }
        public string CodeText { get; init; }
        public string CodeBorder { get; init; }
        public string TableBorder { get; init; }
        public string BodyBg { get; init; }
        public string Main { get; init; }
        public string Editor { get; init; }
        public string Splitter { get; init; }
    }
}