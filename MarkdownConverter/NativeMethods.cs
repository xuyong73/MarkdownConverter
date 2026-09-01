using System.Runtime.InteropServices;

namespace MarkdownConverter;

/// <summary>原生 Win32 API 方法</summary>
internal static class NativeMethods
{
    internal const byte VK_CONTROL = 0x11;
    internal const byte VK_F = 0x46;
    internal const uint KEYEVENTF_KEYUP = 0x0002;

    /// <summary>模拟键盘按键事件（已过时但可靠，适用于触发 WebView2 原生快捷键）</summary>
    [DllImport("user32.dll")]
    internal static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, nint dwExtraInfo);
}