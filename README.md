# Markdown 转换器 v3.2 (MDConv)

一款基于 WPF 的 Markdown 编辑与预览工具，支持实时转换、多格式导出、深色主题和双向定位同步。

## 功能特性

- **实时预览** — 左侧编辑 Markdown，右侧通过 WebView2 实时渲染 HTML
- **增量转换** — 基于 `##` 标题分段，只转换变化部分，大文件编辑更流畅
- **自动转换** — 编辑完成后自动触发转换（300ms 防抖），无需手动操作
- **并行转换** — 全量转换时并行处理所有分段，充分利用多核 CPU
- **分片渲染** — 使用 CSS `content-visibility` 优化大文档的渲染性能
- **高性能编辑** — 基于 AvalonEdit 的虚拟化渲染引擎，流畅处理大文本
- **多格式导出** — 支持导出为 Markdown (.md)、HTML (.html)、Word (.docx)
- **深色主题** — 一键切换亮色/深色主题
- **双向定位** — 编辑器→预览（Ctrl+Q / 右键菜单 / 双击）和预览→编辑器（点击预览元素），基于结构路径精确定位到视口中央；转换进行中若光标所在分段已渲染仍可立即定位

## 定位原理

双向导航的核心是**结构路径四元组** `(sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal)`，它将 Markdown 文档和 HTML 预览统一抽象为同一棵结构树：

```
文档
├── Section 0 (preamble / ## 标题前区域)          ← sectionIdx=0
│   ├── 块 0: <p> 段落内容                         ← blockOrdinal=0, lineOrdinal=0/1/...
│   └── 块 1: <pre> 代码块                        ← blockOrdinal=1, lineOrdinal=0/1/... (行级)
├── Section 1 (## 第一个标题)                      ← sectionIdx=1
│   ├── 标题: h2 "第一章"                          ← headingOrdinal=0, blockOrdinal=-1
│   ├── 块 0: <ul> 列表                           ← blockOrdinal=0, lineOrdinal=0/1/...
│   └── 块 1: <table> 表格                        ← blockOrdinal=1, lineOrdinal=0/1/... (行级)
└── Section 2 (## 第二个标题)                      ← sectionIdx=2
    ├── 标题: h2 "第二章"                          ← headingOrdinal=1, blockOrdinal=-1
    └── 块 0: <blockquote> 引用                    ← blockOrdinal=0
```

### 左→右：编辑器 → 预览

1. **确定段索引** — 通过 `GetSectionIndexAtCaret()` 二分查找光标所在的 `sectionIdx`（基于 `IncrementalConverter.ComputeSectionCharRanges` 的段字符范围）
2. **构建行映射** — 对该段文本调用 `BuildSectionLineMap()` 单次遍历，产出：
   - `Headings[]`：每个 `#` 标题行的字符偏移范围
   - `ContentOffsets[]` + `ContentBlockOrdinals[]` + `ContentLineOrdinals[]`：每行内容对应的块序号和块内行号
3. **计算路径** — `GetStructuralPath()` 将光标偏移映射为四元组（先查是否在标题行上，再在内容行中二分查找）
4. **JS 端执行** — C# 调用 `nav_byPath(sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal)`，JS 通过 `[data-heading][data-block]` 属性选择器直接定位目标元素，再按 `lineOrdinal` 精确高亮子元素（表格行、列表项、代码行、`<span data-line>` 等）

### 右→左：预览 → 编辑器

1. **点击捕获** — JS 监听 `click` 事件，通过 `closest('[data-block]')` 向上查找最近的带属性祖先元素，读取 `data-heading` 和 `data-block`
2. **计算 lineOrdinal** — 根据被点击元素的标签类型用不同策略计算行号：
   - `<pre>`：`caretRangeFromPoint` 计算字符偏移 → 按 `\n` 分割得行号
   - `<td>/<th>`：查找所在 `<tr>` 在表格所有行中的索引
   - `<li>`：查找在父列表 `<li>` 集合中的索引
   - 含 `<br>` 的段落：TreeWalker 统计 `<br>` 数量
3. **回传 C#** — 通过 `chrome.webview.postMessage` 发送 JSON 四元组
4. **编辑器定位** — `NavigateEditorByPath()` 反向查询 `SectionLineMap` 得到目标字符偏移，选中对应行并滚动

### 关键设计决策

- **属性注入替代计数** — `HtmlGenerator.AddBlockOrdinals()` 在 C# 端为每个顶级块级元素注入 `data-block` 和 `data-heading` 属性，JS 直接读取而非自行遍历计数，彻底消除两侧计数不一致
- **段级行映射缓存** — `_lineMapCache` 以 `sectionIdx` 为键缓存 `SectionLineMap`，同一段在增量转换间内容不变，避免重复构建
- **代码块行级高亮** — `<pre>` 内通过 Range API + TreeWalker 映射字符偏移到 DOM 文本节点位置，用 `surroundContents` 包裹目标行为 `<span data-pre-line-highlight>` 实现单行高亮，切换时自动拆包清理避免 DOM 污染累积
- **`<p>` 内行 span 包裹** — `WrapLineSpans()` 将 `<br>` 分割的每行包裹为 `<span data-line="N">`，使非代码块也能精确到行高亮

- **表格支持** — 宽表格自动横向滚动、表格行整行高亮、含数学公式的单元格定位
- **数学公式** — 支持 LaTeX 数学公式（行内 `$...$` 和块级 `$$...$$` / `\[...\]`），通过 Pandoc MathML 转换
- **任务列表** — 支持 `- [ ]` / `- [x]` 语法
- **图片支持** — 自动处理本地图片路径，支持相对路径和绝对路径、WebView2 虚拟主机映射
- **文件拖放** — 直接拖入 .md 文件到编辑器
- **命令行打开** — 支持命令行参数直接打开 .md 文件

## 系统要求

- Windows 7 及以上
- [Microsoft Edge WebView2 Runtime](https://go.microsoft.com/fwlink/?linkid=2124701)（如未安装，启动时会提示下载）
- [Pandoc](https://pandoc.org/installing.html)（用于 Markdown 转换和 Word 导出）

## 快速开始

1. 安装 Pandoc 并确保 `pandoc` 命令在系统 PATH 中
2. 运行 `MDConv.exe`
3. 在左侧编辑器中输入或粘贴 Markdown 文本（自动实时预览）
4. 点击 **保存文件** 导出结果

### 命令行用法

```shell
MDConv.exe path/to/document.md
```

## 构建

```shell
dotnet build -c Release
```

输出路径：`bin\Release\net10.0-windows7.0\MDConv.dll`

## 技术栈

- **框架**: .NET 10 / WPF
- **编辑器**: AvalonEdit (ICSharpCode.AvalonEdit)
- **Markdown 渲染**: Pandoc CLI
- **预览引擎**: WebView2 (Microsoft Edge)
- **语言**: C#

## 项目结构

```
MarkdownConverter/
├── MainWindow.xaml           # 主窗口布局
├── MainWindow.xaml.cs        # 主窗口入口、状态管理、生命周期
├── MainWindow.Navigation.cs  # 双向导航：结构路径计算与编辑器↔预览定位
├── MainWindow.Conversion.cs  # 转换逻辑：增量/全量转换、WebView2 初始化与渲染
├── MainWindow.Editor.cs      # 编辑器操作：右键菜单、快捷键、文件操作、主题切换
├── Preview.js                # 预览端 JS（DOM 工具 + 结构路径导航 + 点击上报）
├── HtmlGenerator.cs          # HTML 生成、CSS 主题注入、块级属性注入、行 span 包裹
├── IncrementalConverter.cs   # 基于 ## 标题的分段增量转换缓存
├── ImagePathProcessor.cs     # 图片路径处理与安全检查
├── PandocConverter.cs        # Pandoc CLI 封装
├── ThemeManager.cs           # 主题颜色管理与 Brush 缓存
├── FileService.cs            # 文件读写、临时文件管理、导出
├── App.xaml.cs               # 应用程序入口
└── MarkdownConverter.csproj  # 项目配置
```

## 版本记录

### v3.2 (2026-06-06)

- **修复点击预览全选问题** — `FindMultiLineFormulaSpan` 中闭合 `$`（前有 `}` `)` `]` 等括号字符）不再被误判为公式开头，避免与后续 `$` 配对成虚假的跨行公式范围导致左侧全选
- **版本号集中管理** — 标题栏版本号改为运行时从程序集读取，仅需修改 `.csproj` 中的 `<Version>` 一处，不再四处硬编码
- **新增 `\${}\$` 闭合字符保护** — `TryFindSingleDollarCrossLine` 增加 `}` `)` `]` 前导检查，防止 LaTeX 闭合 `$` 被当作公式开头

### v3.1 (2026-06-05)

- **代码块行级高亮** — `<pre><code>` 块内支持精确到单行的高亮定位，不再整块高亮。左侧双击代码块中某行（如 `"MarkdownConverter/"`），右侧精确定位并高亮该行；右侧点击代码块内某行，同样仅高亮该行而非整个代码块。通过 Range API + TreeWalker 映射字符偏移到 DOM 位置，用 `surroundContents` 包裹目标行为 `<span data-pre-line-highlight>` 实现行级高亮，切换时自动拆包清理避免 DOM 污染累积
- **结构路径导航体系** — 彻底重构双向导航为纯结构路径 `(sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal)` 四元组定位，C# 端通过 `SectionLineMap` 单次遍历构建段内行映射（标题行 + 内容行），JS 端通过 `AddBlockOrdinals` 注入的 `data-block`/`data-heading` 属性直接读取，消除 C#/JS 两端计数不一致的根因
- **HTML 端属性注入** — `AddBlockOrdinals` 为所有顶级块级元素自动添加 `data-block="N"` 和 `data-heading="M"` 属性，JS 导航无需自行计算块序号，标题/段落/列表/表格/代码块/引用等统一走同一查找路径
- **段内行映射缓存** — `_lineMapCache` 以 sectionIdx 为键缓存 `SectionLineMap`，同一段内容在转换间不变，避免重复构建
- **代码现代化** — 元组析构声明（IDE0042）、集合初始化简写 `new()` → `[]`（IDE0028）

### v3.0 (2026-05-31)

- **AvalonEdit 替换 WPF TextBox** — 虚拟化渲染引擎，大文本编辑不再卡顿
- **Section 级联定位** — 左侧→右侧定位精确到光标所在分段，避免跨区误匹配，定位精度与右侧→左侧一致
- **视口居中定位** — 定位时目标元素滚动到预览区中央，视觉效果更好
- **转换中智能定位** — 转换进行时，若光标所在分段已渲染完成，仍可立即定位（双击/右键/Ctrl+Q），无需等待全部分段完成
- **编辑时自动转换** — 编辑完成后自动触发转换（300ms 防抖），无需手动点击转换按钮
- **DispatcherTimer 防抖** — 用 DispatcherTimer 替代 Task.Run+Task.Delay，消除线程池调度与 UI 切换开销
- **内存分配优化** — 修复 LastIndexOf 子串分配、span.ToString() 冗余、SectionCharRanges 缓存、LineMap 重复获取等性能瓶颈
- **死代码清理** — 移除 CandidateExists、StringBuilderPool、冗余 FindName 调用等无效代码
- 清理冗余代码：移除 `AssemblyConfiguration` 硬编码、遗留注释；统一 heading 提取逻辑（消除 4 处重复）

### v2.9 (2026-05-28)

- 新增双向定位：编辑器→预览（Ctrl+Q）和预览→编辑器（点击预览），基于文本模糊匹配（v3.1 已重构为结构路径体系）
- 表格定位全面修复：宽表格横向滚动、表格行整行高亮、含数学公式表格单元格双向定位
- 修复列表项和任务列表的定位问题
- 有序/无序列表项精确定位
- 公式定位支持：行内公式 `$...$` 和行间公式 `$$...$$` 双向定位
- 图片定位支持：点击/定位到图片
- 清理无效代码，简化定位逻辑

### v2.8 (2026-05-25)

- 合并预处理为一次遍历：将任务列表规范化、行间距调整、数学块保护、尖括号转义合并到单次字符遍历，消除冗余正则扫描
- 转换后自动定位到光标所在行

### v2.7

- 增量转换：基于 `##` 标题分段，仅重转换变化的分段
- 并行转换：全量转换时并行处理全部分段，充分利用多核 CPU
- 分片渲染：使用 CSS `content-visibility` 优化大文档渲染性能
- 修复大文件右侧预览不显示和显示慢的问题
- 修复数学公式中 `>` 被错误转义的问题
- 修复含数学公式行的定位失效问题
- 转换后自动定位到光标所在行

### v2.6

- 修复 WebView2 的 `NavigateToString` 大小限制（改用文件导航）
- 新增任务列表 `- [ ]` / `- [x]` 支持
- 实现 Brush 缓存，优化主题切换性能
- 图片路径快速路径优化

### v2.5

- 新增图片路径处理（相对路径 / 绝对路径 / 虚拟主机映射）
- 新增文件拖放支持
- 新增命令行参数打开文件
- 新增 `figcaption` 图片标题支持

### v2.4

- 新增 LaTeX 数学公式支持（`$...$` / `$$...$$` / `\[...\]`）
- 语法高亮和代码块样式优化

### v2.3

- 新增"定位到预览"功能：右键菜单中定位光标所在行到预览区
- 修复大文件编辑时右侧预览跑位问题

### v2.2

- 新增 Word (.docx) 导出支持
- 保存对话框支持多格式选择

### v2.1

- 新增深色/亮色主题一键切换
- 主题颜色管理和缓存

### v2.0

- 初始版本：Markdown 编辑、实时预览、HTML 导出
- WebView2 预览引擎集成
- Pandoc 转换管道
