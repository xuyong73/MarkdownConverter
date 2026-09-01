namespace MarkdownConverter;

/// <summary>
/// 基于 ## 标题的分段增量转换缓存。
/// 将 Markdown 按 ## 行首标题拆分为独立分段，仅重转换内容变化的段。
/// </summary>
internal class IncrementalConverter
{
    private const int MaxCachedSections = 5000;

    private List<string>? _cachedSections;
    private List<string>? _cachedHtml;
    private List<(int Start, int End)>? _cachedRanges;

    private string? CachedMarkdown
    {
        get;
        set
        {
            if (field == value) return;
            field = value;
            _cachedRanges = null;
        }
    }

    /// <returns>
    /// sections: 所有分段内容列表。
    /// changedIndices: null = 首次或需全量转换；空列表 = 无变化。
    /// </returns>
    public (List<string> Sections, List<int>? ChangedIndices) Analyze(string rawMarkdown)
    {
        var sections = SplitByH2(rawMarkdown);

        if (_cachedSections == null || sections.Count != _cachedSections.Count)
        {
            return (sections, null);
        }

        List<int>? changed = null;
        for (int i = 0; i < sections.Count; i++)
        {
            if (!string.Equals(sections[i], _cachedSections[i], StringComparison.Ordinal))
                (changed ??= []).Add(i);
        }

        if (changed == null)
            return (sections, []);

        // 仅当所有分段都变化时才退化到全量转换
        if (changed.Count >= sections.Count)
        {
            _cachedSections = sections;
            _cachedHtml = null;
            return (sections, null);
        }

        for (int i = 0; i < sections.Count; i++)
            _cachedSections[i] = sections[i];

        return (sections, changed);
    }

    /// <summary>缓存全量转换结果，超过阈值时放弃缓存避免内存膨胀</summary>
    public void CacheFullResult(List<string> sections, List<string> htmlParts)
    {
        if (sections.Count > MaxCachedSections)
        {
            Clear();
            return;
        }
        _cachedSections = sections;
        _cachedHtml = htmlParts;
    }

    /// <summary>获取缓存的 HTML 分段</summary>
    public List<string>? GetCachedHtml() => _cachedHtml;

    /// <summary>当分段数量变化时，尝试通过内容精确匹配来复用旧缓存的 HTML</summary>
    /// <returns>映射表 [新索引] = 旧索引（-1 表示无匹配），若无可复用的段则返回 null</returns>
    public List<int>? TryMatchAfterReshuffle(List<string> newSections)
    {
        if (_cachedSections == null || _cachedHtml == null) return null;
        if (newSections.Count == _cachedSections.Count) return null;

        // 建立旧段内容→索引映射（相同内容取最后一个索引）
        var oldMap = new Dictionary<string, int>(_cachedSections.Count);
        for (int i = 0; i < _cachedSections.Count; i++)
            oldMap[_cachedSections[i]] = i;

        var reuseMap = new List<int>(newSections.Count);
        var used = new HashSet<int>();

        foreach (var newSection in newSections)
        {
            if (oldMap.TryGetValue(newSection, out var oldIdx) && used.Add(oldIdx))
                reuseMap.Add(oldIdx);
            else
                reuseMap.Add(-1);
        }

        if (reuseMap.TrueForAll(idx => idx < 0))
            return null;

        return reuseMap;
    }

    public void Clear()
    {
        _cachedSections = null;
        _cachedHtml = null;
        CachedMarkdown = null;
        _cachedRanges = null;
    }

        public List<(int Start, int End)> GetSectionCharRanges(string markdown)
        {
            if (CachedMarkdown == markdown && _cachedRanges != null)
                return _cachedRanges;
            var ranges = ComputeSectionCharRanges(markdown);
            _cachedRanges = ranges;
            CachedMarkdown = markdown;
            return ranges;
        }

        private static List<(int Start, int End)> ComputeSectionCharRanges(string markdown)
        {
            List<(int Start, int End)> ranges = [];
            var span = markdown.AsSpan().TrimEnd();
            if (span.IsEmpty) return ranges;

            int sectionStart = -1;
            int i = 0;
            bool inCodeFence = false;

            while (i < span.Length)
            {
                // 代码围栏：逐行扫描，确保 inCodeFence 状态正确切换
                if (span[i] == '`' && i + 2 < span.Length && span[i + 1] == '`' && span[i + 2] == '`')
                {
                    inCodeFence = !inCodeFence;
                    // 跳过当前行（开/闭围栏行）
                    var nl = span[i..].IndexOf('\n');
                    i += nl >= 0 ? nl + 1 : span.Length;
                    continue;
                }

                // 行首检测 ## 标题
                if (!inCodeFence)
                {
                    // 快速定位到下一个换行（行首位置）
                    int lineStart = i;
                    // 跳过行首空白
                    int j = lineStart;
                    while (j < span.Length && span[j] == ' ') j++;

                    if (j + 2 < span.Length && span[j] == '#' && span[j + 1] == '#' && span[j + 2] == ' ')
                    {
                        if (sectionStart >= 0)
                        {
                            var sectionContent = span[sectionStart..i].TrimEnd();
                            if (!sectionContent.IsEmpty)
                                ranges.Add((sectionStart, i));
                        }
                        else
                        {
                            var preamble = span[..i].TrimEnd();
                            if (!preamble.IsEmpty)
                                ranges.Add((0, i));
                        }
                        sectionStart = i;
                        var nl = span[i..].IndexOf('\n');
                        i += nl >= 0 ? nl + 1 : span.Length;
                        continue;
                    }
                }

                // 快速跳到下一行
                var nextNl = span[i..].IndexOf('\n');
                if (nextNl >= 0)
                    i += nextNl + 1;
                else
                    i = span.Length;
            }

            if (sectionStart >= 0)
            {
                var lastSection = span[sectionStart..].TrimEnd();
                if (!lastSection.IsEmpty)
                    ranges.Add((sectionStart, span.Length));
            }
            else if (span.Length > 0)
            {
                // 无 ## 标题：整个内容作为一个 preamble 段
                ranges.Add((0, span.Length));
            }
            return ranges;
        }

    /// <summary>按 ## 行首标题拆分 Markdown，跳过代码围栏内的内容</summary>
    private List<string> SplitByH2(string markdown)
    {
        var ranges = GetSectionCharRanges(markdown);
        if (ranges is [_, ..])
            return ranges.ConvertAll(r => markdown[r.Start..r.End].TrimEnd().ToString());

        // 无 ## 标题：全部内容作为一个分段
        var trimmed = markdown.AsSpan().TrimEnd();
        return trimmed.IsEmpty ? [] : [trimmed.ToString()];
    }
}