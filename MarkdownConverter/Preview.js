// MarkdownConverter Preview.js — 预览 DOM 工具 + 结构路径导航 + 点击上报
// 合并自 PreviewCommon.js / PreviewNavigation.js / PreviewClick.js

var MDConv = MDConv || {};

// ═══════════════════════════════════════════════════════════════
//  DOM 工具函数
// ═══════════════════════════════════════════════════════════════
MDConv.isVisible = function(el) {
    var rect = el.getBoundingClientRect();
    // 部分可见即返回 true（与视口有交集）
    return rect.top < window.innerHeight && rect.bottom > 0 &&
        rect.left < window.innerWidth && rect.right > 0;
};

/** 当前高亮元素（用于清除） */
MDConv._highlightEl = null;

/**
 * 长条状高亮：纯背景色（与左侧编辑器选中色一致）
 */
MDConv.highlight = function(el) {
    // 清除上次高亮
    if (MDConv._highlightEl && MDConv._highlightEl !== el) {
        // 如果上次高亮是 <pre> 内的行级 span，需要拆包清理
        if (MDConv._highlightEl.hasAttribute('data-pre-line-highlight')) {
            var pre = MDConv._highlightEl.closest('pre');
            if (pre) cleanupPreLineHighlight(pre);
        } else {
            MDConv._highlightEl.style.background = '';
            MDConv._highlightEl.style.borderLeft = '';
            MDConv._highlightEl.style.boxShadow = '';
        }
    }

    var r = el.tagName === 'TD' || el.tagName === 'TH' ? (el.closest('tr') || el) : el;

    r.style.background = 'rgba(0,102,204,0.35)';

    MDConv._highlightEl = r;

    // 清除文本选中
    var sel = window.getSelection();
    if (sel) sel.removeAllRanges();

    if (!r.hasAttribute('tabindex')) r.setAttribute('tabindex', '-1');
    r.focus({preventScroll: true});
    r.style.outline = 'none';

    window._pvcProgSel = true;

    if (!MDConv.isVisible(r))
        r.scrollIntoView({behavior: 'smooth', block: 'center'});
};

MDConv.getRoot = function(sectionIdx) {
    return (sectionIdx >= 0 && sectionIdx !== null)
        ? document.querySelector('[data-section-idx="' + sectionIdx + '"]') || document
        : document;
};

// ═══════════════════════════════════════════════════════════════
//  结构路径导航（C# → JS）
//  使用 (sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal) 四元组定位
// ═══════════════════════════════════════════════════════════════

/**
 * 按结构路径导航到预览中的目标位置
 * 直接使用 AddBlockOrdinals 生成的 data-block / data-heading 属性查找，
 * 彻底消除 C#/JS 两侧计数不一致。
 * @param {number} sectionIdx    段索引
 * @param {number} headingOrdinal 段内标题序号（-1 = 无标题区域/preamble）
 * @param {number} blockOrdinal   标题后块序号（0-based，-1 = 标题本身）
 * @param {number} lineOrdinal    块内行序号（0-based）
 * @returns {boolean} 是否导航成功
 */
function nav_byPath(sectionIdx, headingOrdinal, blockOrdinal, lineOrdinal) {
    var root = MDConv.getRoot(sectionIdx);
    if (!root) return false;

    // 定位到标题本身
    if (blockOrdinal === -1) {
        var target = root.querySelector('[data-heading="' + headingOrdinal + '"][data-block="-1"]');
        if (target) { MDConv.highlight(target); return true; }
        return false;
    }

    // 用 data-heading + data-block 属性精确查找目标块
    var target = root.querySelector('[data-heading="' + headingOrdinal + '"][data-block="' + blockOrdinal + '"]');
    if (!target) {
        // 降级：尝试在 preamble 区域查找（headingOrdinal = -1）
        target = root.querySelector('[data-heading="-1"][data-block="' + blockOrdinal + '"]');
    }
    if (!target) return false;

    return selectLineInBlock(target, lineOrdinal);
}

/**
 * 在 <pre> 元素中高亮指定行
 * 通过 Range API 精确定位目标行文本，包裹为 span 后高亮，支持左→右导航和右→左点击
 * @param {HTMLElement} preEl       <pre> 元素
 * @param {number}     lineOrdinal  行号（0-based）
 * @returns {boolean} 是否成功
 */
function highlightPreLine(preEl, lineOrdinal) {
    if (!preEl || preEl.tagName !== 'PRE') return false;

    // 清除之前在同一个 <pre> 内创建的行高亮 span（避免 DOM 污染）
    cleanupPreLineHighlight(preEl);

    var text = preEl.textContent;
    var lines = text.split('\n');

    if (lineOrdinal < 0 || lineOrdinal >= lines.length) {
        MDConv.highlight(preEl); // 降级：整块高亮
        return true;
    }

    // 计算目标行的字符偏移范围 [startOffset, endOffset)
    var startOffset = 0;
    for (var i = 0; i < lineOrdinal; i++) {
        startOffset += lines[i].length + 1; // +1 for \n
    }
    var endOffset = startOffset + lines[lineOrdinal].length;

    // 用 TreeWalker 遍历文本节点，映射字符偏移到 DOM 位置
    var range = document.createRange();
    var walker = document.createTreeWalker(preEl, NodeFilter.SHOW_TEXT, null, false);

    var currentOffset = 0;
    var startNode = null, startNodeOffset = 0;
    var endNode = null, endNodeOffset = 0;

    while (walker.nextNode()) {
        var node = walker.currentNode;
        var len = node.nodeValue.length;

        if (startNode === null && currentOffset + len > startOffset) {
            startNode = node;
            startNodeOffset = startOffset - currentOffset;
        }
        if (currentOffset + len >= endOffset) {
            endNode = node;
            endNodeOffset = endOffset - currentOffset;
            break;
        }
        currentOffset += len;
    }

    if (!startNode || !endNode) {
        MDConv.highlight(preEl);
        return true;
    }

    range.setStart(startNode, startNodeOffset);
    range.setEnd(endNode, endNodeOffset);

    // 包裹目标行文本为 span，以便统一调用 MDConv.highlight
    var span = document.createElement('span');
    span.setAttribute('data-pre-line-highlight', '');
    try {
        range.surroundContents(span);
    } catch (e) {
        // surroundContents 可能因跨元素边界失败（如含子标签），降级到整块高亮
        MDConv.highlight(preEl);
        return true;
    }

    MDConv.highlight(span);
    return true;
}

/** 清除 <pre> 内之前创建的行高亮 span */
function cleanupPreLineHighlight(preEl) {
    var oldSpan = preEl.querySelector('span[data-pre-line-highlight]');
    if (oldSpan) {
        var parent = oldSpan.parentNode;
        if (parent) {
            // 将 span 的子节点移出，替换 span 本身
            while (oldSpan.firstChild) {
                parent.insertBefore(oldSpan.firstChild, oldSpan);
            }
            parent.removeChild(oldSpan);
        }
    }
}

/**
 * 在块级元素中选中指定行
 * @param {Element} el          块级元素
 * @param {number}  lineOrdinal 行序号（0-based）
 * @returns {boolean}
 */
function selectLineInBlock(el, lineOrdinal) {
    if (!el) return false;

    // 特殊处理表格行（支持 table-wrapper div 包裹的表格）
    var targetTable = el.tagName === 'TABLE' ? el : el.querySelector('table');
    if (targetTable) {
        var rows = targetTable.querySelectorAll('tr');
        if (lineOrdinal < rows.length) {
            MDConv.highlight(rows[lineOrdinal]);
            return true;
        }
        MDConv.highlight(targetTable);
        return true;
    }

    // 代码块：精确定位到行
    if (el.tagName === 'PRE') {
        return highlightPreLine(el, lineOrdinal);
    }

    // 特殊处理列表
    if (el.tagName === 'UL' || el.tagName === 'OL') {
        var items = el.querySelectorAll('li');
        if (lineOrdinal < items.length) {
            MDConv.highlight(items[lineOrdinal]);
            return true;
        }
        MDConv.highlight(el);
        return true;
    }

    // 含 <br> 的段落/块引用等：查找对应的 <span data-line="N"> 精确高亮
    var lineSpan = el.querySelector('span[data-line="' + lineOrdinal + '"]');
    if (lineSpan) {
        MDConv.highlight(lineSpan);
        return true;
    }

    // 普通块：整块高亮
    MDConv.highlight(el);
    return true;
}

// ═══════════════════════════════════════════════════════════════
//  点击上报（JS → C#）
//  点击预览时从 data-block / data-heading 属性读取结构路径并发送给 C#
// ═══════════════════════════════════════════════════════════════

(function(){
if (window._pvcInit) return;
window._pvcInit = true;

// 跟踪鼠标按下位置，用于区分点击和拖拽选择
document.addEventListener('mousedown', function(e) {
    window._pvcMouseX = e.clientX;
    window._pvcMouseY = e.clientY;
});

document.addEventListener('click', function(e) {
    try {
        // 拖拽选择识别（移动超过 4px 视为拖拽）
        if (window._pvcMouseX !== undefined) {
            var dx = e.clientX - window._pvcMouseX;
            var dy = e.clientY - window._pvcMouseY;
            if (dx * dx + dy * dy > 16) {
                window._pvcMouseX = undefined;
                return;
            }
            window._pvcMouseX = undefined;
        }

        // 清除上次编程选中
        var sel = window.getSelection();
        if (window._pvcProgSel) {
            if (sel) sel.removeAllRanges();
            window._pvcProgSel = false;
        }
        if (sel && sel.toString().trim().length > 0) {
            return; // 用户手动选中文本，不触发导航
        }

        // 找到点击的块级元素
        var blockQuery = 'p,h1,h2,h3,h4,h5,h6,li,pre,blockquote,td,th,figcaption,figure,div';
        var targetEl = e.target.closest(blockQuery + ',a');
        if (!targetEl) return;

        // 链接特殊处理：不触发导航
        if (targetEl.tagName === 'A') return;

        // 1. 获取段索引
        var sectionDiv = targetEl.closest('[data-section-idx]') || document.body;
        var sectionIdx = parseInt(sectionDiv.getAttribute('data-section-idx')) || 0;

        // 2. 从最近的 data-block 祖先读取结构路径
        var dataBlockEl = targetEl.closest('[data-block]');
        if (!dataBlockEl) return;
        var headingOrdinal = parseInt(dataBlockEl.getAttribute('data-heading'));
        var blockOrdinal = parseInt(dataBlockEl.getAttribute('data-block'));

        // 3. 检查是否是标题本身
        var tag = targetEl.tagName;

        var headingTag = (tag === 'H1' || tag === 'H2' || tag === 'H3' ||
                          tag === 'H4' || tag === 'H5' || tag === 'H6');
        if (headingTag) {
            window.chrome.webview.postMessage(JSON.stringify({
                sectionIdx: sectionIdx,
                headingOrdinal: headingOrdinal,
                blockOrdinal: -1,
                lineOrdinal: 0
            }));
            MDConv.highlight(targetEl);
            return;
        }

        // 4. 计算 lineOrdinal
        var lineOrdinal = 0;

        if (tag === 'PRE') {
            var caretRange = document.caretRangeFromPoint(e.clientX, e.clientY);
            if (caretRange) {
                var codeText = targetEl.textContent;
                var range = document.createRange();
                range.setStart(targetEl, 0);
                range.setEnd(caretRange.startContainer, caretRange.startOffset);
                var offset = range.toString().length;
                lineOrdinal = offset > 0 ? codeText.substring(0, offset).split('\n').length - 1 : 0;
            }
        } else if (tag === 'TD' || tag === 'TH') {
            var tr = targetEl.closest('tr');
            var targetTable = targetEl.closest('table');
            if (tr && targetTable) {
                var rows = Array.from(targetTable.querySelectorAll('tr'));
                lineOrdinal = rows.indexOf(tr);
                if (lineOrdinal < 0) lineOrdinal = 0;
            }
        } else if (tag === 'LI' || (tag === 'P' && targetEl.closest('li'))) {
            // 每个 <li> 都是独立 block，lineOrdinal 恒为 0
            lineOrdinal = 0;
        } else {
            var brs = targetEl.querySelectorAll('br');
            if (brs.length > 0) {
                var caretRange = document.caretRangeFromPoint(e.clientX, e.clientY);
                if (caretRange) {
                    var targetNode = caretRange.startContainer;
                    var brCount = 0;
                    var walker = document.createTreeWalker(targetEl, NodeFilter.SHOW_ALL, null, false);
                    while (walker.nextNode()) {
                        var node = walker.currentNode;
                        if (node === targetNode) break;
                        if (node.nodeType === 1 && node.tagName === 'BR') {
                            brCount++;
                        }
                    }
                    lineOrdinal = brCount;
                }
            }
        }

        // 5. 高亮目标（精确到行）
        if (tag === 'PRE') {
            highlightPreLine(targetEl, lineOrdinal);
        } else {
            var highlightEl = targetEl;
            if (tag !== 'LI' && tag !== 'TD' && tag !== 'TH' && tag !== 'FIGURE') {
                var lineSpan = targetEl.querySelector('span[data-line="' + lineOrdinal + '"]');
                if (lineSpan) highlightEl = lineSpan;
            }
            MDConv.highlight(highlightEl);
        }

        // 6. 发送结构路径消息
        window.chrome.webview.postMessage(JSON.stringify({
            sectionIdx: sectionIdx,
            headingOrdinal: headingOrdinal,
            blockOrdinal: blockOrdinal,
            lineOrdinal: lineOrdinal
        }));

    } catch (err) {
        // 忽略点击处理错误
    }
});

})();
