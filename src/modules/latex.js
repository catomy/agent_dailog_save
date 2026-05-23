const { log } = require('./utils.js');
const katex = require('katex');

// 缓存正则表达式对象，避免重复创建
const LATEX_REGEX = {
  block: /(\$\$([\s\S]+?)\$\$)|(\\\[([\s\S]+?)\\\])/g,
  inline: /(\$([^\$\n]+?)\$)|(\\\(([\s\S]+?)\\\))/g,
  blockTest: /(\$\$([\s\S]+?)\$\$)|(\\\[([\s\S]+?)\\\])/,
  inlineTest: /(\$([^\$\n]+?)\$)|(\\\(([\s\S]+?)\\\))/
};

// 缓存 katex 渲染结果，避免重复渲染相同的公式
const katexCache = new Map();

/**
 * 处理纯文本 LaTeX 公式
 * @param {HTMLElement} root - 根元素
 * @returns {number} 处理的公式数量
 */
function processRawLatex(root) {
  try {
    // 遍历所有文本节点
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
    const nodesToReplace = [];
    
    let node;
    while (node = walker.nextNode()) {
      // 跳过 script, style 等标签内的文本
      if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEXTAREA'].includes(node.parentNode.tagName)) continue;
      
      const text = node.nodeValue;
      if (!text) continue;

      // 使用缓存的正则表达式进行测试
      try {
        if (LATEX_REGEX.blockTest.test(text) || LATEX_REGEX.inlineTest.test(text)) {
          nodesToReplace.push(node);
        }
      } catch (error) {
        log(`警告: 正则表达式测试失败 - ${error.message}`);
        continue;
      }
    }

    let latexCount = 0;

    // 批量处理节点，减少 DOM 操作次数
    nodesToReplace.forEach(node => {
      try {
        const parent = node.parentNode;
        const text = node.nodeValue;
        
        // 使用缓存的正则表达式进行替换
        let newHtml = text
          .replace(LATEX_REGEX.block, (match) => {
              latexCount++;
              let raw = match.startsWith('$$') ? match.slice(2, -2) : match.slice(2, -2);
              // 检查缓存
              const cacheKey = `block_${raw}`;
              if (katexCache.has(cacheKey)) {
                return katexCache.get(cacheKey);
              }
              try {
                  const result = katex.renderToString(raw, { throwOnError: false, displayMode: true });
                  katexCache.set(cacheKey, result);
                  return result;
              } catch(error) { return match; }
          })
          .replace(LATEX_REGEX.inline, (match) => {
              latexCount++;
              let raw = match.startsWith('$') ? match.slice(1, -1) : match.slice(2, -2);
              // 检查缓存
              const cacheKey = `inline_${raw}`;
              if (katexCache.has(cacheKey)) {
                return katexCache.get(cacheKey);
              }
              try {
                  const result = katex.renderToString(raw, { throwOnError: false, displayMode: false });
                  katexCache.set(cacheKey, result);
                  return result;
              } catch(error) { return match; }
          });

        if (newHtml !== text) {
            const span = document.createElement('span');
            span.innerHTML = newHtml;
            parent.replaceChild(span, node);
        }
      } catch (error) {
        log(`警告: 处理 LaTeX 公式失败 - ${error.message}`);
      }
    });
    
    // 清理缓存，避免内存泄漏
    try {
      katexCache.clear();
    } catch (error) {
      log(`警告: 清理缓存失败 - ${error.message}`);
    }
    
    return latexCount;
  } catch (error) {
    log(`错误: LaTeX 处理失败 - ${error.message}`);
    return 0;
  }
}

module.exports = {
  processRawLatex
};
