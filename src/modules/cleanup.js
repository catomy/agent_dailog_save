const { log } = require('./utils.js');

/**
 * 修复链接，将相对路径转换为绝对路径
 * @param {HTMLElement} clonedRoot - 克隆的根元素
 */
function fixLinks(clonedRoot) {
  try {
    let links = [];
    try {
      links = clonedRoot.querySelectorAll('a');
    } catch (error) {
      log(`警告: 获取链接元素失败 - ${error.message}`);
      return;
    }
    
    links.forEach(a => {
      try {
        // 转换为绝对路径
        if (a.href) {
          a.href = a.href;
        }
        // 强制样式，防止被清洗
        a.style.color = 'blue';
        a.style.textDecoration = 'underline';
      } catch (error) {
        log(`警告: 修复链接失败 - ${error.message}`);
      }
    });
  } catch (error) {
    log(`错误: 修复链接失败 - ${error.message}`);
  }
}

/**
 * 清理克隆的文档，移除不需要的元素
 * @param {HTMLElement} clonedRoot - 克隆的根元素
 */
function cleanClone(clonedRoot) {
  try {
    // 移除脚本、按钮、输入框等
    try {
      const trash = clonedRoot.querySelectorAll('script, style, button, input, textarea, noscript, iframe, link[rel="stylesheet"], link[as="script"], link[rel="preload"], link[rel="modulepreload"], svg:not([data-docx-id])');
      trash.forEach(el => {
        try {
          el.remove();
        } catch (error) {
          log(`警告: 移除垃圾元素失败 - ${error.message}`);
        }
      });
    } catch (error) {
      log(`警告: 选择垃圾元素失败 - ${error.message}`);
    }
    
    // 移除 hidden 元素
    try {
      const all = clonedRoot.querySelectorAll('*');
      all.forEach(el => {
        try {
          if (el.style.display === 'none' || el.style.visibility === 'hidden' || el.getAttribute('aria-hidden') === 'true') {
              el.remove();
              return;
          }
          
          // 移除事件处理器属性
          const attrs = el.attributes;
          if (attrs) {
              for (let i = attrs.length - 1; i >= 0; i--) {
              try {
                const name = attrs[i].name;
                if (name.startsWith('on') || name === 'src' && el.tagName === 'SCRIPT') {
                    el.removeAttribute(name);
                }
              } catch (error) {
                log(`警告: 移除属性失败 - ${error.message}`);
              }
              }
          }
        } catch (error) {
          log(`警告: 处理元素失败 - ${error.message}`);
        }
      });
    } catch (error) {
      log(`警告: 获取所有元素失败 - ${error.message}`);
    }

    // 再次确保没有任何 script 标签（防止 querySelectorAll 遗漏动态插入的）
    try {
      const scripts = clonedRoot.getElementsByTagName('script');
      while (scripts.length > 0) {
        try {
          scripts[0].parentNode.removeChild(scripts[0]);
        } catch (error) {
          log(`警告: 移除脚本标签失败 - ${error.message}`);
          break;
        }
      }
    } catch (error) {
      log(`警告: 获取脚本标签失败 - ${error.message}`);
    }

    // 深度清理空白元素（递归多次以处理嵌套空白）
    try {
      let removed;
      let maxIterations = 10; // 防止无限循环
      let iterations = 0;
      do {
        removed = false;
        iterations++;
        if (iterations > maxIterations) {
          log('警告: 清理空白元素达到最大迭代次数');
          break;
        }
        // 选择所有空的 div, span, p, section, aside 等容器元素
        try {
          const emptyCandidates = clonedRoot.querySelectorAll('div, span, p, section, article, aside, nav, header, footer');
          emptyCandidates.forEach(el => {
            try {
              // 如果没有子节点，或者子节点全是空白文本
              if (!el.hasChildNodes() || (el.childNodes.length === 1 && el.childNodes[0].nodeType === 3 && !el.childNodes[0].textContent.trim())) {
                  // 排除 img, br, hr 等自闭合标签
                  // 但 querySelectorAll 选中的都是容器，所以通常安全
                  // 额外检查是否包含 img
                  if (!el.querySelector('img')) {
                      el.remove();
                      removed = true;
                  }
              }
            } catch (error) {
              log(`警告: 清理空白元素失败 - ${error.message}`);
            }
          });
        } catch (error) {
          log(`警告: 选择空白元素失败 - ${error.message}`);
          break;
        }
      } while (removed); // 如果有移除，可能产生新的父级空白，继续循环
    } catch (error) {
      log(`警告: 深度清理空白元素失败 - ${error.message}`);
    }

    // 移除多余的换行 (br)
    try {
      const brs = clonedRoot.querySelectorAll('br');
      brs.forEach(br => {
        try {
          // 如果 br 的下一个兄弟也是 br，或者 br 是父元素的最后一个子元素，移除
          // 简单策略：连续的 br 只保留一个
          if (br.nextElementSibling && br.nextElementSibling.tagName === 'BR') {
              br.remove();
          }
        } catch (error) {
          log(`警告: 移除多余换行失败 - ${error.message}`);
        }
      });
    } catch (error) {
      log(`警告: 选择换行元素失败 - ${error.message}`);
    }
  } catch (error) {
    log(`错误: 清理克隆文档失败 - ${error.message}`);
  }
}

module.exports = {
  fixLinks,
  cleanClone
};
