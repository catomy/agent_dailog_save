const { log, delay } = require('./utils.js');

/**
 * 内联 CSS 样式到克隆的文档
 * @param {HTMLElement} realRoot - 真实的根元素
 * @param {HTMLElement} clonedRoot - 克隆的根元素
 * @returns {Promise} 样式内联完成的Promise
 */
async function inlineStyles(realRoot, clonedRoot) {
  try {
    // 遍历 Clone 节点，去 Real 中找
    let allCloned = [];
    try {
      allCloned = clonedRoot.querySelectorAll('*');
    } catch (error) {
      log(`警告: 获取克隆元素失败 - ${error.message}`);
      return;
    }
    
    // 批量处理，避免阻塞
    const CHUNK_SIZE = 100;
    let processedCount = 0;
    
    for (let i = 0; i < allCloned.length; i++) {
      try {
        const cloneEl = allCloned[i];
        const id = cloneEl.getAttribute('data-docx-id');
        if (!id) continue;

        // 优化：每 100 个元素让出主线程
        if (i % CHUNK_SIZE === 0) {
          await delay(0);
        }

        let realEl;
        try {
          realEl = document.querySelector(`[data-docx-id="${id}"]`);
        } catch (error) {
          log(`警告: 查找真实元素失败 - ${error.message}`);
          continue;
        }
        
        if (realEl) {
          let computed;
          try {
            computed = window.getComputedStyle(realEl);
          } catch (error) {
            log(`警告: 获取计算样式失败 - ${error.message}`);
            continue;
          }
          
          // 只保留关键样式，减少体积
          const stylesToCopy = [
            'color', 'background-color', 
            'font-size', 'font-family', 'font-weight', 'font-style',
            'text-align', 'text-decoration',
            // 'padding-left', 'margin-left', 'margin-bottom', // 移除缩进和段间距，避免留白过多
            'border', 'display'
          ];
          
          // 对于 img，保留尺寸
          if (realEl.tagName === 'IMG') {
            stylesToCopy.push('width', 'height');
            try {
              cloneEl.setAttribute('width', realEl.width);
              cloneEl.setAttribute('height', realEl.height);
            } catch (error) {
              log(`警告: 设置图片尺寸失败 - ${error.message}`);
            }
          }

          let styleStr = '';
          stylesToCopy.forEach(prop => {
            try {
              const val = computed.getPropertyValue(prop);
              if (val && val !== 'rgba(0, 0, 0, 0)' && val !== 'transparent' && val !== 'auto' && val !== 'normal') {
                 styleStr += `${prop}:${val};`;
              }
            } catch (error) {
              log(`警告: 获取样式属性失败 - ${error.message}`);
            }
          });
          
          try {
            cloneEl.setAttribute('style', styleStr);
            processedCount++;
          } catch (error) {
            log(`警告: 设置样式失败 - ${error.message}`);
          }
        }
      } catch (error) {
        log(`警告: 处理元素样式失败 - ${error.message}`);
      }
    }
    
    log(`成功内联 ${processedCount} 个元素的样式`);
  } catch (error) {
    log(`错误: 内联样式失败 - ${error.message}`);
  }
}

module.exports = {
  inlineStyles
};
