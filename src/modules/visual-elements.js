const { log, isElementVisible } = require('./utils.js');

/**
 * 捕获可视化元素（公式和图标）
 * @returns {Promise<Array>} 处理后的可视化元素数组
 */
async function captureVisualElements() {
  try {
    // 1. 数学公式容器
    const mathSelectors = [
      '.MathJax', '.MathJax_Display', 'mjx-container', '.katex', '.katex-display'
    ];
    
    // 2. 常见图标容器 (FontAwesome, Material Icons, Glyphicons 等)
    const iconSelectors = [
      '.fa', '.fas', '.far', '.fal', '.fab', // FontAwesome
      '.material-icons', '.material-icons-outlined', // Google Material
      '.glyphicon', // Bootstrap
      '.icon', '.iconfont', // 通用
      'svg:not([data-docx-id])' // 页面上独立的 SVG (排除已经被处理过的)
    ];

    const allSelectors = [...mathSelectors, ...iconSelectors].join(',');
    let candidates = [];
    try {
      candidates = Array.from(document.querySelectorAll(allSelectors));
    } catch (error) {
      log(`警告: 选择器查询失败 - ${error.message}`);
      return [];
    }
    
    // 过滤掉不可见或极小的元素
    const visibleElements = candidates.filter(isElementVisible);
    
    const results = [];
    
    log(`发现 ${visibleElements.length} 个公式和图标，正在处理...`);

    for (let i = 0; i < visibleElements.length; i++) {
      try {
        const element = visibleElements[i];
        const id = element.getAttribute('data-docx-id');
        if (!id) continue;

        if (element.tagName.toLowerCase() === 'svg') {
           // SVG 直接处理
           try {
             const svgData = new XMLSerializer().serializeToString(element);
             const base64 = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
             results.push({ id, dataUrl: base64, width: element.offsetWidth, height: element.offsetHeight });
           } catch (error) {
             log(`警告: SVG 处理失败 - ${error.message}`);
           }
        } else {
           // 检查内部是否有 SVG
           try {
             const innerSvg = element.querySelector('svg');
             if (innerSvg) {
               try {
                  const svgData = new XMLSerializer().serializeToString(innerSvg);
                  const base64 = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
                  results.push({ id, dataUrl: base64, width: element.offsetWidth, height: element.offsetHeight });
               } catch (error) {
                 log(`警告: 内部 SVG 处理失败 - ${error.message}`);
               }
             } else {
               // 字体图标：尝试用 Canvas "画" 出这个字符
               try {
                 const style = window.getComputedStyle(element, '::before');
                 const content = style.content; // e.g. "\f007"
                 
                 // 只有当有伪元素内容且不是 none 时才处理
                 if (content && content !== 'none' && content !== '""') {
                   const cleanContent = content.replace(/['"]/g, '');
                   if (cleanContent) {
                       const canvas = document.createElement('canvas');
                       const size = Math.max(element.offsetWidth, element.offsetHeight, 16); // 最小 16px
                       canvas.width = size;
                       canvas.height = size;
                       const ctx = canvas.getContext('2d');
                       
                       // 复制字体样式
                       const fontSize = style.fontSize || '16px';
                       const fontFamily = style.fontFamily || 'Arial';
                       const color = style.color || '#000';
                       
                       ctx.font = `${fontSize} ${fontFamily}`;
                       ctx.fillStyle = color;
                       ctx.textAlign = 'center';
                       ctx.textBaseline = 'middle';
                       
                       // 绘制字符
                       ctx.fillText(cleanContent, size/2, size/2);
                       
                       results.push({ id, dataUrl: canvas.toDataURL('image/png'), width: size, height: size });
                       
                       // 释放临时创建的 Canvas 对象
                       canvas.width = 0;
                       canvas.height = 0;
                       ctx.clearRect(0, 0, canvas.width, canvas.height);
                   }
                 }
               } catch (error) {
                 log(`警告: 字体图标处理失败 - ${error.message}`);
               }
             }
           } catch (error) {
             log(`警告: 元素处理失败 - ${error.message}`);
           }
        }
      } catch (error) {
        log(`警告: 处理可视化元素失败 - ${error.message}`);
      }
    }

    return results;
  } catch (error) {
    log(`错误: 可视化元素处理失败 - ${error.message}`);
    return [];
  }
}

/**
 * 应用公式图片到克隆的文档
 * @param {HTMLElement} clonedRoot - 克隆的根元素
 * @param {Array} mathImages - 公式图片数组
 */
function applyMathImages(clonedRoot, mathImages) {
  try {
    mathImages.forEach(item => {
      try {
        const target = clonedRoot.querySelector(`[data-docx-id="${item.id}"]`);
        if (target) {
          const img = document.createElement('img');
          img.src = item.dataUrl;
          // 转换为 pt 或保留 px，Word 处理 px 还可以
          img.width = item.width;
          img.height = item.height;
          img.style.verticalAlign = 'middle';
          
          target.parentNode.replaceChild(img, target);
        }
      } catch (error) {
        log(`警告: 应用公式图片失败 - ${error.message}`);
      }
    });
  } catch (error) {
    log(`错误: 应用公式图片失败 - ${error.message}`);
  }
}

module.exports = {
  captureVisualElements,
  applyMathImages
};
