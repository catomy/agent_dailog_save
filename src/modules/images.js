const { log, cleanupMemory, delay } = require('./utils.js');

/**
 * 处理网页图片
 * @returns {Promise<Array>} 处理后的图片数组
 */
async function processImages() {
  try {
    let images = [];
    try {
      images = Array.from(document.querySelectorAll('img'));
    } catch (error) {
      log(`警告: 获取图片元素失败 - ${error.message}`);
      return [];
    }
    
    const results = [];
    let successCount = 0;
    let errorCount = 0;
    
    // 内存监控和批量大小动态调整
    let batchSize = 5;
    const adjustBatchSize = () => {
      try {
        if (navigator.deviceMemory) {
          // 根据设备内存调整批量大小
          const memory = navigator.deviceMemory;
          if (memory < 4) batchSize = 3;
          else if (memory < 8) batchSize = 5;
          else batchSize = 8;
        }
        
        // 根据图片数量调整批量大小
        if (images.length > 50) batchSize = Math.min(batchSize, 3);
      } catch (error) {
        log(`警告: 调整批量大小失败 - ${error.message}`);
      }
    };
    
    adjustBatchSize();
    
    // 图片压缩函数
    const compressImage = (canvas, quality = 0.8) => {
      try {
        // 对于大图片进行压缩
        const maxWidth = 1920;
        const maxHeight = 1080;
        
        let { width, height } = canvas;
        
        if (width > maxWidth || height > maxHeight) {
          const ratio = Math.min(maxWidth / width, maxHeight / height);
          width *= ratio;
          height *= ratio;
          
          const tempCanvas = document.createElement('canvas');
          tempCanvas.width = width;
          tempCanvas.height = height;
          const tempCtx = tempCanvas.getContext('2d');
          tempCtx.drawImage(canvas, 0, 0, width, height);
          return tempCanvas.toDataURL('image/jpeg', quality);
        }
        
        // 对于小图片使用 PNG 格式
        return canvas.toDataURL('image/png');
      } catch (error) {
        log(`警告: 图片压缩失败 - ${error.message}`);
        return canvas.toDataURL('image/png');
      }
    };
    
    const convertImage = async (img) => {
      try {
        // 1. 获取真实的图片 URL (处理懒加载)
        // 优先检查 data-src, data-original, data-url 等常见懒加载属性
        let src = img.src;
        const lazyAttrs = ['data-src', 'data-original', 'data-original-src', 'data-url', 'data-lazy-src'];
        
        // 如果 src 不存在，或者 src 是 base64 占位符（通常很短），或者是 1x1 像素点
        const isPlaceholder = !src || (src.startsWith('data:') && src.length < 2000) || src.includes('spacer.gif');
        
        if (isPlaceholder) {
          for (const attr of lazyAttrs) {
            try {
              const val = img.getAttribute(attr);
              if (val) {
                src = val;
                // 如果是相对路径，转绝对路径
                if (!src.startsWith('http') && !src.startsWith('data:')) {
                   const a = document.createElement('a');
                   a.href = src;
                   src = a.href;
                   // 释放临时创建的 a 元素
                   a.remove();
                }
                break;
              }
            } catch (error) {
              log(`警告: 获取懒加载属性失败 - ${error.message}`);
            }
          }
        }

        if (!src) return;
        if (src.startsWith('data:')) {
            // 如果已经是高清的 base64 (长度够长)，直接保存
            if (src.length > 2000) {
               const id = img.getAttribute('data-docx-id');
               results.push({ id, dataUrl: src });
               successCount++;
            }
            return;
        }

        const id = img.getAttribute('data-docx-id');
        
        // 方案 A: 尝试使用 Canvas 转换 (速度最快，且能处理一部分格式问题)
        // 需要创建一个新的 Image 对象以避免污染页面上的元素或处理跨域属性
        const newImg = new Image();
        newImg.crossOrigin = "Anonymous";
        
        try {
            await new Promise((resolve, reject) => {
                newImg.onload = resolve;
                newImg.onerror = reject;
                newImg.src = src;
                // 超时保护 8s
                const timeoutId = setTimeout(() => reject(new Error('Image load timeout')), 8000);
                // 清理超时器
                newImg.onload = () => {
                    clearTimeout(timeoutId);
                    resolve();
                };
                newImg.onerror = () => {
                    clearTimeout(timeoutId);
                    reject(new Error('Image load error'));
                };
            });

            const canvas = document.createElement('canvas');
            canvas.width = newImg.naturalWidth;
            canvas.height = newImg.naturalHeight;
            const ctx = canvas.getContext('2d');
            ctx.drawImage(newImg, 0, 0);
            
            // 使用压缩函数减少内存占用
            const dataUrl = compressImage(canvas);
            results.push({ id, dataUrl });
            successCount++;
            
            // 释放临时创建的对象
            canvas.width = 0;
            canvas.height = 0;
            // 清除 canvas 内容
            ctx.clearRect(0, 0, canvas.width, canvas.height);
        } catch (canvasError) {
            // Canvas 被污染 (Tainted) 或加载失败
            // 方案 B: 尝试使用 Fetch (如果 CSP 允许)
            // 增加 no-referrer 策略以绕过防盗链
            try {
                const fetchImage = async (url, options = {}) => {
                    const controller = new AbortController();
                    const timeoutId = setTimeout(() => controller.abort(), 8000);
                    try {
                        const res = await fetch(url, { 
                            ...options, 
                            signal: controller.signal,
                            credentials: 'omit' // 不发送 Cookie
                        });
                        clearTimeout(timeoutId);
                        if (!res.ok) throw new Error(`HTTP ${res.status}`);
                        return await res.blob();
                    } catch (error) {
                        clearTimeout(timeoutId);
                        throw error;
                    }
                };

                // 第一次尝试：无 Referrer
                let blob;
                try {
                    blob = await fetchImage(src, { referrerPolicy: 'no-referrer' });
                } catch (error) {
                    // 第二次尝试：默认策略 (有些图片可能需要 Referrer)
                    blob = await fetchImage(src);
                }

                const reader = new FileReader();
                const base64 = await new Promise(resolve => {
                    reader.onloadend = () => resolve(reader.result);
                    reader.readAsDataURL(blob);
                });
                results.push({ id, dataUrl: base64 });
                successCount++;
                
                // 释放 blob 对象
                blob = null;
            } catch (fetchError) {
                errorCount++;
                log(`警告: 图片加载失败 - ${fetchError.message}`);
            }
        } finally {
            // 确保释放 newImg 对象
            if (newImg) {
                newImg.src = '';
            }
        }
      } catch (error) {
        errorCount++;
        log(`警告: 图片处理失败 - ${error.message}`);
      }
    };

    // 批量处理，动态调整批量大小
    for (let i = 0; i < images.length; i += batchSize) {
      try {
        // 每处理10张图片后重新评估批量大小
        if (i > 0 && i % 10 === 0) {
          adjustBatchSize();
        }
        
        const chunk = images.slice(i, i + batchSize);
        await Promise.all(chunk.map(convertImage));
        
        // 内存清理：定期清理不再需要的对象
        if (i % 20 === 0) {
          // 强制垃圾回收（如果可用）
          cleanupMemory();
        }
        
        if (i % 10 === 0) log(`处理图片: ${i}/${images.length}`);
      } catch (error) {
        log(`警告: 批量处理图片失败 - ${error.message}`);
      }
    }
    
    // 计算成功率
    const successRate = images.length > 0 ? (successCount / images.length * 100).toFixed(1) : 100;
    log(`图片处理完成，成功率: ${successRate}%`);
    if (errorCount > 0) {
      log(`图片处理过程中遇到 ${errorCount} 个错误`);
    }
    
    return results;
  } catch (error) {
    log(`错误: 图片处理失败 - ${error.message}`);
    return [];
  }
}

/**
 * 应用图片替换到克隆的文档
 * @param {HTMLElement} clonedRoot - 克隆的根元素
 * @param {Array} docImages - 处理后的图片数组
 */
function applyDocImages(clonedRoot, docImages) {
  try {
    docImages.forEach(item => {
      try {
        const target = clonedRoot.querySelector(`img[data-docx-id="${item.id}"]`);
        if (target) {
          target.src = item.dataUrl;
          // 移除 srcset 避免干扰
          target.removeAttribute('srcset');
          target.removeAttribute('loading'); // 移除懒加载属性
        }
      } catch (error) {
        log(`警告: 应用图片替换失败 - ${error.message}`);
      }
    });
  } catch (error) {
    log(`错误: 应用图片替换失败 - ${error.message}`);
  }
}

/**
 * 清理未处理的图片
 * @param {HTMLElement} clonedRoot - 克隆的根元素
 */
function cleanUnprocessedImages(clonedRoot) {
  try {
    let images = [];
    try {
      images = clonedRoot.querySelectorAll('img');
    } catch (error) {
      log(`警告: 获取图片元素失败 - ${error.message}`);
      return;
    }
    
    let removedCount = 0;
    images.forEach(img => {
      try {
        // 检查 src 是否存在且是否为 base64
        if (!img.src || !img.src.startsWith('data:')) {
           // 如果有 alt 文本，替换为文本说明，否则直接移除
           if (img.alt && img.alt.trim()) {
             const span = document.createElement('span');
             span.textContent = ` [图: ${img.alt}] `;
             span.style.color = '#666';
             span.style.fontSize = '0.8em';
             img.parentNode.replaceChild(span, img);
           } else {
             img.remove();
           }
           removedCount++;
        }
      } catch (error) {
        log(`警告: 清理图片失败 - ${error.message}`);
      }
    });
    if (removedCount > 0) {
      log(`已清理 ${removedCount} 张无法转换的图片以确保导出成功。`);
    }
  } catch (error) {
    log(`错误: 清理未处理图片失败 - ${error.message}`);
  }
}

module.exports = {
  processImages,
  applyDocImages,
  cleanUnprocessedImages
};
