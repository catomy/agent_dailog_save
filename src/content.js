const htmlDocx = require('html-docx-js/dist/html-docx');
const html2canvas = require('html2canvas');
const { saveAs } = require('file-saver');
const katex = require('katex');

const CTX_ID = chrome.runtime.id;
const GLOBAL_FLAG = `hasRunContentScript_${CTX_ID}`;
const STATE_KEY = '__wordExportState';

if (!window[STATE_KEY]) {
  window[STATE_KEY] = {
    isExporting: false,
  };
}

const exportState = window[STATE_KEY];

// 始终注册监听器，因为我们将在 popup 中控制注入时机
// 为了避免旧代码残留（虽然很难完全避免），我们只注册一次核心逻辑
// 但由于无法移除旧监听器，我们通过简单的全局变量来防止重复执行业务逻辑
if (!window[GLOBAL_FLAG]) {
  setupListener();
  window[GLOBAL_FLAG] = true;

  // 仅在首次初始化时通知 background，避免重复注入导致的无意义状态抖动
  chrome.runtime.sendMessage({ action: 'engine_ready' });
}
// 为了语义准确，我们在 background 中处理 'export_start' 或 'ready' 都可以
// 这里简单起见，既然用户说“加载以后”，我们理解为页面加载完成或插件注入完成。
// 由于我们是按需注入，所以只有点击后注入才会变亮。
// 如果用户希望“只要打开网页，插件如果是可用的，就亮”，那需要在 manifest 中配置 content_scripts 自动注入。
// 但我们目前是 lazy injection。
// 妥协方案：在 popup 点击注入后，图标变亮，表示“正在工作”。
// 或者：我们在 manifest 中定义 content_scripts 匹配 <all_urls>，只为了点亮图标？这会增加内存。
// 根据当前架构（按需注入），“加载以后”指的是“点击插件开始工作后”。
// 所以代码逻辑是：开始导出 -> 图标变亮 -> 导出结束 -> 图标变回默认（或保持亮色）。

function log(msg) {
  chrome.runtime.sendMessage({ action: 'log', message: msg });
  console.log('[WordExport]', msg);
}

function setupListener() {
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    // 1. Ping 响应
    if (request.action === 'ping') {
      sendResponse({ status: 'pong' });
      return;
    }

    // 2. 导出逻辑
    if (request.action === 'start_export') {
      // 检查是否已经在运行
      if (exportState.isExporting) {
        log('正在进行中，请稍候...');
        sendResponse({ status: 'busy' });
        return;
      }
      exportState.isExporting = true;

      chrome.runtime.sendMessage({ action: 'export_start' });
      
      runExport(request.config)
        .catch(err => {
          console.error(err);
          chrome.runtime.sendMessage({ action: 'export_error', message: err.toString() });
        })
        .finally(() => {
          exportState.isExporting = false;
        });

      sendResponse({ status: 'started' });
      
      // 异步响应需要返回 true，但这里我们主要通过 sendMessage 回传状态，所以不需要
    }
  });
}

async function runExport(config) {
  log('开始处理...');
  
  // 1. 标记 DOM 节点，以便后续映射
  log('正在索引页面元素...');
  let nodeIdCounter = 0;
  const allElements = document.querySelectorAll('*');
  allElements.forEach(el => {
    el.setAttribute('data-docx-id', nodeIdCounter++);
  });

  // 2. 自动滚动
  if (config.autoScroll) {
    log('正在滚动加载内容...');
    await autoScrollPage();
  }

  // 3. 处理 MathJax 公式和 SVG 图标
  log('正在可视化处理公式和图标...');
  const mathImages = await captureVisualElements();

  // 3.5. 处理纯文本 LaTeX 公式
  log('正在扫描并渲染纯文本 LaTeX...');
  processRawLatex(document.body);

  // 4. 处理常规图片和图标 (转 Base64 以支持离线/Word)
  log('正在处理网页图片和图标...');
  const docImages = await processImages();

  // 5. 克隆页面
  log('正在创建文档副本...');
  const clonedBody = document.body.cloneNode(true);

  // 6. 内联样式 (最耗时步骤)
  log('正在内联 CSS 样式 (保留颜色、字体)...');
  await inlineStyles(document.body, clonedBody);

  // 7. 替换公式为图片
  log('正在应用公式图片...');
  applyMathImages(clonedBody, mathImages);

  // 8. 应用常规图片替换 (替换为 Base64)
  applyDocImages(clonedBody, docImages);

  // 9. 清理未成功转换的图片 (防止 html-docx-js 报错)
  // html-docx-js 如果遇到非 data: 的 src 且无法下载，会抛出 Unable to download 错误
  // 我们必须移除所有残留的外部链接图片
  cleanUnprocessedImages(clonedBody);

  // 10. 修复链接 (转为绝对路径)
  fixLinks(clonedBody);

  // 10. 清理垃圾元素
  cleanClone(clonedBody);

  // 11. 生成 Word
  log('正在生成 Word 文档...');
  const contentHtml = `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
        <style>
          body { font-family: 'SimSun', 'Arial', sans-serif; }
        </style>
      </head>
      <body>
        ${clonedBody.innerHTML}
      </body>
    </html>
  `;

  // html-docx-js 需要 Buffer 或 ArrayBuffer，但在浏览器端它接受字符串并返回 Blob
  const converted = htmlDocx.asBlob(contentHtml, {
    orientation: 'portrait',
    margins: { top: 720, bottom: 720, left: 720, right: 720 } // twips
  });

  saveAs(converted, `Page_Export_${Date.now()}.docx`);
  
  log('完成！');
  chrome.runtime.sendMessage({ action: 'export_done' });

  // 清理 ID (可选，为了性能暂时不清理，刷新页面即可)
}

async function autoScrollPage() {
  return new Promise(resolve => {
    let totalHeight = 0;
    const distance = 100;
    const timer = setInterval(() => {
      const scrollHeight = document.body.scrollHeight;
      window.scrollBy(0, distance);
      totalHeight += distance;

      if (totalHeight >= scrollHeight || (window.innerHeight + window.scrollY) >= scrollHeight) {
        clearInterval(timer);
        window.scrollTo(0, 0); // 回到顶部
        setTimeout(resolve, 500);
      }
    }, 20); // 速度快一点
  });
}

function processRawLatex(root) {
  // 遍历所有文本节点
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null, false);
  const nodesToReplace = [];
  
  let node;
  while (node = walker.nextNode()) {
    // 跳过 script, style 等标签内的文本
    if (['SCRIPT', 'STYLE', 'NOSCRIPT', 'TEXTAREA'].includes(node.parentNode.tagName)) continue;
    
    // 简单的 LaTeX 识别正则
    // 匹配 $...$ 或 \(...\) 或 $$...$$ 或 \[...\]
    // 注意：这里只是简单匹配，可能会误伤（比如 $100），但对于大部分技术文档是有效的
    // 改进正则：要求 $ 后面不是数字，或者长度大于一定值
    const text = node.nodeValue;
    if (!text) continue;

    // 匹配块级公式 $$...$$ 或 \[...\]
    const blockRegex = /(\$\$([\s\S]+?)\$\$)|(\\\[([\s\S]+?)\\\])/g;
    // 匹配行内公式 $...$ 或 \(...\)
    const inlineRegex = /(\$([^\$\n]+?)\$)|(\\\(([\s\S]+?)\\\))/g;

    if (blockRegex.test(text) || inlineRegex.test(text)) {
      nodesToReplace.push(node);
    }
  }

  // 替换节点
  nodesToReplace.forEach(node => {
    const parent = node.parentNode;
    const text = node.nodeValue;
    
    // 我们需要将文本分割成 "普通文本" + "公式" + "普通文本"
    // 简单的解析器：逐字符扫描或正则 split
    // 这里使用正则 split 可能会比较复杂，因为有多种模式
    // 我们采用一种简单的替换策略：将文本替换为 span 容器，内部包含 text 和 img
    
    // 正则替换
    let newHtml = text
      .replace(/(\$\$([\s\S]+?)\$\$)|(\\\[([\s\S]+?)\\\])/g, (match) => {
          let raw = match.startsWith('$$') ? match.slice(2, -2) : match.slice(2, -2);
          try {
              return katex.renderToString(raw, { throwOnError: false, displayMode: true });
          } catch(e) { return match; }
      })
      .replace(/(\$([^\$\n]+?)\$)|(\\\(([\s\S]+?)\\\))/g, (match) => {
          // 排除 $100 这种情况：如果 $ 后面紧跟数字或空格，可能不是公式
          // 简单判断：如果 match 包含空格且长度较短，或者是纯数字..
          // 这里不做过于复杂的判断，信任 katex
          let raw = match.startsWith('$') ? match.slice(1, -1) : match.slice(2, -2);
          try {
              return katex.renderToString(raw, { throwOnError: false, displayMode: false });
          } catch(e) { return match; }
      });

    if (newHtml !== text) {
        const span = document.createElement('span');
        span.innerHTML = newHtml;
        parent.replaceChild(span, node);
    }
  });
}

async function captureVisualElements() {
  // 1. 数学公式容器
  const mathSelectors = [
    '.MathJax', '.MathJax_Display', 'mjx-container', '.katex', '.katex-display'
  ];
  
  // 2. 常见图标容器 (FontAwesome, Material Icons, Glyphicons 等)
  // 这些通常是 i, span 标签，通过 ::before 显示内容，或者包含 svg
  const iconSelectors = [
    '.fa', '.fas', '.far', '.fal', '.fab', // FontAwesome
    '.material-icons', '.material-icons-outlined', // Google Material
    '.glyphicon', // Bootstrap
    '.icon', '.iconfont', // 通用
    'svg:not([data-docx-id])' // 页面上独立的 SVG (排除已经被处理过的)
  ];

  const allSelectors = [...mathSelectors, ...iconSelectors].join(',');
  const candidates = Array.from(document.querySelectorAll(allSelectors));
  
  // 过滤掉不可见或极小的元素
  const visibleEls = candidates.filter(el => el.offsetWidth > 5 && el.offsetHeight > 5);
  
  const results = [];
  
  log(`发现 ${visibleEls.length} 个公式和图标，正在处理...`);

  // 辅助函数：将元素转换为 Canvas 图片 (光栅化)
  // 对于 SVG，优先尝试序列化；对于字体图标，必须用 html2canvas 或类似技术
  // 但 html2canvas 有 CSP 问题。
  // 策略：
  // 1. 如果是 SVG 标签 -> 序列化为 Base64 SVG (Word 支持)
  // 2. 如果是字体图标 -> 尝试用 Canvas 绘制文字 (需获取 Computed Style)
  
  for (let i = 0; i < visibleEls.length; i++) {
    const el = visibleEls[i];
    const id = el.getAttribute('data-docx-id');
    if (!id) continue;

    if (el.tagName.toLowerCase() === 'svg') {
       // SVG 直接处理
       try {
         const svgData = new XMLSerializer().serializeToString(el);
         const base64 = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
         results.push({ id, dataUrl: base64, width: el.offsetWidth, height: el.offsetHeight });
       } catch (e) {}
    } else {
       // 检查内部是否有 SVG
       const innerSvg = el.querySelector('svg');
       if (innerSvg) {
         try {
            const svgData = new XMLSerializer().serializeToString(innerSvg);
            const base64 = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgData)));
            results.push({ id, dataUrl: base64, width: el.offsetWidth, height: el.offsetHeight });
         } catch (e) {}
       } else {
         // 字体图标：这是一个难点。如果不截图，很难在 Word 里还原。
         // 我们尝试用 Canvas "画" 出这个字符
         try {
           const style = window.getComputedStyle(el, '::before');
           const content = style.content; // e.g. "\f007"
           
           // 只有当有伪元素内容且不是 none 时才处理
           if (content && content !== 'none' && content !== '""') {
             const cleanContent = content.replace(/['"]/g, '');
             if (cleanContent) {
                 const canvas = document.createElement('canvas');
                 const size = Math.max(el.offsetWidth, el.offsetHeight, 16); // 最小 16px
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
             }
           }
         } catch (e) {}
       }
    }
  }

  return results;
}

function applyMathImages(clonedRoot, mathImages) {
  mathImages.forEach(item => {
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
  });
}

async function processImages() {
  const imgs = Array.from(document.querySelectorAll('img'));
  const results = [];
  
  const convertImg = async (img) => {
    try {
      // 1. 获取真实的图片 URL (处理懒加载)
      // 优先检查 data-src, data-original, data-url 等常见懒加载属性
      let src = img.src;
      const lazyAttrs = ['data-src', 'data-original', 'data-original-src', 'data-url', 'data-lazy-src'];
      
      // 如果 src 不存在，或者 src 是 base64 占位符（通常很短），或者是 1x1 像素点
      // 简单的判断：如果 src 包含 "data:image" 且长度小于 2000，可能只是占位符
      const isPlaceholder = !src || (src.startsWith('data:') && src.length < 2000) || src.includes('spacer.gif');
      
      if (isPlaceholder) {
        for (const attr of lazyAttrs) {
          const val = img.getAttribute(attr);
          if (val) {
            src = val;
            // 如果是相对路径，转绝对路径
            if (!src.startsWith('http') && !src.startsWith('data:')) {
               const a = document.createElement('a');
               a.href = src;
               src = a.href;
            }
            break;
          }
        }
      }

      if (!src) return;
      if (src.startsWith('data:')) {
          // 如果已经是高清的 base64 (长度够长)，直接保存
          if (src.length > 2000) {
             const id = img.getAttribute('data-docx-id');
             results.push({ id, dataUrl: src });
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
              setTimeout(() => reject(new Error('Image load timeout')), 8000);
          });

          const canvas = document.createElement('canvas');
          canvas.width = newImg.naturalWidth;
          canvas.height = newImg.naturalHeight;
          const ctx = canvas.getContext('2d');
          ctx.drawImage(newImg, 0, 0);
          
          const dataUrl = canvas.toDataURL('image/png');
          results.push({ id, dataUrl });
      } catch (canvasErr) {
          // Canvas 被污染 (Tainted) 或加载失败
          // 方案 B: 尝试使用 Fetch (如果 CSP 允许)
          // 增加 no-referrer 策略以绕过防盗链
          try {
              const fetchImage = async (url, options = {}) => {
                  const controller = new AbortController();
                  const id = setTimeout(() => controller.abort(), 8000);
                  try {
                      const res = await fetch(url, { 
                          ...options, 
                          signal: controller.signal,
                          credentials: 'omit' // 不发送 Cookie
                      });
                      clearTimeout(id);
                      if (!res.ok) throw new Error(`HTTP ${res.status}`);
                      return await res.blob();
                  } catch (e) {
                      clearTimeout(id);
                      throw e;
                  }
              };

              // 第一次尝试：无 Referrer
              let blob;
              try {
                  blob = await fetchImage(src, { referrerPolicy: 'no-referrer' });
              } catch (e) {
                  // 第二次尝试：默认策略 (有些图片可能需要 Referrer)
                  blob = await fetchImage(src);
              }

              const reader = new FileReader();
              const base64 = await new Promise(r => {
                  reader.onloadend = () => r(reader.result);
                  reader.readAsDataURL(blob);
              });
              results.push({ id, dataUrl: base64 });
          } catch (fetchErr) {
              // console.warn('Fetch also failed', fetchErr);
              // 如果都失败了，我们保留原 img 标签的 src (如果是 http)，让 Word 自己去尝试加载（虽然 Word 也可能失败）
              // 但为了 cleanUnprocessedImages 不误删，我们可以不做处理，或者标记一下
          }
      }
    } catch (e) {
      // console.warn('Image processing failed', e);
    }
  };

  // 批量处理
  const BATCH_SIZE = 5; // 减小并发，避免网络拥塞
  for (let i = 0; i < imgs.length; i += BATCH_SIZE) {
    const chunk = imgs.slice(i, i + BATCH_SIZE);
    await Promise.all(chunk.map(convertImg));
    if (i % 10 === 0) log(`处理图片: ${i}/${imgs.length}`);
  }
  
  return results;
}

function applyDocImages(clonedRoot, docImages) {
  docImages.forEach(item => {
    const target = clonedRoot.querySelector(`img[data-docx-id="${item.id}"]`);
    if (target) {
      target.src = item.dataUrl;
      // 移除 srcset 避免干扰
      target.removeAttribute('srcset');
      target.removeAttribute('loading'); // 移除懒加载属性
    }
  });
}

function cleanUnprocessedImages(clonedRoot) {
  const imgs = clonedRoot.querySelectorAll('img');
  let removedCount = 0;
  imgs.forEach(img => {
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
  });
  if (removedCount > 0) {
    console.log(`[WordExport] 已清理 ${removedCount} 张无法转换的图片以确保导出成功。`);
  }
}

function fixLinks(clonedRoot) {
  const links = clonedRoot.querySelectorAll('a');
  links.forEach(a => {
    // 转换为绝对路径
    if (a.href) {
      a.href = a.href; 
    }
    // 强制样式，防止被清洗
    a.style.color = 'blue';
    a.style.textDecoration = 'underline';
  });
}

async function inlineStyles(realRoot, clonedRoot) {
  // 使用 TreeWalker 遍历
  const realWalker = document.createTreeWalker(realRoot, NodeFilter.SHOW_ELEMENT);
  const cloneWalker = document.createTreeWalker(clonedRoot, NodeFilter.SHOW_ELEMENT);

  let realNode = realWalker.nextNode();
  let cloneNode = cloneWalker.nextNode();

  let count = 0;

  // 假设结构一致（因为是直接 clone 的）
  // 但为了安全，我们使用 ID 查找
  // 遍历 Clone 节点，去 Real 中找
  const allCloned = clonedRoot.querySelectorAll('*');
  
  // 批量处理，避免阻塞
  const CHUNK_SIZE = 100;
  
  for (let i = 0; i < allCloned.length; i++) {
    const cloneEl = allCloned[i];
    const id = cloneEl.getAttribute('data-docx-id');
    if (!id) continue;

    // 优化：每 100 个元素让出主线程
    if (i % CHUNK_SIZE === 0) {
      await new Promise(r => setTimeout(r, 0));
    }

    const realEl = document.querySelector(`[data-docx-id="${id}"]`);
    if (realEl) {
      const computed = window.getComputedStyle(realEl);
      
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
        cloneEl.setAttribute('width', realEl.width);
        cloneEl.setAttribute('height', realEl.height);
      }

      let styleStr = '';
      stylesToCopy.forEach(prop => {
        const val = computed.getPropertyValue(prop);
        if (val && val !== 'rgba(0, 0, 0, 0)' && val !== 'transparent' && val !== 'auto' && val !== 'normal') {
           styleStr += `${prop}:${val};`;
        }
      });
      
      cloneEl.setAttribute('style', styleStr);
    }
  }
}

function cleanClone(clonedRoot) {
  // 移除脚本、按钮、输入框等
  const trash = clonedRoot.querySelectorAll('script, style, button, input, textarea, noscript, iframe, link[rel="stylesheet"], link[as="script"], link[rel="preload"], link[rel="modulepreload"], svg:not([data-docx-id])');
  trash.forEach(el => el.remove());
  
  // 移除 hidden 元素
  const all = clonedRoot.querySelectorAll('*');
  all.forEach(el => {
    if (el.style.display === 'none' || el.style.visibility === 'hidden' || el.getAttribute('aria-hidden') === 'true') {
        el.remove();
        return;
    }
    
    // 移除事件处理器属性
    const attrs = el.attributes;
    if (attrs) {
        for (let i = attrs.length - 1; i >= 0; i--) {
        const name = attrs[i].name;
        if (name.startsWith('on') || name === 'src' && el.tagName === 'SCRIPT') {
            el.removeAttribute(name);
        }
        }
    }
  });

  // 再次确保没有任何 script 标签（防止 querySelectorAll 遗漏动态插入的）
  const scripts = clonedRoot.getElementsByTagName('script');
  while (scripts.length > 0) {
    scripts[0].parentNode.removeChild(scripts[0]);
  }

  // 深度清理空白元素（递归多次以处理嵌套空白）
  let removed;
  do {
    removed = false;
    // 选择所有空的 div, span, p, section, aside 等容器元素
    const emptyCandidates = clonedRoot.querySelectorAll('div, span, p, section, article, aside, nav, header, footer');
    emptyCandidates.forEach(el => {
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
    });
  } while (removed); // 如果有移除，可能产生新的父级空白，继续循环

  // 移除多余的换行 (br)
  const brs = clonedRoot.querySelectorAll('br');
  brs.forEach(br => {
      // 如果 br 的下一个兄弟也是 br，或者 br 是父元素的最后一个子元素，移除
      // 简单策略：连续的 br 只保留一个
      if (br.nextElementSibling && br.nextElementSibling.tagName === 'BR') {
          br.remove();
      }
  });
}
