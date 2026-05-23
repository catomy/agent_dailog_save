const htmlDocx = require('html-docx-js/dist/html-docx');
const { saveAs } = require('file-saver');

// 导入模块
const { log, cleanupMemory } = require('./modules/utils.js');
const { autoScrollPage } = require('./modules/scroll.js');
const { processRawLatex } = require('./modules/latex.js');
const { captureVisualElements, applyMathImages } = require('./modules/visual-elements.js');
const { processImages, applyDocImages, cleanUnprocessedImages } = require('./modules/images.js');
const { inlineStyles } = require('./modules/styles.js');
const { fixLinks, cleanClone } = require('./modules/cleanup.js');

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
if (!window[GLOBAL_FLAG]) {
  setupListener();
  window[GLOBAL_FLAG] = true;
  chrome.runtime.sendMessage({ action: 'engine_ready' });
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
        .catch(error => {
          console.error(error);
          chrome.runtime.sendMessage({ action: 'export_error', message: error.toString() });
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
  let errorCount = 0;
  const errors = [];
  
  try {
    // 1. 标记 DOM 节点，以便后续映射
    log('正在索引页面元素...');
    let nodeIdCounter = 0;
    const allElements = document.querySelectorAll('*');
    const totalElements = allElements.length;
    log(`发现 ${totalElements} 个页面元素`);
    allElements.forEach((el, index) => {
      el.setAttribute('data-docx-id', nodeIdCounter++);
      // 每处理 1000 个元素更新一次进度
      if (index % 1000 === 0 && index > 0) {
        log(`已索引 ${index}/${totalElements} 个元素`);
      }
    });
    log(`完成元素索引，共 ${nodeIdCounter} 个元素`);

    // 2. 自动滚动
    if (config.autoScroll) {
      log('正在滚动加载内容...');
      log('这可能需要一些时间，取决于页面长度...');
      try {
        await autoScrollPage();
        log('滚动加载完成');
      } catch (error) {
        errorCount++;
        errors.push(`自动滚动失败: ${error.message}`);
        log(`警告: 自动滚动失败，将继续执行导出 - ${error.message}`);
      }
    }

    // 3. 处理 MathJax 公式和 SVG 图标
    log('正在可视化处理公式和图标...');
    let mathImages = [];
    try {
      mathImages = await captureVisualElements();
      log(`处理了 ${mathImages.length} 个可视化元素`);
    } catch (error) {
      errorCount++;
      errors.push(`公式和图标处理失败: ${error.message}`);
      log(`警告: 公式和图标处理失败，将继续执行导出 - ${error.message}`);
    }

    // 3.5. 处理纯文本 LaTeX 公式
    log('正在扫描并渲染纯文本 LaTeX...');
    let latexCount = 0;
    try {
      latexCount = processRawLatex(document.body);
      log(`渲染了 ${latexCount} 个纯文本 LaTeX 公式`);
    } catch (error) {
      errorCount++;
      errors.push(`LaTeX 公式处理失败: ${error.message}`);
      log(`警告: LaTeX 公式处理失败，将继续执行导出 - ${error.message}`);
    }

    // 4. 处理常规图片和图标 (转 Base64 以支持离线/Word)
    log('正在处理网页图片和图标...');
    let docImages = [];
    try {
      docImages = await processImages();
      log(`处理了 ${docImages.length} 个图片和图标`);
    } catch (error) {
      errorCount++;
      errors.push(`图片处理失败: ${error.message}`);
      log(`警告: 图片处理失败，将继续执行导出 - ${error.message}`);
    }

    // 5. 克隆页面
    log('正在创建文档副本...');
    let clonedBody;
    try {
      clonedBody = document.body.cloneNode(true);
      log('文档副本创建完成');
    } catch (error) {
      errorCount++;
      errors.push(`文档克隆失败: ${error.message}`);
      log(`警告: 文档克隆失败，将使用原始文档 - ${error.message}`);
      clonedBody = document.body;
    }

    // 6. 内联样式 (最耗时步骤)
    log('正在内联 CSS 样式 (保留颜色、字体)...');
    log('这是最耗时的步骤，请耐心等待...');
    try {
      await inlineStyles(document.body, clonedBody);
      log('样式内联完成');
    } catch (error) {
      errorCount++;
      errors.push(`样式内联失败: ${error.message}`);
      log(`警告: 样式内联失败，将继续执行导出 - ${error.message}`);
    }

    // 7. 替换公式为图片
    log('正在应用公式图片...');
    try {
      applyMathImages(clonedBody, mathImages);
      log('公式图片应用完成');
    } catch (error) {
      errorCount++;
      errors.push(`公式图片应用失败: ${error.message}`);
      log(`警告: 公式图片应用失败，将继续执行导出 - ${error.message}`);
    }

    // 8. 应用常规图片替换 (替换为 Base64)
    log('正在应用图片 Base64 替换...');
    try {
      applyDocImages(clonedBody, docImages);
      log('图片 Base64 替换完成');
    } catch (error) {
      errorCount++;
      errors.push(`图片 Base64 替换失败: ${error.message}`);
      log(`警告: 图片 Base64 替换失败，将继续执行导出 - ${error.message}`);
    }

    // 9. 清理未成功转换的图片 (防止 html-docx-js 报错)
    log('正在清理未处理的图片...');
    try {
      cleanUnprocessedImages(clonedBody);
      log('未处理图片清理完成');
    } catch (error) {
      errorCount++;
      errors.push(`图片清理失败: ${error.message}`);
      log(`警告: 图片清理失败，将继续执行导出 - ${error.message}`);
    }

    // 10. 修复链接 (转为绝对路径)
    log('正在修复链接...');
    try {
      fixLinks(clonedBody);
      log('链接修复完成');
    } catch (error) {
      errorCount++;
      errors.push(`链接修复失败: ${error.message}`);
      log(`警告: 链接修复失败，将继续执行导出 - ${error.message}`);
    }

    // 11. 清理垃圾元素
    log('正在清理垃圾元素...');
    try {
      cleanClone(clonedBody);
      log('垃圾元素清理完成');
    } catch (error) {
      errorCount++;
      errors.push(`垃圾元素清理失败: ${error.message}`);
      log(`警告: 垃圾元素清理失败，将继续执行导出 - ${error.message}`);
    }

    // 12. 生成 Word
    log('正在生成 Word 文档...');
    log('正在构建最终 HTML 内容...');
    let contentHtml;
    try {
      contentHtml = `
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
    } catch (error) {
      errorCount++;
      errors.push(`HTML 构建失败: ${error.message}`);
      log(`警告: HTML 构建失败，将使用简化版本 - ${error.message}`);
      contentHtml = `
        <!DOCTYPE html>
        <html>
          <head>
            <meta charset="UTF-8">
          </head>
          <body>
            <p>导出过程中遇到错误，文档可能不完整</p>
          </body>
        </html>
      `;
    }

    log('正在转换为 Word 格式...');
    let converted;
    try {
      // html-docx-js 需要 Buffer 或 ArrayBuffer，但在浏览器端它接受字符串并返回 Blob
      converted = htmlDocx.asBlob(contentHtml, {
        orientation: 'portrait',
        margins: { top: 720, bottom: 720, left: 720, right: 720 } // twips
      });
    } catch (error) {
      errorCount++;
      errors.push(`Word 格式转换失败: ${error.message}`);
      log(`错误: Word 格式转换失败 - ${error.message}`);
      throw new Error(`Word 格式转换失败: ${error.message}`);
    }

    log('正在保存文件...');
    const fileName = `Page_Export_${Date.now()}.docx`;
    try {
      saveAs(converted, fileName);
    } catch (error) {
      errorCount++;
      errors.push(`文件保存失败: ${error.message}`);
      log(`错误: 文件保存失败 - ${error.message}`);
      throw new Error(`文件保存失败: ${error.message}`);
    }
    
    log('完成！');
    log(`Word 文档 "${fileName}" 已生成并开始下载`);
    
    // 发送完成通知，包含错误信息
    chrome.runtime.sendMessage({ 
      action: 'export_done',
      errorCount,
      errors
    });

    // 清理添加的 data-docx-id 属性
    log('正在清理临时属性...');
    try {
      const elementsWithDocxId = document.querySelectorAll('*[data-docx-id]');
      const cleanupCount = elementsWithDocxId.length;
      elementsWithDocxId.forEach((el, index) => {
        el.removeAttribute('data-docx-id');
        // 每清理 1000 个元素更新一次进度
        if (index % 1000 === 0 && index > 0) {
          log(`已清理 ${index}/${cleanupCount} 个临时属性`);
        }
      });
      log('临时属性清理完成');
    } catch (error) {
      log(`警告: 临时属性清理失败 - ${error.message}`);
    }

    // 强制垃圾回收（如果可用）
    cleanupMemory();

    // 如果有错误，在完成通知中显示
    if (errorCount > 0) {
      log(`导出完成，但遇到 ${errorCount} 个警告`);
      errors.forEach((error, index) => {
        log(`警告 ${index + 1}: ${error}`);
      });
    }
  } catch (error) {
    console.error('导出过程中发生严重错误:', error);
    chrome.runtime.sendMessage({ 
      action: 'export_error', 
      message: error.toString(),
      errorCount,
      errors: [...errors, `严重错误: ${error.message}`]
    });
    throw error;
  }
}
