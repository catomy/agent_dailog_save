// 监听来自 Content Script 的消息
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'export_done') {
    // 弹出成功通知
    let message = 'Word 文档已生成并开始下载。';
    if (request.errorCount && request.errorCount > 0) {
      message += `\n注意：导出过程中遇到 ${request.errorCount} 个警告，部分内容可能未完全处理。`;
    }
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon_color.svg',
      title: '导出成功',
      message: message,
      priority: 2
    });
    // 重置图标（可选）
    if (sender.tab) {
        chrome.action.setIcon({ 
            tabId: sender.tab.id, 
            path: {
                "16": "icons/icon_gray.svg",
                "48": "icons/icon_gray.svg",
                "128": "icons/icon_gray.svg"
            }
        });
    }

  } else if (request.action === 'export_error') {
    // 弹出失败通知
    let message = request.message || '未知错误';
    if (request.errorCount && request.errorCount > 0) {
      message += `\n共遇到 ${request.errorCount} 个错误。`;
    }
    if (request.errors && request.errors.length > 0) {
      // 只显示前3个错误，避免通知过长
      const shownErrors = request.errors.slice(0, 3);
      shownErrors.forEach((error, index) => {
        message += `\n${index + 1}. ${error}`;
      });
      if (request.errors.length > 3) {
        message += `\n...等 ${request.errors.length - 3} 个错误`;
      }
    }
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon_gray.svg',
      title: '导出失败',
      message: message,
      priority: 2
    });
  } else if (request.action === 'engine_ready' || request.action === 'export_start') {
      // engine_ready 表示转换引擎注入完成，export_start 表示导出任务开始。
      if (sender.tab) {
          chrome.action.setIcon({ 
              tabId: sender.tab.id, 
              path: {
                  "16": "icons/icon_color.svg",
                  "48": "icons/icon_color.svg",
                  "128": "icons/icon_color.svg"
              }
          });
      }
  } else if (request.action === 'log') {
      // 转发日志消息给所有打开的 popup 窗口
      chrome.runtime.sendMessage(request);
  }
});
