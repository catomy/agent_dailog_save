// 监听来自 Content Script 的消息
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'export_done') {
    // 弹出成功通知
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon_color.svg',
      title: '导出成功',
      message: 'Word 文档已生成并开始下载。',
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
    chrome.notifications.create({
      type: 'basic',
      iconUrl: 'icons/icon_gray.svg',
      title: '导出失败',
      message: request.message || '未知错误',
      priority: 2
    });
  } else if (request.action === 'engine_ready' || request.action === 'export_start') {
      // engine_ready: content script 注入并完成初始化
      // export_start: 导出任务开始
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
  }
});
