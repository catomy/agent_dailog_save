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
  } else if (request.action === 'export_start') {
      // 当导出开始时，点亮图标（如果是灰色默认）
      // 或者我们可以反过来，默认灰色，只有在特定页面或点击后变色
      // 这里根据用户需求：加载后变亮。
      // 实际上，用户说“加载以后需要变成亮色”，通常指页面加载完成（content script注入成功）
      // 我们可以在 content script 初始化时发送一个 'ready' 消息
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
