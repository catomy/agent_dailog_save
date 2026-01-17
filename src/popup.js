document.getElementById('exportBtn').addEventListener('click', async () => {
  const btn = document.getElementById('exportBtn');
  const statusEl = document.getElementById('status');
  const autoScroll = document.getElementById('autoScroll').checked;

  btn.disabled = true;
  statusEl.innerHTML = '<div class="log-item">正在初始化...</div>';

  try {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

    // 1. 尝试连接（按需注入）
    statusEl.innerHTML += '<div class="log-item">正在连接页面...</div>';
    
    let isConnected = false;
    try {
      // 尝试发送 ping
      const response = await chrome.tabs.sendMessage(tab.id, { action: 'ping' });
      if (response && response.status === 'pong') {
        isConnected = true;
        statusEl.innerHTML += '<div class="log-item" style="color:blue">复用现有转换引擎...</div>';
      }
    } catch (e) {
      // 连接失败，说明未注入
    }

    if (!isConnected) {
      statusEl.innerHTML += '<div class="log-item">正在注入转换引擎...</div>';
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        files: ['content.js']
      });
      // 给一点时间让脚本初始化
      await new Promise(r => setTimeout(r, 200));
    }

    // 2. 发送开始命令
    try {
      // await 确保消息成功发送到 content script
      // 由于 content script 中没有 return true，sendMessage 会在 handler 执行完同步代码后立即返回
      // 不会等待 runExport 异步任务完成，符合我们"后台运行"的需求
      await chrome.tabs.sendMessage(tab.id, {
        action: 'start_export',
        config: { autoScroll }
      });
      
      // 立即反馈给用户
      statusEl.innerHTML += '<div class="log-item" style="color:green">✅ 任务已启动！</div>';
      statusEl.innerHTML += '<div class="log-item">您现在可以关闭此窗口或切换网页。</div>';
      statusEl.innerHTML += '<div class="log-item">导出完成后会弹出通知。</div>';
      
      // 禁用按钮但不需要一直等待
      btn.textContent = "后台运行中...";
      
    } catch (retryErr) {
      // 如果注入后立刻发送失败，再试一次
      await new Promise(r => setTimeout(r, 500));
      await chrome.tabs.sendMessage(tab.id, {
        action: 'start_export',
        config: { autoScroll }
      });
      statusEl.innerHTML += '<div class="log-item" style="color:green">✅ 任务已启动！(重试成功)</div>';
    }

  } catch (err) {
    console.error(err);
    statusEl.innerHTML += `<div class="log-item" style="color:red">错误: ${err.message}</div>`;
    btn.disabled = false;
  }
});

// 移除对 content script 日志的实时监听，因为 popup 可能关闭
// chrome.runtime.onMessage.addListener... (不再需要)
