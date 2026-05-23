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
      const startResp = await chrome.tabs.sendMessage(tab.id, {
        action: 'start_export',
        config: { autoScroll }
      });

      if (startResp && startResp.status === 'busy') {
        statusEl.innerHTML += '<div class="log-item" style="color:orange">⚠️ 当前页面已有导出任务正在进行，请稍后再试。</div>';
        btn.disabled = false;
        return;
      }
      
      // 立即反馈给用户
      statusEl.innerHTML += '<div class="log-item" style="color:green">✅ 任务已启动！</div>';
      statusEl.innerHTML += '<div class="log-item">您现在可以关闭此窗口或切换网页。</div>';
      statusEl.innerHTML += '<div class="log-item">导出完成后会弹出通知。</div>';
      
      // 禁用按钮但不需要一直等待
      btn.textContent = "后台运行中...";
      
    } catch (retryErr) {
      // 如果注入后立刻发送失败，再试一次
      await new Promise(r => setTimeout(r, 500));
      const retryResp = await chrome.tabs.sendMessage(tab.id, {
        action: 'start_export',
        config: { autoScroll }
      });
      if (retryResp && retryResp.status === 'busy') {
        statusEl.innerHTML += '<div class="log-item" style="color:orange">⚠️ 当前页面已有导出任务正在进行，请稍后再试。</div>';
        btn.disabled = false;
        return;
      }
      statusEl.innerHTML += '<div class="log-item" style="color:green">✅ 任务已启动！(重试成功)</div>';
    }

  } catch (err) {
    console.error(err);
    statusEl.innerHTML += `<div class="log-item" style="color:red">错误: ${err.message}</div>`;
    btn.disabled = false;
  }
});

// 监听来自 content script 的日志消息
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'log') {
    const statusEl = document.getElementById('status');
    statusEl.innerHTML += `<div class="log-item">${request.message}</div>`;
    statusEl.scrollTop = statusEl.scrollHeight; // 自动滚动到底部
  }
});
