const CTX_ID = chrome.runtime.id;
const GLOBAL_FLAG = `hasRunContentScript_${CTX_ID}`;

/**
 * 日志记录函数
 * @param {string} message - 日志消息
 */
function log(message) {
  chrome.runtime.sendMessage({ action: 'log', message });
  console.log('[WordExport]', message);
}

/**
 * 生成唯一标识符
 * @returns {string} 唯一标识符
 */
function generateId() {
  return Math.random().toString(36).substr(2, 9);
}

/**
 * 检查元素是否可见
 * @param {HTMLElement} element - 要检查的元素
 * @returns {boolean} 是否可见
 */
function isElementVisible(element) {
  try {
    return element.offsetWidth > 5 && element.offsetHeight > 5;
  } catch (error) {
    return false;
  }
}

/**
 * 等待指定时间
 * @param {number} ms - 等待的毫秒数
 * @returns {Promise} 等待完成的Promise
 */
function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

/**
 * 内存清理函数
 */
function cleanupMemory() {
  if (typeof window.gc === 'function') {
    try {
      window.gc();
      log('垃圾回收完成');
    } catch (error) {
      log(`警告: 垃圾回收失败 - ${error.message}`);
    }
  }
}

module.exports = {
  log,
  generateId,
  isElementVisible,
  delay,
  cleanupMemory
};
