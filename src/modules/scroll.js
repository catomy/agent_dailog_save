const { log } = require('./utils.js');

/**
 * 自动滚动页面以加载所有内容
 * @returns {Promise} 滚动完成的Promise
 */
async function autoScrollPage() {
  return new Promise((resolve, reject) => {
    try {
      let totalHeight = 0;
      const distance = 100;
      const maxAttempts = 500; // 最大尝试次数，防止无限循环
      let attempts = 0;
      
      const timer = setInterval(() => {
        try {
          attempts++;
          if (attempts > maxAttempts) {
            clearInterval(timer);
            window.scrollTo(0, 0);
            log('警告: 自动滚动达到最大尝试次数，可能未完全加载所有内容');
            resolve();
            return;
          }
          
          const scrollHeight = document.body.scrollHeight;
          window.scrollBy(0, distance);
          totalHeight += distance;

          if (totalHeight >= scrollHeight || (window.innerHeight + window.scrollY) >= scrollHeight) {
            clearInterval(timer);
            window.scrollTo(0, 0); // 回到顶部
            setTimeout(resolve, 500);
          }
        } catch (error) {
          clearInterval(timer);
          window.scrollTo(0, 0);
          reject(new Error(`自动滚动过程中出错: ${error.message}`));
        }
      }, 20); // 速度快一点
    } catch (error) {
      reject(new Error(`自动滚动初始化失败: ${error.message}`));
    }
  });
}

module.exports = {
  autoScrollPage
};
