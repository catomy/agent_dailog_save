# Webpage to Word Exporter (Chrome Extension)

这是一个强大的 Chrome 浏览器扩展，能够将任意网页完美导出为 Microsoft Word 文档 (.docx)。它专为处理复杂的技术文档设计，能够完美保留网页布局、渲染数学公式、转换图标，并支持离线查看。

## ✨ 核心功能 (Features)

*   **一键导出**: 将当前浏览的网页内容转换为 Word 文档，自动处理分页和边距。
*   **数学公式完美支持**:
    *   **自动识别**: 智能捕获网页中已渲染的 MathJax 和 KaTeX 公式。
    *   **原生 LaTeX 支持**: 内置 **KaTeX 引擎**，自动扫描并渲染纯文本格式的 LaTeX 源码（如 `$E=mc^2$` 或 `\[...\]`），将其转换为图片嵌入文档，完美解决 Markdown 预览页或论文网站的公式显示问题。
*   **智能图像处理**:
    *   **离线支持**: 自动将所有图片（包括跨域图片）转换为 Base64 编码，确保导出的 Word 文档在无网络环境下也能正常显示。
    *   **图标光栅化**: 自动识别并转换 SVG 图标和字体图标（如 FontAwesome, Material Icons, Glyphicons），确保在 Word 中不乱码。
    *   **容错机制**: 自动清理无法加载或损坏的图片链接，防止导出中断。
*   **样式保留**: 深度内联 CSS 样式，最大程度保留网页原本的字体、颜色、背景和布局结构。
*   **自动滚动**: 提供“自动滚动”选项，强制触发网页的懒加载（Lazy Load）图片和内容，确保导出完整页面。
*   **交互反馈**: 实时进度日志显示，插件图标状态随任务自动切换（灰色待机 -> 蓝色工作 -> 通知完成）。

## 🛠️ 技术实现 (Implementation)

本项目基于 **Manifest V3** 架构开发，采用现代化的前端工程化方案。

*   **核心架构**:
    *   `content.js`: 核心逻辑引擎。负责 DOM 遍历、样式内联、资源抓取和文档生成。
    *   `html-docx-js`: 用于在浏览器端直接生成 Word 文档 Blob 流。
    *   **KaTeX**: 集成在 Content Script 中，用于动态渲染纯文本 LaTeX。
    *   **Webpack 5**: 用于项目构建和打包。
*   **关键技术点**:
    *   **编码兼容性**: 配置了 `TerserPlugin` 的 `ascii_only: true` 选项，强制将所有 Unicode 字符（如数学符号、中文）转义为 ASCII 序列，彻底解决了 Chrome 扩展在不同操作系统下的 `Could not load file... encoding` 加载错误。
    *   **动态注入**: 采用 Lazy Injection 策略，仅在用户点击时注入 Content Script，减少内存占用。
    *   **Canvas 光栅化**: 使用 HTML5 Canvas 技术将复杂的 DOM 元素（如字体图标）“截图”为 PNG 图片。

## 🚀 使用方法 (Usage)

### 开发与构建
1.  **克隆仓库**:
    ```bash
    git clone https://github.com/catomy/agent_dailog_save.git
    cd webpage-to-word
    ```
2.  **安装依赖**:
    ```bash
    npm install
    # 建议使用 cnpm 或 yarn 以获得更快的速度
    ```
3.  **构建项目**:
    ```bash
    npm run build
    ```
    构建完成后，会在 `dist` 目录下生成插件代码。

### 安装到 Chrome
1.  打开 Chrome 浏览器，访问 `chrome://extensions/`。
2.  开启右上角的 **"开发者模式" (Developer mode)**。
3.  点击左上角的 **"加载已解压的扩展程序" (Load unpacked)**。
4.  选择本项目下的 `dist` 文件夹。

### 开始使用
1.  打开任意需要导出的网页（例如一篇包含公式的技术博客）。
2.  点击浏览器右上角的插件图标（灰色的 "W" 图标）。
3.  （可选）勾选 "自动滚动页面" 以加载更多内容。
4.  点击 **"开始转换"** 按钮。
5.  图标会变蓝，表示正在后台处理。处理完成后，浏览器会自动下载生成的 `.docx` 文件，并弹出系统通知。

## 📝 目录结构

```
webpage-to-word/
├── src/
│   ├── background.js      # 后台服务，处理图标状态和通知
│   ├── content.js         # 核心业务逻辑，DOM 操作与导出
│   ├── popup.js           # 弹窗交互逻辑
│   ├── popup.html         # 弹窗 UI
│   ├── manifest.json      # 插件配置清单
│   └── icons/             # 图标资源
├── dist/                  # 构建产物 (直接加载此目录)
├── webpack.config.js      # Webpack 配置
└── package.json           # 项目依赖
```

## License
MIT
