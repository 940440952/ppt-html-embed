# HTML 嵌入 · PowerPoint 插件

把本地 HTML 嵌入 PowerPoint 幻灯片：HTML 会保存进 .pptx 文件，编辑和放映时都能交互。

## 安装（每台电脑做一次）
1. 联网，在终端里运行 `bash install-mac.sh`（`manifest.xml` 要和它放在同一个文件夹）。
2. 完全退出 PowerPoint（Cmd+Q）后重新打开：**插入 → 加载项 → 我的加载项 → HTML 嵌入**。
3. 联网插入一次，之后插件会缓存在本机，离线也能用。

## 使用
- **选择 HTML 文件**：单个 HTML 文件。
- **选择文件夹**：HTML 引用了同目录的图片、CSS、JS 时用这个，插件会自动打包成一个文件。
- **粘贴代码**：直接粘贴 HTML 代码，比如 AI 生成的页面。
- 鼠标移到框上会出现工具栏：刷新 / 更换。放映时工具栏会自动隐藏。

## 限制
- 内容最好不要超过 2 MB。
- HTML 里用 `fetch()` 读取同目录 JSON 等文件的代码不会被打包。
- 别人打开这个 PPT 时，也需要装好这个插件。

## 部署到自己的 GitHub Pages
`./build-manifest.sh https://<用户名>.github.io/<仓库名>`，生成 `manifest.xml` 后提交并推送。
