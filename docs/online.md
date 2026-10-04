# GitHub Pages 在线部署

本工具是静态网站。服务器只提供页面、脚本和样式；工作簿解析、筛选、排序和导出均在访问者的浏览器内完成，不需要上传文件接口、数据库或 API 密钥。

## 构建与托管

在独立子项目目录执行：

```bash
npm ci
npm run build
```

仓库的 `.github/workflows/pages.yml` 在 `main` 分支推送或手动触发时，使用 Node.js 24 执行 `npm ci`、`npm test` 和 `npm run build`，然后将 `dist/` 作为 GitHub Pages 资源发布。网站入口为 `index.html`，资源使用相对路径，适用于 `https://jialep38-png.github.io/excel-workbench/` 这样的项目子目录。构建时须保留 `assets/` 和 `vendor/` 的目录结构。

在仓库的 **Settings → Pages → Build and deployment** 中将 **Source** 设为 **GitHub Actions**。工作流部署成功后，通过 [公开页面](https://jialep38-png.github.io/excel-workbench/)访问；不需要登录或配置应用账号。首次启用或修改 Pages 设置后，以 GitHub Actions 的部署结果和实际页面为准。

如需单文件离线版，在本地执行上述构建命令后使用 `release/Excel筛选工作台.html`。它包含运行脚本、样式和许可，可在断网环境直接打开。

本地检查常规构建可用：

```bash
npm run preview
```

## 访问与数据

公开页面仅加载应用文件。工作簿解析、筛选、排序和导出都在访问者的浏览器内进行，不需要把表格上传到 GitHub Pages。

文件不会上传，保存方案也不会保存完整工作簿。用户主动保存方案时，当前浏览器、当前网站来源的 `localStorage` 会保留表头、筛选值、搜索词和视图设置；筛选值可能包含原表中的分类文字，用完可在“保存或恢复方案”中删除。在线站点与本地离线页面的方案不自动同步。关闭页面后，重新导入原文件，再选择保存的方案即可继续。

访问在线页面需要网络。需要完全断网使用时，下载并保留单文件离线版。

## 更新与发布内容

修改源码并推送到 `main` 后，工作流会重新运行测试、构建和部署；也可以在 GitHub Actions 中手动运行。部署资源仅来自 `dist/`。

公开仓库和部署资源不得包含父项目环境文件、业务名单、开发依赖、验收下载文件或访问凭据。

应用源码采用 MIT 许可。SheetJS CE 和其他依赖的许可见项目 `THIRD_PARTY_NOTICES.md`；单文件构建内也保留许可全文。
