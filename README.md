# DSH Office Canvas

为 DeepSeek Harness 的右侧文档面板提供 DOCX / PPTX 编辑、页面预览与浮动批注。当前版本 **0.2.1**，为试用版。

## 功能

- **编辑**：修改正文文字、加粗、斜体、下划线、字号和颜色；PPT 支持普通文本及可解析坐标对象的位置调整。
- **预览**：通过已安装的 LibreOffice Kit 显示保存后的原版式，页面自适应面板宽度。
- **提出修改请求**：选中文字、对象或框选页面区域，在附近的浮动输入框填写更改要求。
- **多条批注**：连续添加批注，从底部操作条查看、删除、取消、加入对话或发送。
- **版本检查与副本保存**：请求携带路径、SHA-256 和选择目标；发送前核对文件版本，保存默认创建新副本。

## 安装

本插件需要支持本地组合包的 DSH 桌面环境、Python（含 lxml）及 LibreOffice Kit。未捆绑 DSH、Python 或 LibreOffice 的运行时。

1. 下载本仓库源码或 `dist/local-dsh-office-canvas-0.2.1.tgz`。
2. 在 DSH 插件管理器安装本地目录，或在支持本地包的入口安装 tgz。
3. 在插件 Host Config 配置本机 `pythonPath`、`nodePath`、`libreofficeModulePath`。源码中的默认值来自开发环境，其他机器应覆盖这些配置。
4. 重载插件或重启 DSH。在当前会话打开工作区内 DOCX/PPTX，右侧预览器选择「Office 画布」。

包名：`@local/dsh-office-canvas`；插件行为：`local-office-canvas`。

## 使用

在「编辑」视图中修改内容并保存副本。选中文字或在「预览」中框选区域，点击「提出修改请求」，填写要求后按上箭头加入批注队列。底部「加入对话」只插入草稿，「发送」才提交到当前会话。

有未保存编辑时，先保存再提交修改请求。保存或重新读取文件后，旧版本批注会清空。关闭面板前请保存；当前没有跨面板关闭的草稿恢复。

## 验证与限制

- 0.2.0 客户端编辑逻辑回归 11/11 通过；新浮动批注真实 Chrome / React DOM 检查 5/5 通过，后端响应使用测试夹具。
- 后端已有 15 项回归与真实 DOCX/PPTX 保存重开通过记录；Host 集成检查通过。
- 0.2.1 仅把入口名称调整为「编辑 / 预览」，语法检查通过。
- 完整 DSH 安装后的端到端现场验收尚未完成。多行输入和完整 PPT 编辑仍需进一步验证。
- DOCX 编辑视图按段落流显示，不复现全部分页、浮动图文、页眉页脚和表格网格；准确排版请使用预览。
- PPT 组合对象、SmartArt、复杂图表等保留在原文件中，不具备完整可视编辑能力。

建议先用副本试用，保存后重新打开确认结果。

## 开发测试

Node.js 需支持 ESM、fetch、Request/Response。Python 测试需要 `lxml`、`python-docx` 和 `python-pptx`。

```sh
node test_client.mjs
node test_host.mjs
python -m unittest -v test_backend
```

安装包可用 `pnpm pack` 生成。文件操作通过当前会话的 shell 与 sandboxPolicy 执行，限定在会话工作区；保存校验源文件 hash，未修改的 OOXML package parts 保留。

## 说明

这是独立开发的 DSH 本地插件，不是 OpenAI 官方产品。交互参考公开文件审阅体验；未使用 OpenAI 私有实现。图标为本插件原创 SVG。当前仓库未声明开源许可证。
