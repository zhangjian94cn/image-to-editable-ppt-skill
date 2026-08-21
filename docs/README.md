# Image to Editable PPT Skill

把页面图片、扫描 PDF 或视觉型 PPTX 重建为对象级可编辑 PowerPoint。

该 Fork 只维护一个简单架构：一页图片对应一个本机 Codex 任务，Codex
使用 `$image-to-editable-ppt` 及 Skill 内的确定性工具生成 `page.pptx`，
经 Microsoft PowerPoint 真实渲染并查看后写出 `result.json`。多页由调用方
按原顺序调用 `editppt assemble`。

## 文档

- [快速开始](quickstart.md)
- [设计与边界](design.md)
- [安装与环境](installation.md)
- [标准工作流](workflow.md)
- [常见问题](faq.md)
- [示例提示词](prompts.md)

## 关键能力

- OCR、布局和结构证据，而不是强制 page plan。
- 原像素局部素材提取和已审核品牌素材。
- 实际字体解析与适配、富文本、原生表格、原生连接线和箭头。
- `editppt.authoring.SlideManifest` 复用组件。
- 只渲染目标文件的 Microsoft PowerPoint 权威预览。
- 独立 `page.pptx` 的关系感知多页合并。

禁止整页或接近整页的源图覆盖。没有页面子 Agent、Controller、session
状态、coverage/containment 或 Hybrid 自动兜底。
