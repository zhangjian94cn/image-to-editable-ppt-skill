# 标准工作流

本 Skill 只维护“一页图片、一个 Codex、一个 Skill”的页面流程。

1. 调用方用 `editppt prepare` 将图片、PDF 或视觉型 PPTX 拆成有序的
   `pages/page_NNN/source.png`。
2. 每一页启动一个独立的本机 Codex 任务，并显式使用
   `$image-to-editable-ppt`。
3. Codex 先观察整页并列出主要区域，再按需调用 OCR、布局、结构、素材、
   Builder 和字体工具。
4. 文字、数字、表格、普通图表、时间轴、容器和连接关系使用原生对象；
   Logo、照片、地图、截图和复杂插图可保留为紧凑的局部图片。
5. Codex 生成 `page.pptx`，用 `editppt render` 让 Microsoft PowerPoint
   打开并真实渲染该文件，然后查看 `preview.png` 并针对可见问题修改。
6. 只有真实渲染成功后才写 `result.json`。多页任务最后调用
   `editppt assemble` 按源顺序合并独立的 `page.pptx`。

页面目录的最小结果是：

```text
source.png
page.pptx
preview.png
result.json
```

没有 page worker、Controller、session 恢复、dispatch/record/finalize、
coverage/containment 或 Hybrid 兜底。Microsoft PowerPoint 不可用、文件被
修复或渲染失败时，不得写 `ready`，也不得改用整页截图交付。

确定性工具是 Codex 可按需选择的能力，不是外部强制流水线。
