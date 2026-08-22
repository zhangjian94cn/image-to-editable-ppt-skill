# 常见问题

## 为什么不能只检查对象数量？

对象存在不代表版式正确。标题换行、字体替换、文字重叠和连接线错位只有
在 Microsoft PowerPoint 的真实画面中才能可靠发现，因此必须运行并查看
`editppt render` 的结果。

## PowerPoint 渲染失败怎么办？

不要写 `ready`，也不要切换到截图版或 LibreOffice 伪装成功。保留
`.editppt/render/` 证据，处理 PowerPoint 弹窗、权限或文件问题后再运行。
Skill 的 renderer 只打开并关闭唯一命名的测试副本，不创建额外 canary，
也不关闭用户已有文稿。

## 可以使用图片对象吗？

可以用于 Logo、照片、地图、产品截图和复杂插图，但裁剪必须紧凑且来自
源像素。标题、正文、数字、表格、时间轴、容器和普通图表不能整体栅格化。

## 多页怎么处理？

每页各启动一个独立 Codex 任务，得到 `page.pptx` 后按源顺序调用
`editppt assemble`。不需要 page worker、Controller 或 manifest 状态机。

## OCR 是必需的吗？

不是。PaddleOCR-VL 可提供更好的文本和坐标证据；缺少 token 时工具会
明确标记为几何提示。OCR 不能代替 Codex 观察整页。
