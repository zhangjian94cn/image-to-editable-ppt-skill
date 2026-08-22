# 示例提示词

## 单页

```text
使用 $image-to-editable-ppt，将 source.png 重建为对象级可编辑的单页 PPTX。
源图文字、数据、分组和相对几何是权威。普通内容使用原生对象，只允许
Logo、照片、截图和复杂插图作为紧凑局部图片。真实 PowerPoint 渲染并
查看 page.pptx 后再写 result.json。
```

## 多页

```text
使用 $image-to-editable-ppt 处理这份 PDF。每页分别运行一个本机 Codex
任务，完成后按原页序用 editppt assemble 合并，只交付可编辑 PPTX。
```

失败后重新运行该页面任务，不恢复旧 Controller session，也不使用整页
截图兜底。
