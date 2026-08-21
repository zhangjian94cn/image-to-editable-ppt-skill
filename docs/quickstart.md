# 快速开始

先安装 Skill 和 CLI，并运行 `editppt doctor --json`。然后给 Codex 一张
页面图片并输入：

```text
使用 $image-to-editable-ppt，将 source.png 重建为一页对象级可编辑的 PowerPoint。
不要把整张 source.png 作为背景或覆盖图片。
生成 page.pptx，用 Microsoft PowerPoint 真实渲染并查看后再写 result.json。
```

正常页面目录包含 `source.png`、`page.pptx`、`preview.png` 和
`result.json`。多页输入先用 `editppt prepare` 拆页，每页分别启动一个
Codex，最后按原顺序运行 `editppt assemble`。

PaddleOCR-VL token 是可选增强。第一次建议先跑一页，人工对照源图、
PowerPoint 渲染和可编辑对象，再决定是否扩大到整份文稿。
