# Image to Editable PPT Skill

**简体中文** · [English](README_en.md) · [한국어](README_ko.md)

把幻灯片图片、扫描 PDF 或图片版 PPT/PPTX 重建为对象级可编辑 PowerPoint。该 Fork 以 `ningzimu/main@fb869763` 为同步基线，只维护一个开放的 Codex + Skill 工作流，不包含页面子 Agent、Controller 状态机或强制质量门禁。

## 核心用法

在 Codex 中给页面图片并明确选择 Skill：

```text
使用 $image-to-editable-ppt，将 source.png 重建为一页对象级可编辑的 PowerPoint。

你可以自主使用 Skill 提供的 OCR、素材提取、PPT 构建和预览工具。
不要把整张 source.png 作为全页背景或全页覆盖图片。
请在当前页面目录生成 page.pptx，并在完成后写出 result.json。
```

一个 Codex 任务负责一页，从观察、构建到按需预览修改。Skill 不会再创建 page worker、子 Agent、session controller、coverage/containment 图或固定 repair 循环。

## 输出契约

```text
source.png
page.pptx
result.json
preview.png   # 可选
```

`result.json`：

```json
{
  "status": "ready",
  "output_pptx": "page.pptx",
  "warnings": []
}
```

禁止把整页源图作为背景或覆盖层伪装成可编辑结果；Logo、照片、地图、截图和复杂插图可以作为独立局部图片对象。

## CLI

```bash
editppt prepare <input...>
editppt inspect <page-dir>
editppt extract-assets --input <image> --out-dir <dir>
editppt build <page-dir>
editppt render <page-dir>
editppt assemble <page-dir...> --out <deck.pptx>
editppt doctor --json
```

这些命令是 Codex 可按需调用的确定性工具，不构成强制流水线。多页任务由调用方为每页分别启动 Codex，完成后按原顺序调用 `editppt assemble`。

## 安装与验证

```bash
pipx install --force --python python3.12 ./skills/image-to-editable-ppt/cli
ln -sfn "$PWD/skills/image-to-editable-ppt" ~/.codex/skills/image-to-editable-ppt
editppt doctor --json
python -m unittest discover -s tests -v
```

PaddleOCR token 是可选增强：

```bash
editppt config --paddle-ocr-token "<token>"
```

缺少 OCR 或图片后端时，Codex仍可使用原生文本、形状、表格和连接线工具完成页面；这些辅助能力不应阻断整个任务。

## 可复用素材

`skills/image-to-editable-ppt/assets/brand-catalog.json` 登记了已审阅素材，包括 China Mobile 横版 Logo 和品牌色。只有源页实际包含对应品牌时才使用。

## 边界

- 本 Skill 用于从视觉页面重建可编辑 PPT，不用于从文章或大纲创建全新演示文稿。
- 文字、数字、表格、普通容器、时间轴与结构关系优先使用原生对象。
- 复杂视觉元素可保留为独立局部素材。
- 预览和诊断是 Codex 自主选择的辅助，不是调用方的下载门禁。
- 最终质量通过 Skill、工具和真实样例持续优化，不通过扩张工作台编排实现。
