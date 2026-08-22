# 安装与环境

```bash
pipx install --force --python python3.12 ./skills/image-to-editable-ppt/cli
ln -sfn "$PWD/skills/image-to-editable-ppt" ~/.codex/skills/image-to-editable-ppt
editppt doctor --json
```

运行需要本机 Codex CLI、Python 3.10+ 和 Microsoft PowerPoint for macOS。
`editppt doctor` 会检查 OCR 配置、字体、PowerPoint、Skill 版本、PPTX
依赖和合并能力。

PaddleOCR-VL 是可选增强，可通过 `PADDLE_OCR_TOKEN` 或
`~/.editppt/config.yaml` 提供。没有 OCR token 时仍可重建，但输出会明确
标记 OCR 降级。没有 PowerPoint 或真实渲染失败时不得写 `ready`。

页面 Codex 任务需要能够在隔离页面目录写文件，并允许操作本机
PowerPoint。benchmark runner 只用于冻结的可信语料，会使用 host 访问权限；
妙笔产品运行时不导入 benchmark 代码。
