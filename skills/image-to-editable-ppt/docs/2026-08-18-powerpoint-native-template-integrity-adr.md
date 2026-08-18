# ADR: PowerPoint 原生模板与 Open XML 零错误构建门禁

- 日期: 2026-08-18
- 状态: accepted
- 范围: image-to-editable-ppt / manifest-to-PPTX runtime
- 相关源码: [builder](../cli/editppt/runtime/build_pptx_from_manifest.py)、[integrity gate](../cli/editppt/runtime/pptx_integrity.py)、[validator project](../tools/openxml-validator/openxml-validator.csproj)、[validator source](../tools/openxml-validator/Program.cs)

## Context and Problem Statement

旧生成器手写最小 OOXML 包，主题四个样式列表基数不足，并把 16:9 写成非法
`type="wide"`。PowerPoint 因此进入修复流程；第一个坏文件产生的隐藏模态还会污染
同一进程，使后续正常文件看起来也“没有打开”。

## Decision Drivers

- 自己生成的 PPTX 必须在启动 PowerPoint 前得到结构零错误证据。
- 母版、版式、主题和 presentation properties 不应继续由业务代码猜测。
- 文字、表格和形状必须保持原生可编辑，不能让 PowerPoint 修复后另存来擦除错误。
- 部署期可以构建验证工具，任务运行期不得联网恢复 NuGet 依赖。

## Considered Options

- 继续补丁式完善手写最小包：短期小，但每个 Office 扩展与关系顺序都可能产生新缺口。
- 让 PowerPoint 修复后另存：表面可打开，但可能静默删除对象或改变文本属性。
- 复用 PowerPoint 创建的基础包，只注入 slides/media/notes，再用固定 Open XML SDK 预检。

## Decision Outcome

采用第三种方案。默认基础包来自 python-pptx 随附、属性标识为 Microsoft Macintosh
PowerPoint 的零页模板；部署可用 `EDITPPT_POWERPOINT_TEMPLATE` 指向当前 PowerPoint
创建并验收的固定模板。构建器保留其 master、blank layout、theme、view/presentation
properties 和关系结构，只改合法 slide size、slide 列表和任务对象。生成后使用锁定为
DocumentFormat.OpenXml 3.3.0 的本地工具验证；`error_count != 0` 时直接失败并写
`pptx-integrity-report.json`，不启动 PowerPoint。

对于用户提供或 truth set 使用的 authored PPTX，SDK 结果不再整体降级为 warning。
只有 `/ppt/charts/` 中 SDK 尚不认识的 Microsoft Office / DrawingML chart extension
“unexpected child element”可以进入 PowerPoint 真机验证；重复 ID、关系错误、核心
PresentationML schema 错误及其他未明确允许的错误全部在启动 PowerPoint 前进入
`manual_pending`。分类与证据由 [integrity gate](../cli/editppt/runtime/pptx_integrity.py)
单点维护，调用方不得复制一套宽松判断。

唯一的确定性净化例外是：所有阻断错误都来自 `commentAuthors.xml` 的重复 ID，且包中
不存在任何 comment part 或 comment relationship。此时门禁可以生成一次性的派生副本，
删除孤立的作者 part、对应 relationship 和 Content-Type override；原 authored PPTX
保持逐字节不变。派生副本必须重新通过 SDK 分类，并继续完成 PowerPoint 无修复验收。
若存在真实评论、错误数量与重复数量不一致或仍有其他阻断错误，净化不适用并维持
`manual_pending`。报告必须同时记录原始/派生哈希、操作清单和
`visual_content_changed=false`。

manifest v2 的每个文本框、表格和单元格都必须绑定绝对字体文件、内部名称和 SHA-256；
表格写成原生 `a:tbl`，不允许回退为截图或一组无结构线条。

## Consequences

- 生成包可追溯到模板哈希、生成器版本、build id 和 SDK 版本。
- 模板或验证器变化会改变环境证据，已有结果需要重验。
- SDK 零错误只证明结构可验证；仍必须经过真实 PowerPoint 无修复导出与像素门禁。
- authored 文件的扩展 allowlist 必须保持窄且可测试；未知错误允许假阴性，不允许冒险打开并触发修复模态。
- 孤立评论作者净化只处理无法影响幻灯片画面的遗留元数据，不扩展为通用 OOXML 自动修复器。
- 当前 PowerPoint 会话若被隐藏模态污染，结构重建可以完成，但交付保持 `manual_pending`。
