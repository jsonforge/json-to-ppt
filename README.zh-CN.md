# JSON2PPT （中文）

[![Status – Experimental](https://img.shields.io/badge/status-experimental-orange)](./LIMITATIONS.md)
[![幂等输出](https://img.shields.io/badge/输出-幂等%2F可复现-blue)](#决定性--幂等性-)
[![LLM 友好](https://img.shields.io/badge/LLM-friendly-7B3FE4)](docs/PROMPT_FOR_LLM.md)
[![单文件内核](https://img.shields.io/badge/核心-单文件-lightgrey)](./main.py)

[English README](./README.md) | 中文当前页面

面向 AI 时代的「声明式」PPT 生成器：输入结构化 JSON ➜ 输出可编辑 `.pptx`。让模型 / 自动化脚本直接“操纵”排版，不再手工对齐、调字体、搬数据。

> 愿景：把“写 PPT”转化为“生产结构化叙事数据”，渲染只是副作用。✨

## 为什么存在 🚀
传统方案：要么模板 mail merge 要么截图粗糙。本项目聚焦三点：
1. 面向大模型：Schema 清晰、约束严格、可验证，LLM 快速掌握
2. 完全可控：位置 / 尺寸 / 样式 / 图表 / 表格 全显式，可复现可审计
3. 极简内核：单文件核心（`main.py`），阅读 & 二开成本极低

## 决定性 & 幂等性 🔁
同一份通过校验的 JSON，多次生成结果保持一致（除 Office 自写元数据）。
关键机制：
- 顺序遍历，不引入随机
- 单位换算确定性（`px`→EMU；百分比基于幻灯片尺寸）
- 样式合并规则固定（`merge_styles`）
- 避免隐式推断/自动布局
- 图片推荐 `base64:` / `file:` 源（`url:` 可能变动）

用途价值：
- CI 回归：对生成 `.pptx` 做哈希或二进制比对
- 多 Agent 协作：稳定的元素顺序与索引
- 审计/合规：月度/季度报表可重生成不漂移

## 核心能力一览 🧰
- 文本：段落 / 项目符号 / 编号
- 图片：本地 / Base64 / 远程 URL（含缓存）
- 形状：矩形 / 圆角矩形 / 椭圆
- 表格：头/正文样式分离、斑马纹、列宽控制
- 图表：柱 / 折 / 饼（易扩展）
- 背景：纯色或全幅图片
- 坐标：`px` + 百分比混用（统一换算 EMU）
- 校验：`ppt.schema.json`（Draft 7）

## 快速开始 ⚡
```powershell
uv venv
uv pip install python-pptx requests
```
运行 `python main.py` 查看示例，或调用 `handler(args)`（`args.input.meta` 为 JSON 字符串）。仅使用本地/内嵌图片可不装 `requests`。

## 最小示例 🧪
参考英文 README 的示例（字段一致）。

## 与大模型协作 🤖
1. 准备指标/语料
2. 用 `PROMPT_FOR_LLM.md` 约束生成 JSON
3. （可选）`jsonschema` 校验
4. 送入 `handler` 生成 PPT

## 适用场景 ✅
- 批量日报周报 / KPI 看板
- 模板 / 叙事版本 A/B
- 「自然语言 → 结构化 → 演示」流水线
- Agent / 插件直出 PPT

## 暂不适用 ⛔
- 复杂母版 / 动画 / SmartArt / 多媒体
- 高度自定义高级图表

详见：`LIMITATIONS.md`。

## 扩展点 🔌
搜索：`ELEMENT_TYPES`、`add_chart`、`map_chart_type`、`add_shape`、`apply_run_style`、`apply_paragraph_style`。

## 排查 🔍
- 缺少依赖：安装 `python-pptx`
- 远程图片失败：检查网络 / URL
- 颜色异常：确认十六进制格式
- 图表空白：`series.values` 长度与 `categories` 对齐

## 许可证 & 贡献 🤝
欢迎 Fork 定制内部“自动汇报生成器”。通用增强（新元素 / 图表 / 布局）欢迎 PR。

---
快速上手让模型生成：阅读 `PROMPT_FOR_LLM.md`。
