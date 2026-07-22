# v0.2.0 架构说明

```
只读扫描 -> 解析器 -> 路径/正文/表格上下文 -> AI JSON Schema 校验 -> 规则与词典降级
        -> 哈希/相似度/版本提示 -> 待审核结构化结果 -> JSON + 五个独立 Excel
```

`src/config.py` 读取正式 `config.json`；密钥仅由环境变量提供。`src/parsers.py` 为每条文本保留来源位置；PDF 文字过少时仅检测并标记为待处理。**本版本不具备 OCR 能力**：不渲染页面、不调用视觉模型、也不伪造 OCR 原文。`src/ai.py` 调用任何 OpenAI 兼容接口并验证严格 Schema。`src/preprocessor.py` 维护多标签分类、型号/数控系统字典、审核状态、哈希与版本提示。`src/reports.py` 生成五个独立报表。

旧版 Office 的安全转换默认关闭。启用时在独立进程使用 `DispatchEx` 创建 Office、禁用宏/外链/自动更新，并记录该任务创建的 Office PID；发生超时时仅终止这个 PID，绝不枚举或关闭用户已打开的 Office。`tools/verify_legacy_office_timeout_windows.py` 可在 Windows 上针对一份可丢弃的 DOC/XLS/PPT 副本审计超时清理和残留 PID。本版本不会执行文档内代码。
