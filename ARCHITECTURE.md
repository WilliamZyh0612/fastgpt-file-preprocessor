# v0.2.0 架构说明

```
只读扫描 -> 解析器 -> 路径/正文/表格上下文 -> AI JSON Schema 校验 -> 规则与词典降级
        -> 哈希/相似度/版本提示 -> 待审核结构化结果 -> JSON + 五个独立 Excel
```

`src/config.py` 读取正式 `config.json`；密钥仅由环境变量提供。`src/parsers.py` 为每条文本保留来源位置；PDF 文字过少时会记录 OCR/视觉流程状态，未配置视觉模型时不会伪造 OCR 内容。`src/ai.py` 调用任何 OpenAI 兼容接口并验证严格 Schema。`src/preprocessor.py` 维护多标签分类、型号/数控系统字典、审核状态、哈希与版本提示。`src/reports.py` 生成五个独立报表。

旧版 Office 的安全转换默认关闭。未来启用时必须在独立进程执行、禁用宏/外链/自动更新、设置文件超时，并在 `finally` 中关闭 Office 进程；本版本不会执行文档内代码。
