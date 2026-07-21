# FastGPT 独立 AI 文件预处理器

`v0.2.0` 是本地、只读的知识库资料预处理工具：它不修改 FastGPT、不连接 CRM/ERP、不上传 FastGPT，也不会移动、覆盖或删除原始资料。

## 安全与审核边界

- 只扫描 PDF、DOCX、XLSX/XLSM、PPTX、TXT、Markdown；DOC/XLS/PPT 默认标记为“需安全转换”，不会运行宏、VBA、ActiveX 或嵌入对象。
- 每份资料都保留 SHA-256、标准化文本哈希、来源路径，以及页码/工作表/标题层级等证据。
- 任何缺失或不能由证据支撑的信息都输出“待人工确认”。结果默认“待审核”；只有人工改为“已确认”后，才应由未来的导入程序使用。
- API Key 仅从 `FASTGPT_PREPROCESSOR_API_KEY`（或配置的环境变量名）读取，绝不写入配置、日志、`analysis.json` 或 Excel。

## 配置与 AI

复制并编辑 `config.json`。`api_base_url` 采用 OpenAI 兼容的 `/v1` 地址；`text_model`、`vision_model`、超时、重试、并发、文件限制和相似度阈值均可配置。默认 `ai_enabled=false`，所以离线测试不会外发资料。

```powershell
$env:FASTGPT_PREPROCESSOR_API_KEY = '仅写入当前终端'
# 编辑 config.json：设置 api_base_url、text_model，并将 ai_enabled 改为 true
```

AI 层要求 JSON Schema；非法 JSON、字段不合法或超时会重试，然后安全降级到规则与字典识别。AI 结合路径、目录、文件名、正文、标题和表格文本判断，可返回多标签分类及每项置信度、理由和证据。

## 安装与运行

```powershell
cd E:\Codex工作区\FastGPT文件预处理器
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
.\run.ps1 -Input 'D:\待处理资料' -Output 'D:\预处理结果'
```

输出目录必须在输入目录之外。输出为 `analysis.json` 与五个独立 Excel：

1. `FastGPT知识库总目录.xlsx`
2. `产品选型矩阵.xlsx`
3. `故障案例库.xlsx`
4. `配件适配表.xlsx`
5. `CRM_ERP业务规则.xlsx`

## 测试与测试包

```powershell
$env:PYTHONPATH='.'
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\build_package.ps1
```

测试包会生成在 `dist\fastgpt-file-preprocessor.zip`。详见 [架构说明](ARCHITECTURE.md)。
