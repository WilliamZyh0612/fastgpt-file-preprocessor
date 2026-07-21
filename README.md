# FastGPT 文件预处理器

独立、本地、只读的 FastGPT 知识库资料预处理工具。它不修改 FastGPT 核心代码，也不会修改、移动或删除原始资料。

## 能力边界

- 批量读取 PDF、Word（DOC/DOCX）、Excel（XLS/XLSX/XLSM）、PPT（PPT/PPTX）、TXT、Markdown。旧版 Office 格式通过本机安装的 Microsoft Office 以只读模式提取；如未安装，会明确标记为需人工处理而不会编造结果。
- 仅根据提取到的文本证据识别分类、机床型号、数控系统、版本、发布日期和可见范围；没有明确证据即输出 **待人工确认**。
- 生成摘要、知识点、关键词、归档建议，以及重复/相似/版本冲突/疑似过期标记。
- 生成 `FastGPT知识库总目录.xlsx`：包含知识库总目录、产品选型矩阵、故障案例库、配件适配表、CRM/ERP 业务规则五个工作表。

## 部署

要求 Python 3.11 或更新版本。

```powershell
cd E:\Codex工作区\FastGPT文件预处理器
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## 使用

输出目录必须在资料目录之外，避免生成文件被再次扫描。

```powershell
.\run.ps1 -Input 'D:\资料待整理' -Output 'D:\FastGPT预处理输出'
```

输出包括：

- `FastGPT知识库总目录.xlsx`
- `analysis.json`（供后续导入 FastGPT 或自定义流程使用）

## 测试与打包

```powershell
$env:PYTHONPATH='.'
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
.\build_package.ps1
```

部署包输出为 `dist\fastgpt-file-preprocessor.zip`。

## 审计原则

所有分析结果带有原始绝对路径与 SHA-256。相似性是基于文本 Token 的 Jaccard 指数，阈值为 72%；它是提示，不会自动合并、覆盖或删除任何资料。
