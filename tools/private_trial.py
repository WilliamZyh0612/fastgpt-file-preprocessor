"""Run a local-only, auditable pilot over a user-supplied ZIP of material copies.

Neither the ZIP, extracted copies, generated text nor trial outputs are tracked
by Git.  This tool does not upload files or call FastGPT/CRM/ERP.
"""
from __future__ import annotations

import argparse
import csv
import json
import zipfile
from pathlib import Path, PurePosixPath

from src.ai import AIAnalyzer
from src.config import load_settings
from src.preprocessor import UNKNOWN, run


SUPPORTED = {".pdf", ".docx", ".xlsx", ".xlsm", ".pptx", ".txt", ".md", ".markdown", ".doc", ".xls", ".ppt"}


def extract_copies(archive: Path, destination: Path) -> list[Path]:
    if destination.exists(): raise FileExistsError(f"为避免覆盖已有副本，解压目录必须不存在：{destination}")
    # Most Chinese Windows ZIP tools write legacy filenames without the UTF-8
    # flag.  metadata_encoding affects only those entries, not UTF-8 entries.
    with zipfile.ZipFile(archive, metadata_encoding="gbk") as bundle:
        entries = [item for item in bundle.infolist() if not item.is_dir()]
        if any(PurePosixPath(item.filename).is_absolute() or ".." in PurePosixPath(item.filename).parts for item in entries):
            raise ValueError("ZIP 包含不安全路径，已拒绝解压")
        if sum(item.file_size for item in entries) > 1024 * 1024 * 1024: raise ValueError("ZIP 解压后超过 1GB 安全上限")
        selected = [item for item in entries if PurePosixPath(item.filename).suffix.lower() in SUPPORTED]
        destination.mkdir(parents=True)
        for item in selected: bundle.extract(item, destination)
    return [path for path in destination.rglob("*") if path.is_file() and path.suffix.lower() in SUPPORTED]


def non_unknown(value: str | None) -> bool:
    return bool(value and value != UNKNOWN)


def metadata_coverage(record) -> tuple[int, int]:
    values = [
        bool(record.categories), bool(record.machine_models), bool(record.cnc_systems),
        non_unknown(record.version.value if record.version else None), non_unknown(record.release_date.value if record.release_date else None),
        non_unknown(record.visibility.value if record.visibility else None), non_unknown(record.summary.value if record.summary else None),
        bool(record.knowledge_points), bool(record.keywords),
    ]
    return sum(values), len(values)


def pending_confirmation_fields(record) -> int:
    found, total = metadata_coverage(record)
    return total - found


def structured_coverage(records) -> tuple[int, int]:
    detected = total = 0
    for record in records:
        for rows in record.ai_structured.values():
            for row in rows:
                for key, value in row.items():
                    if key in {"evidence_refs", "confidence"}: continue
                    total += 1
                    if value not in {"", UNKNOWN}: detected += 1
    return detected, total


def write_review_template(records, path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as output:
        writer = csv.DictWriter(output, fieldnames=["文件名", "预测分类", "人工确认分类", "预测机型", "人工确认机型", "人工确认结构化字段数", "审核说明"])
        writer.writeheader()
        for record in records:
            writer.writerow({"文件名": record.filename, "预测分类": "、".join(item.name for item in record.categories), "人工确认分类": "", "预测机型": "、".join(item.value for item in record.machine_models), "人工确认机型": "", "人工确认结构化字段数": "", "审核说明": ""})


def main() -> None:
    parser = argparse.ArgumentParser(description="本地真实资料试跑；不上传、不提交任何资料或输出")
    parser.add_argument("--archive", type=Path, default=Path("测试文件.zip"))
    parser.add_argument("--input-copy-dir", type=Path, default=Path("private-test-input"))
    parser.add_argument("--output", type=Path, default=Path("private-test-output"))
    parser.add_argument("--config", type=Path, default=Path("config.json"))
    parser.add_argument("--allow-under-minimum", action="store_true", help="允许少于 10 份资料的预试跑")
    args = parser.parse_args()
    if not args.archive.is_file(): parser.error(f"未找到试跑 ZIP：{args.archive}")
    if args.output.exists(): parser.error(f"为避免覆盖已有输出，输出目录必须不存在：{args.output}")
    copies = extract_copies(args.archive, args.input_copy_dir)
    if not args.allow_under_minimum and not 10 <= len(copies) <= 20:
        parser.error(f"试跑要求 10 至 20 份资料，当前 ZIP 中可处理资料为 {len(copies)} 份；可在确认后使用 --allow-under-minimum 运行预试跑")
    settings = load_settings(args.config); analyzer = AIAnalyzer(settings)
    records = run(args.input_copy_dir, args.output, settings, analyzer)
    metadata_found, metadata_total = map(sum, zip(*(metadata_coverage(record) for record in records))) if records else (0, 0)
    structured_found, structured_total = structured_coverage(records)
    failures = [{"file": record.filename, "reason": "；".join(record.errors) or "未提取到可分析文本"} for record in records if record.extraction_status != "成功"]
    prompt, completion = analyzer.usage["prompt_tokens"], analyzer.usage["completion_tokens"]
    pricing_configured = settings.input_token_price_per_million > 0 or settings.output_token_price_per_million > 0
    estimate = prompt * settings.input_token_price_per_million / 1_000_000 + completion * settings.output_token_price_per_million / 1_000_000 if pricing_configured else None
    metrics = {
        "sample_files": len(records), "sample_size_target": "10-20", "sample_size_status": "合格" if 10 <= len(records) <= 20 else "不足，属于预试跑",
        "parse_success_rate": {"success": sum(record.extraction_status == "成功" for record in records), "total": len(records)},
        "classification_accuracy": {"value": None, "reason": "需在 人工审核标注模板.csv 填写人工确认分类后计算，当前不可编造准确率"},
        "machine_model_accuracy": {"value": None, "reason": "需在 人工审核标注模板.csv 填写人工确认机型后计算，当前不可编造准确率"},
        "metadata_field_coverage": {"recognized": metadata_found, "total": metadata_total},
        "structured_field_recognition_rate": {"recognized": structured_found, "total": structured_total, "note": "这是已生成专业记录的非待确认字段覆盖率，不等同人工准确率"},
        "pending_manual_confirmation": {
            "records_with_pending_fields": sum(pending_confirmation_fields(record) > 0 for record in records),
            "fields": sum(pending_confirmation_fields(record) for record in records),
            "conflict_records": sum(record.review_status == "待人工确认" for record in records),
            "note": "字段数统计元数据中的待人工确认/缺失值；冲突记录是跨来源值冲突导致的待人工确认",
        },
        "failures": failures,
        "token_usage": analyzer.usage,
        "estimated_model_cost": estimate,
        "estimated_model_cost_note": "未配置单价或服务端未返回 usage 时不计算费用",
    }
    (args.output / "trial_metrics.json").write_text(json.dumps(metrics, ensure_ascii=False, indent=2), encoding="utf-8")
    write_review_template(records, args.output / "人工审核标注模板.csv")
    print(json.dumps({key: metrics[key] for key in ("sample_files", "sample_size_status", "parse_success_rate", "metadata_field_coverage", "structured_field_recognition_rate", "pending_manual_confirmation", "token_usage", "estimated_model_cost")}, ensure_ascii=False))


if __name__ == "__main__": main()
