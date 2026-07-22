from __future__ import annotations

import argparse
import hashlib
import json
import re
from datetime import datetime
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass, field
from pathlib import Path

from .ai import AIAnalyzer, AIResponseError
from .config import Settings, load_settings
from .dictionaries import CATEGORY_RULES, CNC_BRANDS, MACHINE_MODELS
from .parsers import Evidence, Extracted, SUPPORTED, evidence, extract, sha256
from .reports import write_reports

STOPWORDS = {"的", "和", "及", "与", "在", "是", "为", "了", "本", "该", "产品", "文件", "说明", "manual", "the", "and", "for", "with", "this"}
REVIEW_PENDING = "待审核"
UNKNOWN = "待人工确认"  # v0.1 compatibility constant; configured values are used at runtime.


@dataclass
class Classification:
    name: str
    confidence: float
    reason: str
    evidence: Evidence


@dataclass
class FieldValue:
    value: str
    evidence: list[Evidence] = field(default_factory=list)


@dataclass
class Record:
    path: str
    filename: str
    extension: str
    sha256: str
    size_bytes: int
    normalized_text_hash: str = ""
    categories: list[Classification] = field(default_factory=list)
    machine_models: list[FieldValue] = field(default_factory=list)
    cnc_systems: list[dict] = field(default_factory=list)
    version: FieldValue | None = None
    release_date: FieldValue | None = None
    visibility: FieldValue | None = None
    summary: FieldValue | None = None
    knowledge_points: list[FieldValue] = field(default_factory=list)
    keywords: list[str] = field(default_factory=list)
    archive_suggestion: str = ""
    review_status: str = REVIEW_PENDING
    extraction_status: str = "成功"
    flags: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    text: str = field(default="", repr=False)
    evidence: list[Evidence] = field(default_factory=list, repr=False)
    tables: list[dict] = field(default_factory=list, repr=False)
    ai_structured: dict = field(default_factory=dict)
    conflicts: list[dict] = field(default_factory=list)

    @property
    def machine_model(self) -> str:
        return self.machine_models[0].value if self.machine_models else UNKNOWN

    @property
    def category(self) -> str:
        return self.categories[0].name if self.categories else UNKNOWN


def unknown(settings: Settings, evidence: list[Evidence] | None = None) -> FieldValue:
    return FieldValue(settings.unknown_value, evidence or [])


def find_models(context: str, evidence: list[Evidence], settings: Settings) -> list[FieldValue]:
    aliases = {**MACHINE_MODELS, **settings.model_aliases}; hits: list[tuple[str, str]] = []
    for canonical, names in aliases.items():
        for alias in names:
            if re.search(rf"(?<![A-Z0-9-]){re.escape(alias)}(?![A-Z0-9-])", context, re.I): hits.append((canonical, alias)); break
    chosen: list[tuple[str, str]] = []
    for canonical, alias in sorted(hits, key=lambda item: len(item[1]), reverse=True):
        if canonical not in [item[0] for item in chosen] and not any(alias.upper() in longer.upper() for _, longer in chosen): chosen.append((canonical, alias))
    return [FieldValue(canonical, [next((e for e in evidence if re.search(rf"(?<![A-Z0-9-]){re.escape(alias)}(?![A-Z0-9-])", e.excerpt, re.I)), Evidence("name-context", "路径或文件名", "名称", alias))]) for canonical, alias in chosen]


def find_cnc(context: str, evidence: list[Evidence], settings: Settings) -> list[dict]:
    aliases = {**CNC_BRANDS, **settings.cnc_aliases}; result = []
    for brand, names in aliases.items():
        matched = next((name for name in names if name.upper() in context.upper()), None)
        if matched:
            model = re.search(rf"{re.escape(matched)}\s*([0-9A-Z][0-9A-Z ._-]{{0,20}})?", context, re.I)
            item = model.group(0).strip() if model else settings.unknown_value
            source = next((e for e in evidence if matched.upper() in e.excerpt.upper()), Evidence("name-context", "路径或文件名", "名称", matched))
            result.append({"brand": brand, "model": item, "version": settings.unknown_value, "evidence": [asdict(source)]})
    return result


def first_regex(pattern: str, context: str, evidence: list[Evidence], settings: Settings) -> FieldValue:
    matched = re.search(pattern, context, flags=re.I)
    if not matched: return unknown(settings)
    value = matched.group(1).strip(); source = next((e for e in evidence if value.lower() in e.excerpt.lower()), Evidence("name-context", "路径或文件名", "名称", value))
    return FieldValue(value, [source])


def split_sentences(text: str) -> list[str]:
    clean = re.sub(r"\s+", " ", text).strip()
    # A dot between digits belongs to V1.2 / 0.01 / 3.14, never a sentence boundary.
    parts = re.split(r"[。！？]+|(?<!\d)[.!?](?!\d)(?:\s+|$)", clean)
    return [part.strip() for part in parts if len(part.strip()) >= 10]


def summarize(text: str, evidence: list[Evidence], settings: Settings) -> tuple[FieldValue, list[FieldValue], list[str]]:
    sentences = split_sentences(text)
    if not sentences: return unknown(settings), [unknown(settings)], []
    ranked = sorted(sentences, key=lambda item: (sum(word in item.lower() for words in CATEGORY_RULES.values() for word in words), len(item)), reverse=True)
    selected = ranked[0][:300]
    source = next((e for e in evidence if selected[:20] in e.excerpt), evidence[:1][0] if evidence else globals()["evidence"]("", "", ""))
    words = re.findall(r"[A-Za-z][A-Za-z0-9+._-]{2,}|[\u4e00-\u9fff]{2,}", text)
    keys = [word for word, _ in Counter(word for word in words if word.lower() not in STOPWORDS).most_common(10)]
    return FieldValue(selected, [source]), [FieldValue(item[:300], [next((e for e in evidence if item[:16] in e.excerpt), source)]) for item in ranked[:5]], keys


def rule_categories(context: str, evidence: list[Evidence]) -> list[Classification]:
    found = []
    for name, terms in CATEGORY_RULES.items():
        hits = [term for term in terms if term.lower() in context.lower()]
        if hits:
            source = next((e for e in evidence if any(term.lower() in e.excerpt.lower() for term in hits)), globals()["evidence"]("路径或文件名", "名称", hits[0]))
            found.append(Classification(name, min(0.95, 0.55 + .1 * len(hits)), f"命中关键词：{'、'.join(hits)}", source))
    return found


def apply_ai(record: Record, analyzer: AIAnalyzer, context: str, settings: Settings, chunk: list[Evidence]) -> None:
    refs = [asdict(item) for item in chunk]
    try: result = analyzer.analyze(context, refs)
    except AIResponseError as exc: record.errors.append(str(exc)); return
    if not result: return
    for key in ("product_records", "fault_records", "part_records", "business_rule_records"):
        record.ai_structured.setdefault(key, []).extend(result[key])
    meta = result["document_metadata"]; lookup = {item.evidence_id:item for item in chunk}
    def cited(ref_ids: list[str]) -> list[Evidence]:
        # validate_ai_result already checks existence; this second lookup protects
        # against accidental use of evidence from a different document chunk.
        if not ref_ids or not set(ref_ids).issubset(lookup): raise AIResponseError("AI 字段引用了当前分块以外的证据")
        return [lookup[item] for item in ref_ids]
    categories = [Classification(item["name"], float(item["confidence"]), item["reason"], lookup[item["evidence_refs"][0]]) for item in meta["categories"] if item["name"] != settings.unknown_value and item["evidence_refs"]]
    if categories:
        rule_names={x.name for x in record.categories}; ai_names={x.name for x in categories}
        if rule_names and rule_names != ai_names: record.conflicts.append({"field":"categories","rule_value":sorted(rule_names),"ai_value":sorted(ai_names)}); record.review_status="待人工确认"
        record.categories=categories
    if meta["machine_models"]:
        ai_models=[FieldValue(item["value"], cited(item["evidence_refs"])) for item in meta["machine_models"] if item["value"] != settings.unknown_value]
        if ai_models and record.machine_models and {x.value for x in ai_models}!={x.value for x in record.machine_models}: record.conflicts.append({"field":"machine_models","rule_value":[x.value for x in record.machine_models],"ai_value":[x.value for x in ai_models]}); record.review_status="待人工确认"
        if ai_models: record.machine_models=ai_models
    for field in ("visibility","version","release_date","summary"):
        value=meta[field]["value"]
        if value!=settings.unknown_value:
            current=getattr(record,field); ai_field=FieldValue(value, cited(meta[field]["evidence_refs"]))
            if current and current.value not in {settings.unknown_value,value}: record.conflicts.append({"field":field,"rule_value":current.value,"ai_value":value}); record.review_status="待人工确认"
            setattr(record,field,ai_field)
    if meta["cnc_systems"]:
        record.cnc_systems=[{**item, "evidence":[asdict(source) for source in cited(item["evidence_refs"])]} for item in meta["cnc_systems"]]
    if meta["knowledge_points"]: record.knowledge_points=[FieldValue(item["value"], cited(item["evidence_refs"])) for item in meta["knowledge_points"]]
    if meta["keywords"]: record.keywords=[item["value"] for item in meta["keywords"][:10]]


def _aggregate_ai_structured(record: Record) -> None:
    """Merge equal records across chunks and explicitly flag contradictory ones."""
    identity = {
        "product_records": ("product_model", "product_name"),
        "fault_records": ("machine_model", "alarm_code", "fault_symptom"),
        "part_records": ("part_model", "drawing_number", "part_name"),
        "business_rule_records": ("business_module", "rule_name"),
    }
    for key, items in record.ai_structured.items():
        unique: dict[str, dict] = {}
        for item in items:
            signature = json.dumps({name:value for name, value in item.items() if name not in {"evidence_refs", "confidence"}}, ensure_ascii=False, sort_keys=True)
            if signature in unique:
                merged = unique[signature]
                merged["evidence_refs"] = sorted(set(merged["evidence_refs"]) | set(item["evidence_refs"]))
                merged["confidence"] = max(float(merged["confidence"]), float(item["confidence"]))
            else:
                unique[signature] = item
        merged_items = list(unique.values())
        record.ai_structured[key] = merged_items
        by_identity: dict[tuple[str, ...], list[dict]] = {}
        for item in merged_items:
            token = tuple(item[name] for name in identity[key] if item[name] != "待人工确认")
            if token: by_identity.setdefault(token, []).append(item)
        for token, variants in by_identity.items():
            variants_without_refs = {json.dumps({name:value for name, value in item.items() if name not in {"evidence_refs", "confidence"}}, ensure_ascii=False, sort_keys=True) for item in variants}
            if len(variants_without_refs) > 1:
                record.conflicts.append({"field":key, "rule_value":"分块结果", "ai_value":f"同一标识 {token} 存在冲突记录"})
                record.review_status = "待人工确认"


def analyze(path: Path, settings: Settings, analyzer: AIAnalyzer | None = None) -> Record:
    record = Record(str(path.resolve()), path.name, path.suffix.lower(), sha256(path), path.stat().st_size)
    name_context = " / ".join((*path.parts[-4:], path.name))
    try:
        parsed = extract(path, settings); record.text, record.evidence, record.tables = parsed.text, parsed.evidence, parsed.tables
        record.errors.extend(parsed.errors)
        if not parsed.text.strip():
            record.extraction_status = "需人工处理"; record.flags.append("未提取到可分析文本")
            return record
        context = name_context + "\n" + parsed.text
        normalized = re.sub(r"\s+", "", parsed.text).lower(); record.normalized_text_hash = hashlib.sha256(normalized.encode("utf-8")).hexdigest()
        record.machine_models = find_models(context, parsed.evidence, settings)
        record.cnc_systems = find_cnc(context, parsed.evidence, settings)
        record.categories = rule_categories(context, parsed.evidence)
        record.version = first_regex(r"(?:版本|version|ver\.)\s*[:：]?\s*(v?\d+(?:\.\d+){0,3}(?:[-_A-Za-z0-9]+)?)", context, parsed.evidence, settings)
        record.release_date = first_regex(r"(?:发布日期|发布|日期|date)\s*[:：]?\s*((?:20\d{2}|19\d{2})[-/.年]\d{1,2}[-/.月]\d{1,2}日?)", context, parsed.evidence, settings)
        record.visibility = first_regex(r"(?:可见范围|保密级别|visibility)\s*[:：]?\s*(内部|公开|销售|售后|管理层|客户|public|internal)", context, parsed.evidence, settings)
        record.summary, record.knowledge_points, record.keywords = summarize(parsed.text, parsed.evidence, settings)
        if analyzer:
            for chunk in parsed.chunks or [parsed.evidence]:
                apply_ai(record, analyzer, "\n".join(item.excerpt for item in chunk), settings, chunk)
            _aggregate_ai_structured(record)
        category = record.categories[0].name if record.categories else "待分类"; model = record.machine_models[0].value if record.machine_models else settings.unknown_value
        record.archive_suggestion = f"/{category}/{model}/{record.release_date.value if record.release_date else settings.unknown_value}"
    except Exception as exc:
        record.extraction_status = "需人工处理"; record.errors.append(f"解析失败：{type(exc).__name__}: {str(exc)[:300]}")
    return record


def _tokens(text: str) -> set[str]: return {item.lower() for item in re.findall(r"[A-Za-z0-9\u4e00-\u9fff]{2,}", text) if item.lower() not in STOPWORDS}


def enrich(records: list[Record], settings: Settings, analyzer: AIAnalyzer | None = None) -> None:
    by_hash: dict[str, list[Record]] = {}; by_text: dict[str, list[Record]] = {}
    for record in records:
        by_hash.setdefault(record.sha256, []).append(record)
        if record.normalized_text_hash: by_text.setdefault(record.normalized_text_hash, []).append(record)
    for group in by_hash.values():
        if len(group) > 1:
            for record in group: record.flags.append("完全重复文件（SHA-256）")
    for group in by_text.values():
        if group[0] and len(group) > 1:
            for record in group: record.flags.append("标准化文本重复")
    vectors = None
    if analyzer:
        try: vectors = analyzer.embeddings([record.text for record in records])
        except Exception as exc:
            for record in records: record.errors.append(f"语义向量不可用：{type(exc).__name__}")
    for index, left in enumerate(records):
        a = _tokens(left.text)
        for right_index, right in enumerate(records[index + 1:], index + 1):
            b = _tokens(right.text); score = len(a & b) / max(1, len(a | b))
            if vectors:
                first, second = vectors[index], vectors[right_index]
                dot = sum(x * y for x, y in zip(first, second)); left_norm = sum(x * x for x in first) ** .5; right_norm = sum(y * y for y in second) ** .5
                score = dot / max(1e-12, left_norm * right_norm); label = "语义"
            else: label = "文本"
            if left.sha256 != right.sha256 and score >= settings.similarity_threshold:
                left.flags.append(f"相似文件：{right.filename}（{label} {score:.0%}）"); right.flags.append(f"相似文件：{left.filename}（{label} {score:.0%}）")
    groups: dict[tuple[str, str], list[Record]] = {}
    for item in records:
        for category in item.categories:
            for model in item.machine_models: groups.setdefault((category.name, model.value), []).append(item)
    for group in groups.values():
        versions = {item.version.value for item in group if item.version and item.version.value != settings.unknown_value}
        if len(versions) > 1:
            for item in group: item.flags.append("版本冲突：同类同型号存在多个版本")
        dates = sorted(((datetime.strptime(item.release_date.value.replace("/", "-").replace("年", "-").replace("月", "-").replace("日", ""), "%Y-%m-%d"), item) for item in group if item.release_date and item.release_date.value != settings.unknown_value and re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{4}年\d{1,2}月\d{1,2}日?", item.release_date.value)), key=lambda pair: pair[0])
        if len(dates) > 1:
            for item in group:
                if item.release_date and dates and item is dates[0][1]: item.flags.append("疑似过期资料：存在较新发布日期")
    for record in records:
        if not record.flags: record.flags.append("无")


def run(source: Path, output: Path, settings: Settings | None = None, analyzer: AIAnalyzer | None = None) -> list[Record]:
    settings = settings or load_settings(Path(__file__).resolve().parent.parent / "config.json"); source, output = source.resolve(), output.resolve()
    files = [path for path in source.rglob("*") if path.is_file() and path.suffix.lower() in SUPPORTED and not path.is_relative_to(output)]
    if analyzer is None: analyzer = AIAnalyzer(settings)
    with ThreadPoolExecutor(max_workers=settings.concurrency) as executor:
        records = list(executor.map(lambda item: analyze(item, settings, analyzer), files))
    enrich(records, settings, analyzer); output.mkdir(parents=True, exist_ok=True)
    write_reports(records, output, settings)
    def public(record: Record) -> dict:
        data = asdict(record); data.pop("text", None); return data
    (output / "analysis.json").write_text(json.dumps([public(record) for record in records], ensure_ascii=False, indent=2), encoding="utf-8")
    return records


def main() -> None:
    parser = argparse.ArgumentParser(description="FastGPT 独立 AI 文件预处理器（只读源文件）")
    parser.add_argument("--input", required=True, type=Path); parser.add_argument("--output", required=True, type=Path); parser.add_argument("--config", type=Path, default=Path("config.json"))
    args = parser.parse_args()
    if not args.input.is_dir(): parser.error("--input 必须是存在的目录")
    if args.output.resolve().is_relative_to(args.input.resolve()): parser.error("--output 不得位于 --input 内")
    records = run(args.input, args.output, load_settings(args.config)); print(f"完成：处理 {len(records)} 个文件；输出：{args.output.resolve()}")


if __name__ == "__main__": main()
