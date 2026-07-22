import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from dataclasses import asdict
from pathlib import Path

from openpyxl import load_workbook

from src.ai import AI_SCHEMA, AIAnalyzer, validate_ai_result
from src.config import Settings, load_settings
from src.parsers import _chunk_evidence, evidence, extract
from src.preprocessor import Record, _finalize_document_metadata, analyze, apply_ai, enrich
from src.reports import write_reports


def structured_result(evidence_id: str, product_name: str = "BK5030 选型", *, category: str = "产品资料", model: str = "BK5030", version: str = "V1.2", date: str = "2025-01-01", visibility: str = "内部", summary: str = "BK5030 产品参数说明", keyword: str = "BK5030", knowledge: str = "加工范围 500mm", cnc_brand: str = "FANUC") -> dict:
    def field(value: str) -> dict:
        return {"value": value, "confidence": .9, "evidence_refs": [evidence_id]}
    metadata = {
        "categories": [{"name": category, "confidence": .9, "reason": "参数表", "evidence_refs": [evidence_id]}],
        "machine_models": [field(model)],
        "cnc_systems": [{"brand": cnc_brand, "model": "0i-MF", "version": "待人工确认", "confidence": .8, "evidence_refs": [evidence_id]}],
        "visibility": field(visibility), "version": field(version), "release_date": field(date),
        "summary": field(summary), "knowledge_points": [field(knowledge)], "keywords": [field(keyword)],
    }
    product = {"product_model": "BK5030", "product_name": product_name, "processing_object": "齿轮", "processing_range": "500mm", "key_parameters": "行程 500mm", "accuracy": "0.01mm", "standard_configuration": "标准夹具", "optional_configuration": "自动上下料", "applicable_scenarios": "批量加工", "limitations": "需人工确认", "cnc_system": "FANUC 0i-MF", "evidence_refs": [evidence_id], "confidence": .9}
    return {"document_metadata": metadata, "product_records": [product], "fault_records": [], "part_records": [], "business_rule_records": []}


class FixTests(unittest.TestCase):
    def test_strict_schema_is_complete_and_payload_is_valid(self):
        schema = AI_SCHEMA["schema"]
        self.assertEqual(set(schema["required"]), set(schema["properties"]))
        def inspect(node):
            if not isinstance(node, dict): return
            if node.get("type") == "object":
                self.assertTrue(node.get("properties"))
                self.assertEqual(set(node["required"]), set(node["properties"]))
                self.assertIs(node["additionalProperties"], False)
            if node.get("type") == "array": self.assertIn("items", node)
            for value in node.values():
                if isinstance(value, dict): inspect(value)
                elif isinstance(value, list):
                    for item in value: inspect(item)
        inspect(schema)
        payload = AIAnalyzer(Settings(ai_enabled=True, api_base_url="http://x/v1", text_model="m"))._payload("x", [])
        self.assertEqual(payload["response_format"], {"type": "json_schema", "json_schema": AI_SCHEMA})

    def test_last_chunk_and_configured_overlap_are_effective(self):
        parts = [evidence("x", str(i), "a" * 8) for i in range(3)]
        chunks = _chunk_evidence(parts, 12, 4)
        self.assertTrue("".join(item.excerpt for item in chunks[-1]).endswith(parts[-1].excerpt[-4:]))
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "long.txt"; path.write_text("a" * 1200, encoding="utf-8")
            config = Path(tmp) / "config.json"; config.write_text(json.dumps({"chunk_size_chars": 500, "chunk_overlap_chars": 100}), encoding="utf-8")
            parsed = extract(path, load_settings(config))
            self.assertGreater(len(parsed.chunks), 2)
            self.assertEqual(parsed.chunks[0][0].excerpt[-100:], parsed.chunks[1][0].excerpt[:100])

    def test_long_txt_uses_isolated_chunk_evidence_and_aggregates_records(self):
        calls: list[dict] = []
        def transport(payload):
            source = json.loads(payload["messages"][1]["content"]); calls.append(source)
            return {"choices": [{"message": {"content": json.dumps(structured_result(source["evidence"][0]["evidence_id"], f"记录-{len(calls)}"), ensure_ascii=False)}}]}
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "long.txt"; path.write_text("BK5030 产品参数。" * 200, encoding="utf-8")
            record = analyze(path, Settings(ai_enabled=True, api_base_url="http://x/v1", text_model="m", chunk_size_chars=500, chunk_overlap_chars=100), AIAnalyzer(Settings(ai_enabled=True, api_base_url="http://x/v1", text_model="m", chunk_size_chars=500, chunk_overlap_chars=100), transport))
        self.assertGreater(len(calls), 1)
        self.assertEqual(len(record.ai_structured["product_records"]), len(calls))
        for source in calls:
            self.assertLessEqual(sum(len(item["excerpt"]) for item in source["evidence"]), 500)
            self.assertTrue(all(item["excerpt"] in source["context"] for item in source["evidence"]))

    def test_document_metadata_is_merged_across_chunks(self):
        first, second = evidence("x", "第1页", "first"), evidence("x", "第2页", "second")
        def transport(payload):
            current = json.loads(payload["messages"][1]["content"]); ref = current["evidence"][0]["evidence_id"]
            if "second" in current["context"]:
                value = structured_result(ref, category="操作资料", model="TCK56", version="V2.0", date="2025-02-01", visibility="公开", summary="第二分块摘要", keyword="操作", knowledge="安全操作", cnc_brand="西门子")
            else:
                value = structured_result(ref, category="产品资料", model="BK5030", version="V1.0", date="2025-01-01", visibility="内部", summary="第一分块摘要", keyword="参数", knowledge="加工范围")
            return {"choices": [{"message": {"content": json.dumps(value, ensure_ascii=False)}}]}
        settings = Settings(ai_enabled=True, api_base_url="http://x/v1", text_model="m")
        record = Record("x", "x.txt", ".txt", "hash", 1, evidence=[first, second])
        analyzer = AIAnalyzer(settings, transport)
        apply_ai(record, analyzer, "first", settings, [first]); apply_ai(record, analyzer, "second", settings, [second]); _finalize_document_metadata(record, settings)
        self.assertEqual({item.name for item in record.categories}, {"产品资料", "操作资料"})
        self.assertEqual({item.value for item in record.machine_models}, {"BK5030", "TCK56"})
        self.assertEqual({item["brand"] for item in record.cnc_systems}, {"FANUC", "西门子"})
        self.assertEqual(set(record.keywords), {"参数", "操作"}); self.assertEqual({item.value for item in record.knowledge_points}, {"加工范围", "安全操作"})
        self.assertEqual(record.version.value, "待人工确认"); self.assertEqual({item.value for item in record.metadata_candidates["version"]}, {"V1.0", "V2.0"})
        self.assertEqual(record.release_date.value, "待人工确认"); self.assertEqual(record.visibility.value, "待人工确认")
        self.assertIn("第一分块摘要", record.summary.value); self.assertIn("第二分块摘要", record.summary.value)

    def test_chunk_overlap_is_exact_and_chunk_size_is_never_exceeded(self):
        chunks = _chunk_evidence([evidence("x", "正文", "abcdefghijklmnopqrstuvwxyz")], 10, 3)
        text = lambda group: "".join(item.excerpt for item in group)
        self.assertTrue(all(len(text(group)) <= 10 for group in chunks))
        self.assertEqual(text(chunks[0])[-3:], text(chunks[1])[:3])
        self.assertEqual(text(chunks[1])[-3:], text(chunks[2])[:3])

    def test_ai_metadata_requires_current_nonempty_evidence(self):
        item = evidence("x", "第1页", "BK5030 产品参数")
        result = structured_result(item.evidence_id)
        validate_ai_result(result, {item.evidence_id})
        meta = result["document_metadata"]
        for field in ("visibility", "version", "release_date", "summary"):
            self.assertTrue(meta[field]["evidence_refs"])
        for field in ("machine_models", "knowledge_points", "keywords", "cnc_systems"):
            self.assertTrue(meta[field][0]["evidence_refs"])
        bad = structured_result(item.evidence_id); bad["document_metadata"]["summary"]["evidence_refs"] = []
        with self.assertRaises(ValueError): validate_ai_result(bad, {item.evidence_id})
        wrong = structured_result(item.evidence_id); wrong["document_metadata"]["version"]["evidence_refs"] = ["0" * 16]
        with self.assertRaises(ValueError): validate_ai_result(wrong, {item.evidence_id})

    def test_empty_text_is_not_standardized_text_duplicate(self):
        left = Record("a", "a.txt", ".txt", "a", 0); right = Record("b", "b.txt", ".txt", "b", 0)
        enrich([left, right], Settings())
        self.assertNotIn("标准化文本重复", left.flags)
        self.assertNotIn("标准化文本重复", right.flags)

    def test_professional_excel_uses_structured_record_fields(self):
        item = evidence("manual.pdf", "PDF 第2页", "BK5030 加工范围 500mm")
        record = Record("manual.pdf", "manual.pdf", ".pdf", "hash", 10)
        record.evidence = [item]; record.ai_structured = structured_result(item.evidence_id)["product_records"] and {"product_records": structured_result(item.evidence_id)["product_records"]}
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp); write_reports([record], output, Settings())
            book = load_workbook(output / "产品选型矩阵.xlsx", read_only=True); sheet = book.active
            row = list(sheet.iter_rows(min_row=2, max_row=2, values_only=True))[0]; book.close()
            self.assertEqual(row[0], "BK5030"); self.assertEqual(row[1], "BK5030 选型")
            self.assertEqual(row[3], "500mm"); self.assertIn("PDF 第2页", row[-5])

    def test_real_package_can_be_extracted_and_started(self):
        root = Path(__file__).parents[1]
        subprocess.run([sys.executable, str(root / "tools" / "package.py")], cwd=root, check=True)
        archive = root / "dist" / "fastgpt-file-preprocessor.zip"
        with tempfile.TemporaryDirectory() as tmp:
            with zipfile.ZipFile(archive) as bundle: bundle.extractall(tmp)
            package = Path(tmp) / "fastgpt-file-preprocessor"; input_dir = Path(tmp) / "input"; output = Path(tmp) / "output"; input_dir.mkdir()
            input_dir.joinpath("smoke.txt").write_text("BK5030 产品参数", encoding="utf-8")
            environment = {**os.environ, "PYTHONPATH": str(package)}
            subprocess.run([sys.executable, "-m", "src.preprocessor", "--input", str(input_dir), "--output", str(output), "--config", str(package / "config.json")], cwd=package, env=environment, check=True)
            self.assertTrue((output / "analysis.json").exists())


if __name__ == "__main__": unittest.main()
