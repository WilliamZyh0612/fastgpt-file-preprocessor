import tempfile
import unittest
from pathlib import Path

from src.preprocessor import UNKNOWN, run


class PipelineTests(unittest.TestCase):
    def test_text_files_are_catalogued_without_source_mutation(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp); source = root / "source"; output = root / "output"; source.mkdir()
            file = source / "VMC850_维护说明.txt"
            original = "机床型号：VMC850\n数控系统：FANUC 0i-MF\n版本：V1.2\n发布日期：2025-01-10\n故障 报警 维修流程。"
            file.write_text(original, encoding="utf-8")
            records = run(source, output)
            self.assertEqual(file.read_text(encoding="utf-8"), original)
            self.assertEqual(records[0].machine_model, "VMC850")
            self.assertTrue((output / "FastGPT知识库总目录.xlsx").exists())
            self.assertTrue((output / "analysis.json").exists())

    def test_missing_fields_are_never_invented(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp); source = root / "source"; source.mkdir()
            (source / "note.md").write_text("这是一般工作笔记。", encoding="utf-8")
            record = run(source, root / "output")[0]
            self.assertEqual(record.machine_model, UNKNOWN)
            self.assertEqual(record.version.value, UNKNOWN)


if __name__ == "__main__": unittest.main()
