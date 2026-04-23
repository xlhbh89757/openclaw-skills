import importlib.util
from pathlib import Path
from uuid import uuid4

from openpyxl import load_workbook


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "analyze_resume_v3.py"
TEST_TMP = Path(__file__).resolve().parents[1] / ".test_tmp"


def load_module():
    spec = importlib.util.spec_from_file_location("analyze_resume_v3", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_analyze_risk_keeps_trigger_evidence_for_flags():
    module = load_module()
    parsed = module.parse_resume(
        """
        张三
        工作经验：1年
        2024.01-至今 某某科技有限公司 数据开发工程师
        个人技能
        精通Python，精通Java，精通SQL，全面掌握Hive，专家级Spark，顶级Flink。
        项目经历
        负责数据开发。
        """,
        "张三",
        "张三.pdf",
    )

    analyzed = module.analyze_risk(parsed)

    skill_risk = analyzed["risks"]["skill_exaggeration"]
    assert skill_risk["flags"]
    assert skill_risk["evidence"]
    assert "精通Python" in skill_risk["evidence"][0]


def test_excel_report_includes_evidence_and_parse_quality_columns():
    module = load_module()
    parsed = module.parse_resume(
        """
        李四
        2022.01-至今 测试科技有限公司 ETL工程师
        个人技能
        精通Python，精通Java，精通SQL，全面掌握Hive，专家级Spark，顶级Flink。
        """,
        "李四",
        "李四.pdf",
    )
    result = module.analyze_risk(parsed)
    TEST_TMP.mkdir(exist_ok=True)
    output = TEST_TMP / f"report-{uuid4().hex}.xlsx"

    module.create_excel_report([result], str(output))

    wb = load_workbook(output, data_only=True)
    overview_headers = [cell.value for cell in wb["风险总览"][1]]
    detail_headers = [cell.value for cell in wb["详细分析"][1]]
    assert "解析质量" in overview_headers
    assert "文本长度" in overview_headers
    assert "触发原文" in detail_headers


def test_folder_scan_can_be_limited_to_top_level():
    module = load_module()
    root = TEST_TMP / f"scan-{uuid4().hex}"
    root.mkdir(parents=True, exist_ok=True)
    (root / "top.pdf").write_text("top", encoding="utf-8")
    nested = root / "nested"
    nested.mkdir()
    (nested / "child.pdf").write_text("child", encoding="utf-8")

    recursive = module.scan_resumes_folder(root)
    top_level = module.scan_resumes_folder(root, recursive=False)

    assert {Path(item["filepath"]).name for item in recursive} == {"top.pdf", "child.pdf"}
    assert {Path(item["filepath"]).name for item in top_level} == {"top.pdf"}
