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


def test_colon_section_headings_are_parsed():
    module = load_module()
    parsed = module.parse_resume(
        """
        王五
        工作经历：
        2021.08-2023.12 测试科技有限公司 数据开发工程师
        项目经历：
        客户经营分析平台 2022.01-2022.12
        专业技能：
        SQL、Hive、Python
        """,
        "王五",
        "王五.pdf",
    )

    assert parsed["work_experience"] == ["2021.08-2023.12 测试科技有限公司 数据开发工程师"]
    assert parsed["project_experience"]
    assert parsed["skills"] == ["SQL、Hive、Python"]


def test_project_lines_are_not_inferred_as_work_history():
    module = load_module()
    parsed = module.parse_resume(
        """
        赵六
        项目经历
        2023-11 ~ 2024-06 邢台家乐园集团超市有限责任公司数据中台 数据开发工程师
        项目职责：负责ODS、DWD建模。
        专业技能
        SQL、Hive、Kettle
        """,
        "赵六",
        "赵六.pdf",
    )

    assert parsed["work_experience"] == []
    assert parsed["project_experience"]


def test_project_timeline_lines_do_not_trigger_work_overlap():
    module = load_module()
    analyzed = module.analyze_risk(
        {
            "name": "钱七",
            "filename": "钱七.pdf",
            "raw_text": "钱七",
            "clean_text": "钱七",
            "education": ["2018.09-2022.06 测试大学 计算机 本科"],
            "skills": ["SQL、Hive"],
            "project_experience": [],
            "work_experience": [
                "2021.01-2024.12 测试科技有限公司 数据开发工程师",
                "2023.02~2024.10 广发银行星轨个贷风控系统开发",
                "2022.07-2023.08 银川隆基锂电池生产数据整合与分析支撑项目 数据开发工程师",
            ],
        }
    )

    assert analyzed["risks"]["timeline"]["flags"] == []


def test_extraction_debug_files_include_raw_text_and_sections():
    module = load_module()
    parsed = module.parse_resume(
        """
        孙八
        教育背景
        2016.09-2020.06 测试大学 计算机 本科
        工作经历
        2020.07-2024.01 测试科技有限公司 数据开发工程师
        项目经历
        客户经营分析平台 负责数据建模
        专业技能
        SQL、Hive、Python
        """,
        "孙八",
        "数据开发-孙八.pdf",
    )
    output_dir = TEST_TMP / f"debug-{uuid4().hex}"

    written = module.write_extraction_debug_files(parsed, output_dir, 1)

    raw_path = output_dir / "001-数据开发-孙八.txt"
    sections_path = output_dir / "001-数据开发-孙八.sections.json"
    assert written == {"text": str(raw_path), "sections": str(sections_path)}
    assert "测试科技有限公司" in raw_path.read_text(encoding="utf-8")

    import json

    sections = json.loads(sections_path.read_text(encoding="utf-8"))
    assert sections["filename"] == "数据开发-孙八.pdf"
    assert sections["sections"]["work_experience"] == ["2020.07-2024.01 测试科技有限公司 数据开发工程师"]
    assert sections["parse_quality"]["quality"] in {"正常", "需复核", "较差"}


def test_relative_report_and_debug_paths_resolve_under_folder():
    module = load_module()
    folder = Path(r"E:\简历")

    output_path, debug_dir = module.resolve_output_locations(
        output_arg="custom_report.xlsx",
        folder_arg=str(folder),
        file_arg=None,
        debug_arg=None,
        save_extracted_text="extraction_debug",
    )

    assert output_path == str(folder / "custom_report.xlsx")
    assert debug_dir == str(folder / "extraction_debug")


def test_default_report_path_resolves_under_single_file_parent():
    module = load_module()
    file_path = Path(r"E:\简历\数据岗-1.20\candidate.pdf")

    output_path, debug_dir = module.resolve_output_locations(
        output_arg=None,
        folder_arg=None,
        file_arg=str(file_path),
        debug_arg="debug",
        save_extracted_text=None,
    )

    assert output_path == str(file_path.parent / "resume_risk_report.xlsx")
    assert debug_dir == str(file_path.parent / "debug")
