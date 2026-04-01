import os
from pathlib import Path

from dotenv import load_dotenv

from app import DEFAULT_MODEL, DEFAULT_OPENAI_BASE_URL, grade_homework


def require_file(path: Path) -> Path:
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"示例文件不存在: {path}")
    return path


def main() -> None:
    load_dotenv()

    api_key = os.getenv("ARK_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("未找到 ARK_API_KEY")

    protocol = "OpenAI兼容"
    base_url = os.getenv("ARK_BASE_URL", DEFAULT_OPENAI_BASE_URL)
    model = os.getenv("ARK_MODEL", DEFAULT_MODEL)
    use_model_file_inputs = os.getenv("USE_MODEL_FILE_INPUTS", "false").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }

    root = Path(__file__).resolve().parent
    question_path = require_file(root / "example" / "Excel实验题目要求.pdf")
    student_path = require_file(root / "example" / "Excel实验报告-待批改.docx")
    reference_path = require_file(root / "example" / "Excel实验报告-批改后.docx")
    excel_material_path = require_file(root / "example" / "Excel实验原始素材文件.xlsx")

    output_path, overall, analysis_path = grade_homework(
        question_path=question_path,
        student_path=student_path,
        student_id="demo_001",
        reference_path=reference_path,
        teacher_material_paths=[excel_material_path],
        protocol=protocol,
        api_key=api_key,
        base_url=base_url,
        model=model,
        output_dir=root / "outputs",
        use_model_file_inputs=use_model_file_inputs,
    )

    print(str(output_path.resolve()))
    print(str(analysis_path.resolve()))
    print(overall)


if __name__ == "__main__":
    main()
