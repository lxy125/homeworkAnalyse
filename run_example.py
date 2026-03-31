import os
from pathlib import Path

from dotenv import load_dotenv

from app import DEFAULT_MODEL, DEFAULT_OPENAI_BASE_URL, grade_homework


def main() -> None:
    load_dotenv()

    api_key = os.getenv("ARK_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError("未找到 ARK_API_KEY")

    protocol = "OpenAI兼容"
    base_url = os.getenv("ARK_BASE_URL", DEFAULT_OPENAI_BASE_URL)
    model = os.getenv("ARK_MODEL", DEFAULT_MODEL)

    root = Path(__file__).resolve().parent
    question_path = root / "example" / "Excel实验题目要求.pdf"
    student_path = root / "example" / "Excel实验报告-待批改.docx"
    reference_path = root / "example" / "Excel实验报告-批改后.docx"

    output_path, overall, analysis_path = grade_homework(
        question_path=question_path,
        student_path=student_path,
        student_id="demo_001",
        reference_path=reference_path,
        teacher_material_paths=[],
        protocol=protocol,
        api_key=api_key,
        base_url=base_url,
        model=model,
        output_dir=root / "outputs",
    )

    print(str(output_path.resolve()))
    print(str(analysis_path.resolve()))
    print(overall)


if __name__ == "__main__":
    main()
