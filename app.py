import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx
import pythoncom
import streamlit as st
import win32com.client as win32
from dotenv import load_dotenv
from openai import OpenAI
from pypdf import PdfReader

load_dotenv()

DEFAULT_OPENAI_BASE_URL = "https://ark.cn-beijing.volces.com/api/coding/v3"
DEFAULT_ANTHROPIC_BASE_URL = "https://ark.cn-beijing.volces.com/api/coding"
DEFAULT_MODEL = "doubao-seed-2.0-pro"
QUESTION_EXT = {".doc", ".docx", ".pdf", ".xls", ".xlsx"}
STUDENT_WORD_EXT = {".doc", ".docx"}
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)


class ComScope:
    def __enter__(self):
        pythoncom.CoInitialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pythoncom.CoUninitialize()
        return False


def now_suffix() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def normalize_student_id(student_id: str) -> str:
    value = student_id.strip()
    if not value:
        raise ValueError("学生ID不能为空。")
    normalized = re.sub(r"[^0-9A-Za-z_-]+", "_", value).strip("_")
    if not normalized:
        raise ValueError("学生ID仅支持字母、数字、下划线和中划线。")
    return normalized[:64]


def normalize_text(text: str) -> str:
    cleaned = text.replace("\r", "\n")
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def extract_json(text: str) -> dict[str, Any]:
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{[\s\S]*\}", text)
    if not match:
        return {}

    try:
        return json.loads(match.group(0))
    except json.JSONDecodeError:
        return {}


def call_openai_compatible(api_key: str, base_url: str, model: str, system_prompt: str, user_prompt: str) -> str:
    client = OpenAI(api_key=api_key, base_url=base_url)
    response = client.chat.completions.create(
        model=model,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.2,
        max_tokens=1800,
    )
    return (response.choices[0].message.content or "").strip()


def call_anthropic_compatible(api_key: str, base_url: str, model: str, system_prompt: str, user_prompt: str) -> str:
    url = base_url.rstrip("/")
    if not url.endswith("/v1/messages"):
        url = f"{url}/v1/messages"

    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
    }
    payload = {
        "model": model,
        "max_tokens": 1800,
        "temperature": 0.2,
        "system": system_prompt,
        "messages": [{"role": "user", "content": [{"type": "text", "text": user_prompt}]}],
    }

    with httpx.Client(timeout=120.0) as client:
        response = client.post(url, headers=headers, json=payload)
        response.raise_for_status()
        data = response.json()

    parts = data.get("content", [])
    texts = [part.get("text", "") for part in parts if part.get("type") == "text"]
    return "\n".join([t.strip() for t in texts if t.strip()]).strip()


def call_model(protocol: str, api_key: str, base_url: str, model: str, system_prompt: str, user_prompt: str) -> str:
    if protocol == "Anthropic兼容":
        return call_anthropic_compatible(api_key, base_url, model, system_prompt, user_prompt)
    return call_openai_compatible(api_key, base_url, model, system_prompt, user_prompt)


def extract_text_from_pdf(file_path: Path) -> str:
    reader = PdfReader(str(file_path))
    text = []
    for page in reader.pages:
        text.append(page.extract_text() or "")
    return normalize_text("\n".join(text))


def extract_text_from_word(file_path: Path) -> str:
    with ComScope():
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = None
        try:
            doc = word.Documents.Open(str(file_path.resolve()), ReadOnly=True)
            lines = []
            for i in range(1, doc.Paragraphs.Count + 1):
                raw = doc.Paragraphs(i).Range.Text
                txt = raw.replace("\r", "").replace("\x07", "").strip()
                if txt:
                    lines.append(txt)
            return normalize_text("\n".join(lines))
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()


def extract_text_from_excel(file_path: Path) -> str:
    with ComScope():
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = None
        try:
            wb = excel.Workbooks.Open(str(file_path.resolve()), ReadOnly=True)
            lines = []
            for ws in wb.Worksheets:
                used = ws.UsedRange
                row_count = used.Rows.Count
                col_count = used.Columns.Count
                start_row = used.Row
                start_col = used.Column
                for r in range(start_row, start_row + row_count):
                    row_values = []
                    for c in range(start_col, start_col + col_count):
                        val = ws.Cells(r, c).Value
                        if val is not None and str(val).strip():
                            row_values.append(str(val).strip())
                    if row_values:
                        lines.append(f"[{ws.Name}] " + " | ".join(row_values))
            return normalize_text("\n".join(lines))
        finally:
            if wb is not None:
                wb.Close(False)
            excel.Quit()


def extract_text(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_path)
    if ext in {".doc", ".docx"}:
        return extract_text_from_word(file_path)
    if ext in {".xls", ".xlsx"}:
        return extract_text_from_excel(file_path)
    raise ValueError(f"不支持的文件类型: {ext}")


def generate_annotation_plan(
    protocol: str,
    api_key: str,
    base_url: str,
    model: str,
    question_text: str,
    student_segments: list[dict[str, Any]],
    segment_hint: str,
    reference_text: str = "",
) -> dict[str, Any]:
    system_prompt = (
        "你是严谨且鼓励式的老师。你只返回JSON，不要返回Markdown或解释。"
        "批注要简短，聚焦知识点与改进建议。"
    )

    limited_segments = student_segments[:50]
    serialized = []
    for item in limited_segments:
        serialized.append(f"index={item['index']} | content={item['text']}")

    user_prompt = f"""
请依据题目与学生作答给出批注计划。

输出JSON格式必须是：
{{
  "overall": "总评，不超过80字",
  "items": [
    {{"index": 1, "comment": "该段批注，不超过40字"}}
  ]
}}

要求：
1) items 只包含需要批注的条目，最多10条。
2) index 必须来自我提供的 index。
3) comment 使用中文。
4) 不要输出任何JSON以外文本。

题目：
{question_text[:6000]}

参考老师批改样例（可选）：
{reference_text[:3000] if reference_text else '无'}

学生作答条目（{segment_hint}）：
{chr(10).join(serialized)}
""".strip()

    raw = call_model(protocol, api_key, base_url, model, system_prompt, user_prompt)
    parsed = extract_json(raw)
    if not parsed:
        return {"overall": "整体完成较认真，建议根据题目要求进一步完善关键步骤。", "items": []}

    overall = str(parsed.get("overall", "")).strip() or "整体完成较认真，建议根据题目要求进一步完善关键步骤。"
    items = parsed.get("items", [])
    normalized_items = []
    for item in items:
        try:
            idx = int(item.get("index"))
            comment = str(item.get("comment", "")).strip()
            if idx > 0 and comment:
                normalized_items.append({"index": idx, "comment": comment[:120]})
        except Exception:
            continue

    return {"overall": overall[:120], "items": normalized_items}


def annotate_word(
    student_path: Path,
    student_id: str,
    output_dir: Path,
    protocol: str,
    api_key: str,
    base_url: str,
    model: str,
    question_text: str,
    reference_text: str,
) -> tuple[Path, str]:
    with ComScope():
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = None
        try:
            doc = word.Documents.Open(str(student_path.resolve()))
            segments = []
            paragraph_map: dict[int, Any] = {}
            idx = 1
            for i in range(1, doc.Paragraphs.Count + 1):
                para = doc.Paragraphs(i)
                txt = para.Range.Text.replace("\r", "").replace("\x07", "").strip()
                if txt:
                    segments.append({"index": idx, "text": txt})
                    paragraph_map[idx] = para
                    idx += 1

            plan = generate_annotation_plan(
                protocol, api_key, base_url, model, question_text, segments, "Word段落", reference_text
            )

            for item in plan["items"]:
                para = paragraph_map.get(item["index"])
                if para is None:
                    continue
                doc.Comments.Add(Range=para.Range, Text=item["comment"])

            if doc.Paragraphs.Count >= 1:
                first_para = doc.Paragraphs(1).Range
                doc.Comments.Add(Range=first_para, Text=f"总评：{plan['overall']}")

            output_path = output_dir / (
                f"{student_path.stem}-学生ID-{student_id}-批改后-{now_suffix()}{student_path.suffix}"
            )
            doc.SaveAs2(str(output_path.resolve()))
            return output_path, plan["overall"]
        finally:
            if doc is not None:
                doc.Close(False)
            word.Quit()


def grade_homework(
    question_path: Path,
    student_path: Path,
    student_id: str,
    reference_path: Path | None,
    teacher_material_paths: list[Path] | None,
    protocol: str,
    api_key: str,
    base_url: str,
    model: str,
    output_dir: Path,
) -> tuple[Path, str]:
    ext = student_path.suffix.lower()
    if ext not in STUDENT_WORD_EXT:
        raise ValueError(f"待批改文件必须是Word格式(doc/docx)，当前为：{ext}")
    normalized_student_id = normalize_student_id(student_id)

    teacher_paths = [question_path] + (teacher_material_paths or [])
    teacher_text_parts = []
    for p in teacher_paths:
        teacher_text_parts.append(f"[教师材料: {p.name}]")
        teacher_text_parts.append(extract_text(p))
    question_text = normalize_text("\n\n".join(teacher_text_parts))
    reference_text = extract_text(reference_path) if reference_path else ""
    return annotate_word(
        student_path,
        normalized_student_id,
        output_dir,
        protocol,
        api_key,
        base_url,
        model,
        question_text,
        reference_text,
    )


def save_upload(uploaded_file, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    file_name = getattr(uploaded_file, "name", None) or getattr(uploaded_file, "filename", None)
    if not file_name:
        raise ValueError("上传文件缺少文件名。")
    target = target_dir / Path(file_name).name

    if hasattr(uploaded_file, "getbuffer"):
        target.write_bytes(uploaded_file.getbuffer())
        return target

    file_obj = getattr(uploaded_file, "file", None)
    if file_obj is not None and hasattr(file_obj, "read"):
        file_obj.seek(0)
        target.write_bytes(file_obj.read())
        return target

    raise ValueError("不支持的上传文件对象类型。")


def render_upload_summary(
    question_upload,
    teacher_material_uploads,
    student_upload,
    reference_upload,
    student_id: str,
) -> None:
    st.markdown("### 已上传文件")
    st.write(f"- 学生ID：`{student_id.strip() or '未填写'}`")
    if question_upload:
        st.write(f"- 题目：`{question_upload.name}`")
    else:
        st.write("- 题目：`未上传`")

    if teacher_material_uploads:
        for item in teacher_material_uploads:
            st.write(f"- 材料：`{item.name}`")
    else:
        st.write("- 材料：`无`")

    if student_upload:
        st.write(f"- 作业：`{student_upload.name}`")
    else:
        st.write("- 作业：`未上传`")

    if reference_upload:
        st.write(f"- 样例：`{reference_upload.name}`")
    else:
        st.write("- 样例：`无`")


def main() -> None:
    st.set_page_config(page_title="作业批注系统", page_icon="📝", layout="centered")
    st.title("作业批注前端页")
    st.caption("上传题目/材料/学生作业，系统处理后输出批注完成的 Word 文件")

    with st.sidebar:
        st.markdown("### 模型设置")
        protocol = st.selectbox("接口协议", ["OpenAI兼容", "Anthropic兼容"], index=0)
        api_key = st.text_input("ARK API Key", value=os.getenv("ARK_API_KEY", ""), type="password")
        model = st.text_input("模型名", value=os.getenv("ARK_MODEL", DEFAULT_MODEL))

        default_base_url = os.getenv("ARK_BASE_URL", DEFAULT_OPENAI_BASE_URL)
        if protocol == "Anthropic兼容":
            default_base_url = DEFAULT_ANTHROPIC_BASE_URL
        base_url = st.text_input("Base URL", value=default_base_url)

        st.markdown("### 支持格式")
        st.write("- 题目/材料：`pdf/doc/docx/xls/xlsx`")
        st.write("- 学生作业：`doc/docx`")
        st.write("- 输出：批注后的 `doc/docx`")

    with st.form("grading_form"):
        st.markdown("### 0) 填写学生ID")
        student_id = st.text_input("学生ID（必填，仅用于标识和输出文件命名）")

        st.markdown("### 1) 上传题目")
        question_upload = st.file_uploader(
            "题目文件（必填）",
            type=["pdf", "doc", "docx", "xls", "xlsx"],
        )

        st.markdown("### 2) 上传老师材料（可选）")
        teacher_material_uploads = st.file_uploader(
            "补充材料（可多选）",
            type=["pdf", "doc", "docx", "xls", "xlsx"],
            accept_multiple_files=True,
        )

        st.markdown("### 3) 上传学生作业")
        student_upload = st.file_uploader("学生待批改 Word（必填）", type=["doc", "docx"])

        st.markdown("### 4) 上传参考样例（可选）")
        reference_upload = st.file_uploader(
            "老师批改样例",
            type=["pdf", "doc", "docx", "xls", "xlsx"],
        )

        submitted = st.form_submit_button("开始处理并生成批注Word", type="primary")

    render_upload_summary(question_upload, teacher_material_uploads, student_upload, reference_upload, student_id)

    if submitted:
        if not question_upload or not student_upload:
            st.error("请至少上传题目文件和学生待批改文件。")
            return
        if not student_id.strip():
            st.error("请填写学生ID。")
            return
        if not api_key.strip():
            st.error("请填写 API Key。")
            return

        question_ext = Path(question_upload.name).suffix.lower()
        student_ext = Path(student_upload.name).suffix.lower()
        if question_ext not in QUESTION_EXT:
            st.error("题目文件格式不支持，请使用 pdf/doc/docx/xls/xlsx。")
            return
        if student_ext not in STUDENT_WORD_EXT:
            st.error("学生待批改文件必须是 doc/docx。")
            return

        work_dir = Path("workspace_uploads") / now_suffix()

        with st.spinner("正在批改，请稍候..."):
            try:
                normalized_student_id = normalize_student_id(student_id)
                question_path = save_upload(question_upload, work_dir)
                teacher_material_paths = [save_upload(f, work_dir) for f in (teacher_material_uploads or [])]
                student_path = save_upload(student_upload, work_dir)
                reference_path = save_upload(reference_upload, work_dir) if reference_upload else None

                output_path, overall = grade_homework(
                    question_path,
                    student_path,
                    normalized_student_id,
                    reference_path,
                    teacher_material_paths,
                    protocol,
                    api_key,
                    base_url,
                    model,
                    OUTPUT_DIR,
                )

                st.success("批改完成")
                st.write(f"总评：{overall}")
                st.code(f"输出文件：{output_path.resolve()}")
                st.download_button(
                    "下载批改后文件",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime="application/octet-stream",
                )
            except Exception as exc:
                st.error(f"批改失败：{exc}")


if __name__ == "__main__":
    main()
