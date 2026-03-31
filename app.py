import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any

import httpx
import streamlit as st
from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import load_workbook
from pypdf import PdfReader

load_dotenv()

DEFAULT_OPENAI_BASE_URL = "https://ark.cn-beijing.volces.com/api/coding/v3"
DEFAULT_ANTHROPIC_BASE_URL = "https://ark.cn-beijing.volces.com/api/coding"
DEFAULT_MODEL = "doubao-seed-2.0-pro"
QUESTION_EXT = {".doc", ".docx", ".pdf", ".xls", ".xlsx"}
STUDENT_WORD_EXT = {".doc", ".docx"}
OUTPUT_DIR = Path("outputs")
OUTPUT_DIR.mkdir(exist_ok=True)


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


def convert_legacy_office_file(file_path: Path, target_ext: str) -> Path:
    soffice = shutil.which("soffice")
    if not soffice:
        raise ValueError(
            f"文件 `{file_path.name}` 为旧格式 `{file_path.suffix}`，当前环境无法直接解析。"
            f"请改传 `{target_ext}`，或在服务器安装 LibreOffice（soffice）后自动转换。"
        )

    temp_dir = Path(tempfile.mkdtemp(prefix="office_convert_"))
    convert_to = target_ext.lstrip(".")
    result = subprocess.run(
        [soffice, "--headless", "--convert-to", convert_to, "--outdir", str(temp_dir), str(file_path)],
        capture_output=True,
        text=True,
        check=False,
    )
    converted = temp_dir / f"{file_path.stem}.{convert_to}"
    if result.returncode != 0 or not converted.exists():
        err = (result.stderr or result.stdout or "").strip()
        raise ValueError(f"无法将 `{file_path.name}` 转换为 `{target_ext}`。{err}")
    return converted


def ensure_docx(file_path: Path) -> Path:
    ext = file_path.suffix.lower()
    if ext == ".docx":
        return file_path
    if ext == ".doc":
        return convert_legacy_office_file(file_path, ".docx")
    raise ValueError(f"不支持的Word文件类型: {ext}")


def ensure_xlsx(file_path: Path) -> Path:
    ext = file_path.suffix.lower()
    if ext == ".xlsx":
        return file_path
    if ext == ".xls":
        return convert_legacy_office_file(file_path, ".xlsx")
    raise ValueError(f"不支持的Excel文件类型: {ext}")


def extract_text_from_pdf(file_path: Path) -> str:
    reader = PdfReader(str(file_path))
    text = []
    for page in reader.pages:
        text.append(page.extract_text() or "")
    return normalize_text("\n".join(text))


def extract_text_from_word(file_path: Path) -> str:
    docx_path = ensure_docx(file_path)
    doc = Document(str(docx_path))
    lines = []
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            lines.append(txt)
    return normalize_text("\n".join(lines))


def extract_text_from_excel(file_path: Path) -> str:
    xlsx_path = ensure_xlsx(file_path)
    wb = load_workbook(str(xlsx_path), data_only=True, read_only=True)
    try:
        lines = []
        for ws in wb.worksheets:
            for row in ws.iter_rows(values_only=True):
                row_values = []
                for val in row:
                    if val is not None and str(val).strip():
                        row_values.append(str(val).strip())
                if row_values:
                    lines.append(f"[{ws.title}] " + " | ".join(row_values))
        return normalize_text("\n".join(lines))
    finally:
        wb.close()


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


def generate_learning_analysis(
    protocol: str,
    api_key: str,
    base_url: str,
    model: str,
    student_id: str,
    question_text: str,
    student_text: str,
    reference_text: str = "",
) -> dict[str, Any]:
    system_prompt = (
        "你是教学评估专家。你只返回JSON，不要返回Markdown或解释。"
        "结论应客观、可落地、可用于教学跟进。"
    )
    user_prompt = f"""
请基于题目要求和学生作答，输出该学生的学情分析。

输出JSON格式必须是：
{{
  "student_id": "{student_id}",
  "overall_mastery": "总体掌握结论，不超过80字",
  "requirements": [
    {{
      "point": "题目要求点",
      "status": "已完成/部分完成/未完成",
      "evidence": "作答证据，不超过80字",
      "advice": "改进建议，不超过60字"
    }}
  ],
  "knowledge_mastery": [
    {{
      "topic": "知识点",
      "level": "熟练/基本掌握/待加强",
      "analysis": "分析，不超过80字"
    }}
  ]
}}

要求：
1) requirements 至少3条，最多8条，尽量覆盖题目关键要求。
2) status 仅能使用：已完成、部分完成、未完成。
3) knowledge_mastery 至少3条，最多8条。
4) 不要输出任何JSON以外文本。

题目与教师材料：
{question_text[:8000]}

参考样例（可选）：
{reference_text[:3000] if reference_text else '无'}

学生作答：
{student_text[:8000]}
""".strip()

    raw = call_model(protocol, api_key, base_url, model, system_prompt, user_prompt)
    parsed = extract_json(raw)
    if not parsed:
        return {
            "student_id": student_id,
            "overall_mastery": "当前可完成基础要求，综合应用与细节准确性仍需提升。",
            "requirements": [],
            "knowledge_mastery": [],
        }

    overall = str(parsed.get("overall_mastery", "")).strip() or "当前可完成基础要求，综合应用与细节准确性仍需提升。"

    requirements = []
    for item in parsed.get("requirements", []):
        point = str(item.get("point", "")).strip()
        status = str(item.get("status", "")).strip()
        evidence = str(item.get("evidence", "")).strip()
        advice = str(item.get("advice", "")).strip()
        if not point:
            continue
        if status not in {"已完成", "部分完成", "未完成"}:
            status = "部分完成"
        requirements.append(
            {
                "point": point[:120],
                "status": status,
                "evidence": evidence[:160],
                "advice": advice[:120],
            }
        )

    knowledge_mastery = []
    for item in parsed.get("knowledge_mastery", []):
        topic = str(item.get("topic", "")).strip()
        level = str(item.get("level", "")).strip()
        analysis = str(item.get("analysis", "")).strip()
        if not topic:
            continue
        if level not in {"熟练", "基本掌握", "待加强"}:
            level = "基本掌握"
        knowledge_mastery.append(
            {
                "topic": topic[:120],
                "level": level,
                "analysis": analysis[:160],
            }
        )

    return {
        "student_id": student_id,
        "overall_mastery": overall[:160],
        "requirements": requirements[:8],
        "knowledge_mastery": knowledge_mastery[:8],
    }


def save_learning_analysis_report(student_id: str, student_path: Path, analysis: dict[str, Any], output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / f"{student_path.stem}-学生ID-{student_id}-学情分析-{now_suffix()}.docx"

    doc = Document()
    doc.add_heading("学情分析报告", level=1)
    doc.add_paragraph(f"学生ID：{analysis.get('student_id') or student_id}")
    doc.add_paragraph(f"综合掌握结论：{analysis.get('overall_mastery', '')}")

    doc.add_heading("题目要求完成情况", level=2)
    requirements = analysis.get("requirements", [])
    if requirements:
        for idx, item in enumerate(requirements, start=1):
            doc.add_paragraph(f"{idx}. 要求点：{item.get('point', '')}")
            doc.add_paragraph(f"完成状态：{item.get('status', '')}")
            doc.add_paragraph(f"作答证据：{item.get('evidence', '') or '未提取到明确证据'}")
            doc.add_paragraph(f"改进建议：{item.get('advice', '') or '建议补充关键步骤与依据'}")
    else:
        doc.add_paragraph("1. 暂未提取到结构化要求点，请结合批注报告复核。")

    doc.add_heading("知识掌握程度分析", level=2)
    knowledges = analysis.get("knowledge_mastery", [])
    if knowledges:
        for idx, item in enumerate(knowledges, start=1):
            doc.add_paragraph(f"{idx}. 知识点：{item.get('topic', '')}")
            doc.add_paragraph(f"掌握程度：{item.get('level', '')}")
            doc.add_paragraph(f"分析：{item.get('analysis', '') or '建议在练习中强化迁移应用。'}")
    else:
        doc.add_paragraph("1. 暂未提取到结构化知识点，请结合批注报告复核。")

    doc.save(str(report_path.resolve()))
    return report_path


def get_or_add_comments_part(document: Document) -> XmlPart:
    doc_part = document.part
    for rel in doc_part.rels.values():
        if rel.reltype == RT.COMMENTS:
            return rel.target_part

    comments_xml = parse_xml(f"<w:comments {nsdecls('w')}/>")
    comments_part = XmlPart(PackURI("/word/comments.xml"), CT.WML_COMMENTS, comments_xml, doc_part.package)
    doc_part.relate_to(comments_part, RT.COMMENTS)
    return comments_part


def next_comment_id(comments_part: XmlPart) -> int:
    comments_root = comments_part.element
    max_id = -1
    for node in comments_root.findall(qn("w:comment")):
        raw = node.get(qn("w:id"))
        if raw is None:
            continue
        try:
            max_id = max(max_id, int(raw))
        except ValueError:
            continue
    return max_id + 1


def add_comment_to_paragraph(document: Document, paragraph: Any, comment_text: str, author: str = "AI批改助手") -> None:
    text = comment_text.strip()
    if not text:
        return

    comments_part = get_or_add_comments_part(document)
    cid = next_comment_id(comments_part)

    comment = OxmlElement("w:comment")
    comment.set(qn("w:id"), str(cid))
    comment.set(qn("w:author"), author)
    comment.set(qn("w:initials"), "AI")
    comment.set(qn("w:date"), datetime.utcnow().replace(microsecond=0).isoformat() + "Z")

    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    r.append(t)
    p.append(r)
    comment.append(p)
    comments_part.element.append(comment)

    p_elm = paragraph._p
    runs = [child for child in p_elm.iterchildren() if child.tag == qn("w:r")]
    if not runs:
        paragraph.add_run(" ")
        runs = [child for child in p_elm.iterchildren() if child.tag == qn("w:r")]

    if not runs:
        return

    first_run = runs[0]
    last_run = runs[-1]

    range_start = OxmlElement("w:commentRangeStart")
    range_start.set(qn("w:id"), str(cid))
    first_run.addprevious(range_start)

    range_end = OxmlElement("w:commentRangeEnd")
    range_end.set(qn("w:id"), str(cid))
    last_run.addnext(range_end)

    ref_run = OxmlElement("w:r")
    rpr = OxmlElement("w:rPr")
    rstyle = OxmlElement("w:rStyle")
    rstyle.set(qn("w:val"), "CommentReference")
    rpr.append(rstyle)
    ref_run.append(rpr)
    cref = OxmlElement("w:commentReference")
    cref.set(qn("w:id"), str(cid))
    ref_run.append(cref)
    range_end.addnext(ref_run)


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
) -> tuple[Path, str, str]:
    docx_path = ensure_docx(student_path)
    doc = Document(str(docx_path))

    segments = []
    segment_map: dict[int, tuple[Any, str]] = {}
    idx = 1
    for para in doc.paragraphs:
        txt = para.text.strip()
        if txt:
            segments.append({"index": idx, "text": txt})
            segment_map[idx] = (para, txt)
            idx += 1

    plan = generate_annotation_plan(protocol, api_key, base_url, model, question_text, segments, "Word段落", reference_text)

    if doc.paragraphs:
        add_comment_to_paragraph(doc, doc.paragraphs[0], f"总评：{plan['overall']}")

    for item in plan["items"]:
        mapped = segment_map.get(item["index"])
        if not mapped:
            continue
        para, _ = mapped
        add_comment_to_paragraph(doc, para, item["comment"])

    output_path = output_dir / f"{student_path.stem}-学生ID-{student_id}-批改后-{now_suffix()}.docx"
    output_dir.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path.resolve()))
    student_text = normalize_text("\n".join([item["text"] for item in segments]))
    return output_path, plan["overall"], student_text


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
) -> tuple[Path, str, Path]:
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
    output_path, overall, student_text = annotate_word(
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
    analysis = generate_learning_analysis(
        protocol=protocol,
        api_key=api_key,
        base_url=base_url,
        model=model,
        student_id=normalized_student_id,
        question_text=question_text,
        student_text=student_text,
        reference_text=reference_text,
    )
    analysis_path = save_learning_analysis_report(normalized_student_id, student_path, analysis, output_dir)
    return output_path, overall, analysis_path


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
        st.write("- 输出：批注后的 `docx`")
        st.caption("提示：Linux 环境处理 .doc/.xls 需要安装 LibreOffice（soffice）自动转换。")

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

                output_path, overall, analysis_path = grade_homework(
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
                st.code(f"学情分析：{analysis_path.resolve()}")
                st.download_button(
                    "下载批改后文件",
                    data=output_path.read_bytes(),
                    file_name=output_path.name,
                    mime="application/octet-stream",
                )
                st.download_button(
                    "下载学情分析",
                    data=analysis_path.read_bytes(),
                    file_name=analysis_path.name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as exc:
                st.error(f"批改失败：{exc}")


if __name__ == "__main__":
    main()
