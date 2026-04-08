from __future__ import annotations

import uuid
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from starlette.concurrency import run_in_threadpool

from app import (
    CLASS_ID,
    COURSE_ID,
    DEFAULT_MODEL,
    DEFAULT_SOURCE_ID,
    DEFAULT_MEMBER4_INGEST_URL,
    EVENT_TYPE,
    DEFAULT_OPENAI_BASE_URL,
    QUESTION_EXT,
    SOURCE_TYPE,
    STUDENT_WORD_EXT,
    OUTPUT_DIR,
    grade_homework,
    normalize_student_id,
    post_member4_event,
    save_upload,
)

api = FastAPI(
    title="Homework Grading API",
    version="1.0.0",
    description="题目/材料多格式输入，学生Word作业自动批注并输出Word文件",
)

JOBS: dict[str, dict[str, Any]] = {}


def _validate_upload(name: str, ext_set: set[str], field_name: str) -> None:
    ext = Path(name).suffix.lower()
    if ext not in ext_set:
        allowed = ", ".join(sorted(ext_set))
        raise HTTPException(status_code=400, detail=f"{field_name}格式不支持: {ext}，允许: {allowed}")


@api.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok", "time": datetime.now().isoformat()}


@api.post("/api/v1/grade")
async def grade(
    question_file: UploadFile = File(..., description="题目文件: pdf/doc/docx/xls/xlsx"),
    student_file: UploadFile = File(..., description="学生作业: doc/docx"),
    student_id: str = Form(..., description="学生ID(必填，支持字母/数字/_/-)"),
    teacher_material_files: list[UploadFile] | None = File(default=None, description="老师补充材料(可多文件)"),
    reference_file: UploadFile | None = File(default=None, description="老师批改样例(可选)"),
    protocol: str = Form(default="OpenAI兼容"),
    api_key: str = Form(...),
    base_url: str = Form(default=DEFAULT_OPENAI_BASE_URL),
    model: str = Form(default=DEFAULT_MODEL),
) -> JSONResponse:
    if protocol not in {"OpenAI兼容", "Anthropic兼容"}:
        raise HTTPException(status_code=400, detail="protocol 仅支持 OpenAI兼容 或 Anthropic兼容")
    try:
        normalized_student_id = normalize_student_id(student_id)
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc))

    _validate_upload(question_file.filename or "", QUESTION_EXT, "question_file")
    _validate_upload(student_file.filename or "", STUDENT_WORD_EXT, "student_file")

    if reference_file and reference_file.filename:
        _validate_upload(reference_file.filename, QUESTION_EXT, "reference_file")

    material_files = teacher_material_files or []
    for f in material_files:
        _validate_upload(f.filename or "", QUESTION_EXT, "teacher_material_files")

    job_id = uuid.uuid4().hex
    work_dir = Path("workspace_uploads") / f"api_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{job_id[:8]}"

    question_path = save_upload(question_file, work_dir)
    student_path = save_upload(student_file, work_dir)
    reference_path = save_upload(reference_file, work_dir) if reference_file else None
    teacher_paths = [save_upload(f, work_dir) for f in material_files]

    try:
        output_path, overall, kp_event = await run_in_threadpool(
            grade_homework,
            question_path,
            student_path,
            normalized_student_id,
            reference_path,
            teacher_paths,
            protocol,
            api_key,
            base_url,
            model,
            OUTPUT_DIR,
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"批改失败: {exc}")

    kp_results = kp_event.get("payload", {}).get("results", [])
    ingest_url = os.getenv("MEMBER4_INGEST_URL", DEFAULT_MEMBER4_INGEST_URL)
    try:
        push_result = await run_in_threadpool(
            post_member4_event,
            ingest_url,
            normalized_student_id,
            job_id,
            kp_results,
        )
    except Exception as exc:
        push_result = {
            "request": {
                "events": [
                    {
                        "user_id": normalized_student_id,
                        "class_id": CLASS_ID,
                        "course_id": COURSE_ID,
                        "event_type": EVENT_TYPE,
                        "payload": {"source_id": job_id, "source_type": SOURCE_TYPE, "results": kp_results},
                    }
                ]
            },
            "error": str(exc),
        }

    JOBS[job_id] = {
        "job_id": job_id,
        "created_at": datetime.now().isoformat(),
        "overall": overall,
        "student_id": normalized_student_id,
        "output_file": str(output_path.resolve()),
        "output_name": output_path.name,
        "structured_result": {
            "events": [
                {
                    "user_id": normalized_student_id,
                    "class_id": CLASS_ID,
                    "course_id": COURSE_ID,
                    "event_type": EVENT_TYPE,
                    "payload": {
                        "source_id": job_id or DEFAULT_SOURCE_ID,
                        "source_type": SOURCE_TYPE,
                        "results": kp_results,
                    },
                }
            ]
        },
        "member4_push": push_result,
    }

    return JSONResponse(
        {
            "job_id": job_id,
            "overall": overall,
            "student_id": normalized_student_id,
            "output_file_name": output_path.name,
            "structured_result": JOBS[job_id]["structured_result"],
            "member4_push": push_result,
            "download_url": f"/api/v1/download/{job_id}",
        }
    )


@api.get("/api/v1/result/{job_id}")
def result(job_id: str) -> dict[str, Any]:
    info = JOBS.get(job_id)
    if not info:
        raise HTTPException(status_code=404, detail="job_id 不存在")
    return info


@api.get("/api/v1/download/{job_id}")
def download(job_id: str) -> FileResponse:
    info = JOBS.get(job_id)
    if not info:
        raise HTTPException(status_code=404, detail="job_id 不存在")

    file_path = Path(info["output_file"])
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="输出文件不存在")

    return FileResponse(path=str(file_path), filename=info["output_name"], media_type="application/octet-stream")
