# API 接口文档（成员2 / 成员4 对接版）

后端入口文件：`backend_api.py`

## 1. 启动与地址

启动（不运行 Web 页面）：

```powershell
conda activate homework_annotator
python run_backend.py
```

默认地址：`http://127.0.0.1:8000`

在线文档：

- Swagger UI: `http://127.0.0.1:8000/docs`
- ReDoc: `http://127.0.0.1:8000/redoc`

成员4推送地址（可配置环境变量）：

- `MEMBER4_INGEST_URL`，默认：`http://127.0.0.1:8007/api/v1/analytics/events/ingest`

## 2. 固定业务规则（非常重要）

成员3在推送给成员4时，固定字段如下：

- `class_id = "class-demo"`
- `course_id = "course_oop_design"`
- `source_type = "assignment"`
- `event_type = "GRADING_RESULT_RECEIVED"`
- `user_id = student_id`（成员2传入的学生ID会映射为成员4事件中的 user_id）
- `source_id = job_id`（后端API模式下，默认使用本次批改任务ID）

知识点范围（模型只能从以下ID中选择）：

- `kp_object_oriented_basics`
- `kp_class_and_object`
- `kp_encapsulation_and_access_control`
- `kp_inheritance`
- `kp_polymorphism`
- `kp_abstract_class_and_interface`
- `kp_exception_handling`
- `kp_common_collections_and_generics`

---

## 3. 成员2如何接入（上传题目/作业并触发批改）

### 3.1 健康检查

- 方法：`GET`
- 路径：`/health`

返回示例：

```json
{
  "status": "ok",
  "time": "2026-04-08T10:00:00.000000"
}
```

### 3.2 提交批改任务

- 方法：`POST`
- 路径：`/api/v1/grade`
- Content-Type：`multipart/form-data`

表单字段：

- `question_file` (file, 必填): 题目文件，支持 `pdf/doc/docx/xls/xlsx`
- `student_file` (file, 必填): 学生作业，支持 `doc/docx`（输出固定为 `docx`）
- `student_id` (text, 必填): 学生ID（仅支持字母/数字/_/-）
- `teacher_material_files` (file[], 可选): 老师补充材料，多文件
- `reference_file` (file, 可选): 老师批改样例
- `protocol` (text, 可选): `OpenAI兼容` 或 `Anthropic兼容`，默认 `OpenAI兼容`
- `api_key` (text, 必填): ARK API Key
- `base_url` (text, 可选): 默认 `https://ark.cn-beijing.volces.com/api/coding/v3`
- `model` (text, 可选): 默认 `doubao-seed-2.0-pro`

返回示例：

```json
{
  "job_id": "b9d1f8b4f0f54f25ae9d1b4f88123456",
  "student_id": "stu-001",
  "overall": "整体完成较好，建议完善细节。",
  "output_file_name": "对象与类作业-待批改-学生ID-stu-001-批改后-20260408_100000.docx",
  "structured_result": {
    "events": [
      {
        "user_id": "stu-001",
        "class_id": "class-demo",
        "course_id": "course_oop_design",
        "event_type": "GRADING_RESULT_RECEIVED",
        "payload": {
          "source_id": "b9d1f8b4f0f54f25ae9d1b4f88123456",
          "source_type": "assignment",
          "results": [
            {
              "knowledge_point_id": "kp_class_and_object",
              "is_correct": true
            },
            {
              "knowledge_point_id": "kp_polymorphism",
              "is_correct": false
            }
          ]
        }
      }
    ]
  },
  "member4_push": {
    "status_code": 200
  },
  "download_url": "/api/v1/download/b9d1f8b4f0f54f25ae9d1b4f88123456"
}
```

说明：

- 批改Word生成成功后，会自动推送给成员4。
- 即使成员4暂时不可用，批改主流程仍可成功；此时 `member4_push` 中会有错误信息。

### 3.3 查询任务详情

- 方法：`GET`
- 路径：`/api/v1/result/{job_id}`

用途：

- 查看本次任务的 `structured_result` 与 `member4_push` 结果。

### 3.4 下载批改后文档

- 方法：`GET`
- 路径：`/api/v1/download/{job_id}`
- 返回：文件流（二进制）

### 3.5 成员2调用示例（curl）

```bash
curl -X POST "http://127.0.0.1:8000/api/v1/grade" \
  -F "question_file=@D:/project/example/对象与类作业-题目.pdf" \
  -F "student_file=@D:/project/example/对象与类作业-待批改.doc" \
  -F "student_id=stu-001" \
  -F "teacher_material_files=@D:/project/example/对象与类作业-补充材料.docx" \
  -F "protocol=OpenAI兼容" \
  -F "api_key=你的APIKEY" \
  -F "base_url=https://ark.cn-beijing.volces.com/api/coding/v3" \
  -F "model=doubao-seed-2.0-pro"
```

下载批改后文档：

```bash
curl -L "http://127.0.0.1:8000/api/v1/download/<job_id>" -o graded.docx
```

---

## 4. 成员4如何对接（接收成员3推送）

成员3在每次批改结束后，会主动调用：

- 方法：`POST`
- 地址：`http://成员4服务地址:8007/api/v1/analytics/events/ingest`
- Body 要求：最外层必须是 `events`

成员4应兼容的请求体格式：

```json
{
  "events": [
    {
      "user_id": "stu-001",
      "class_id": "class-demo",
      "course_id": "course_oop_design",
      "event_type": "GRADING_RESULT_RECEIVED",
      "payload": {
        "source_id": "b9d1f8b4f0f54f25ae9d1b4f88123456",
        "source_type": "assignment",
        "results": [
          {
            "knowledge_point_id": "kp_class_and_object",
            "is_correct": true
          },
          {
            "knowledge_point_id": "kp_polymorphism",
            "is_correct": false
          }
        ]
      }
    }
  ]
}
```

成员4接收建议：

- `events` 为数组，按事件逐条处理。
- `knowledge_point_id` 必须按ID存储，不要按中文名匹配。
- `is_correct` 为布尔值，直接用于统计正确率。

成员4返回建议：

- 成功返回 `200`（或 `2xx`）+ JSON。
- 失败返回 `4xx/5xx` + 错误说明，便于成员3记录 `member4_push.error`。

---

## 5. 联调检查清单

成员2侧：

- 上传字段名与文档一致（尤其 `question_file`、`student_file`、`student_id`）。
- `student_id` 满足规则（字母/数字/_/-）。
- 先看 `download_url` 能否拿到批改文档。

成员4侧：

- `POST /api/v1/analytics/events/ingest` 已可用。
- 可接收 `events` 包裹格式。
- 可处理 `source_id=job_id` 的幂等或去重。

成员3侧：

- 若批改成功但 `member4_push.error` 非空，优先检查成员4服务日志。

---

## 6. 兼容说明

- Linux 可直接运行本项目，不依赖 Windows COM。
- 上传 `.doc/.xls` 旧格式时：
  - Linux 建议安装 LibreOffice（`soffice`）转换。
  - Windows 会优先尝试 `soffice`，若不可用再尝试本机 Office COM 转换。
- 建议生产统一上传 `.docx/.xlsx` 以降低转换失败风险。
