# API 接口文档

后端入口文件：`backend_api.py`

启动方式（不运行Web页面）：

```powershell
conda activate homework_annotator
python run_backend.py
```

默认地址：`http://0.0.0.0:8000`

在线文档：

- Swagger UI: `http://127.0.0.1:8000/docs`
- ReDoc: `http://127.0.0.1:8000/redoc`

## 1. 健康检查

- 方法：`GET`
- 路径：`/health`
- 返回示例：

```json
{
  "status": "ok",
  "time": "2026-03-31T11:00:00.000000"
}
```

## 2. 提交批改任务

- 方法：`POST`
- 路径：`/api/v1/grade`
- Content-Type: `multipart/form-data`

### 表单字段

- `question_file` (file, 必填): 题目文件，支持 `pdf/doc/docx/xls/xlsx`
- `student_file` (file, 必填): 学生作业，支持 `doc/docx`
- `teacher_material_files` (file[], 可选): 老师补充材料，多文件
- `reference_file` (file, 可选): 老师批改样例
- `protocol` (text, 可选): `OpenAI兼容` 或 `Anthropic兼容`，默认 `OpenAI兼容`
- `api_key` (text, 必填): ARK API Key
- `base_url` (text, 可选): 默认 `https://ark.cn-beijing.volces.com/api/coding/v3`
- `model` (text, 可选): 默认 `doubao-seed-2.0-pro`

### 返回示例

```json
{
  "job_id": "b9d1f8b4f0f54f25ae9d1b4f88123456",
  "overall": "整体完成较好，建议加强函数公式和数据透视细节。",
  "output_file_name": "Excel实验报告-待批改-批改后-20260331_110000.docx",
  "download_url": "/api/v1/download/b9d1f8b4f0f54f25ae9d1b4f88123456"
}
```

## 3. 查询任务结果

- 方法：`GET`
- 路径：`/api/v1/result/{job_id}`

## 4. 下载批改后文件

- 方法：`GET`
- 路径：`/api/v1/download/{job_id}`
- 返回：文件流（二进制）

## curl 调用示例

```bash
curl -X POST "http://127.0.0.1:8000/api/v1/grade" \
  -F "question_file=@D:/project/example/Excel实验题目要求.pdf" \
  -F "teacher_material_files=@D:/project/example/Excel实验原始素材文件.xlsx" \
  -F "student_file=@D:/project/example/Excel实验报告-待批改.docx" \
  -F "reference_file=@D:/project/example/Excel实验报告-批改后.docx" \
  -F "protocol=OpenAI兼容" \
  -F "api_key=你的APIKEY" \
  -F "base_url=https://ark.cn-beijing.volces.com/api/coding/v3" \
  -F "model=doubao-seed-2.0-pro"
```

下载：

```bash
curl -L "http://127.0.0.1:8000/api/v1/download/<job_id>" -o graded.docx
```
