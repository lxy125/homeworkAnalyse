# 作业批注系统（题目多格式 + Word 批注）

支持流程：

1. 老师上传题目文件（`pdf/doc/docx/xls/xlsx`）
2. 可选上传老师补充材料（可多选，`pdf/doc/docx/xls/xlsx`）
3. 填写学生ID（必填，用于输出命名）
4. 学生上传待批改作业（仅 `doc/docx`）
5. 可选上传老师历史批改样例（同格式）
6. 系统调用 `doubao-seed-2.0-pro` 生成批注并输出 Word 文档（文件名包含学生ID）
7. 同时输出结构化知识点正确性结果，并主动推送到成员4接口（`events/ingest`）

## 环境

```powershell
conda activate homework_annotator
pip install -r requirements.txt
```

## 配置

`.env` 示例：

```env
ARK_API_KEY=你的apikey
ARK_BASE_URL=https://ark.cn-beijing.volces.com/api/coding/v3
ARK_MODEL=doubao-seed-2.0-pro
MEMBER4_INGEST_URL=http://127.0.0.1:8007/api/v1/analytics/events/ingest
```

支持两种协议：

- OpenAI 兼容：`https://ark.cn-beijing.volces.com/api/coding/v3`
- Anthropic 兼容：`https://ark.cn-beijing.volces.com/api/coding`（程序内自动补全 `/v1/messages`）

## Web 运行

```powershell
conda activate homework_annotator
streamlit run app.py
```

## 后端运行（不启动Web）

```powershell
conda activate homework_annotator
python run_backend.py
```

默认监听 `0.0.0.0:8000`，可用环境变量覆盖：

- `API_HOST`：默认 `0.0.0.0`
- `API_PORT`：默认 `8000`
- `API_RELOAD`：默认 `false`

接口文档见：[API_DOCS.md](API_DOCS.md)

## 样例一键运行

已按你的目录约定预置：

- `example/对象与类作业-题目.pdf`
- `example/对象与类作业-待批改.doc`
- `example/对象与类作业-批改后.doc`

执行：

```powershell
conda activate homework_annotator
python run_example.py
```

输出文件会保存到 `outputs/`。

## 批注规则

- 题目文件：支持读取 `.doc/.docx/.pdf/.xls/.xlsx`
- 学生作业：仅支持 `.doc/.docx`
- 批注输出：`docx`，以 Word 原生“批注（评论气泡）”写入
- 结构化输出：`events` 包裹格式，`results` 为知识点编码 + 布尔正确性
- 固定值：`class_id=class-demo`、`course_id=course_oop_design`、`source_type=assignment`

Linux 部署说明：

- 本项目已移除 Windows COM 依赖，可直接运行在 Linux。
- 若上传 `.doc/.xls` 旧格式，需服务器安装 LibreOffice（`soffice`）用于自动转换。
- 建议优先上传 `.docx/.xlsx`，可避免格式转换依赖。
