# 作业批注系统（题目多格式 + Word 批注）

支持流程：

1. 老师上传题目文件（`pdf/doc/docx/xls/xlsx`）
2. 可选上传老师补充材料（可多选，`pdf/doc/docx/xls/xlsx`）
3. 学生上传待批改作业（仅 `doc/docx`）
4. 可选上传老师历史批改样例（同格式）
5. 系统调用 `doubao-seed-2.0-pro` 生成批注并写回 Word 文档

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
- 批注输出：Word 原生“批注”形式
