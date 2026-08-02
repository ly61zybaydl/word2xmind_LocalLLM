# Local Word to XMind

[中文](README_CN.md) | [English](README.md)

一个完全本地运行的 Python 工具：读取结构化 Word 需求文档，调用本地 Ollama 模型生成测试用例，并导出为兼容 XMind 8 的思维导图。

> 文档内容默认只发送到本机 `http://127.0.0.1:11434`，不会调用外部大模型 API。首次安装 Ollama 和下载模型仍需要网络。

## 功能特点

- 根据 Word 的“标题 1 / Heading 1”和“标题 2 / Heading 2”识别业务结构并分块，避免一次输入过长。
- 使用等价类、边界值、状态转移、异常流程、权限和数据校验等方法生成测试点。
- 通过 Ollama 在本地运行 `qwen3.5:4b`，默认温度为 `0.1`。
- 使用 JSON Schema 约束模型输出，并对 Markdown 包裹、缺失括号等常见 JSON 问题进行容错和二次修复。
- 生成并验证兼容 XMind 8 的 `.xmind` 文件。
- 支持自动启动本机 Ollama 服务、自定义服务地址、超时和日志级别。

## 工作流程

```text
Word (.docx)
    ↓ 按标题结构解析和分块
本地 Ollama 模型
    ↓ 生成并修复结构化 JSON
测试用例层级映射
    ↓
XMind (.xmind)
```

## 环境要求

- Python 3.10 或更高版本
- [Ollama](https://ollama.com/)
- 可运行所选模型的内存或显存空间
- XMind（仅在需要查看或编辑结果时使用）

## 快速开始

### 1. 获取代码

```bash
git clone https://github.com/<your-username>/local-word-to-xmind.git
cd local-word-to-xmind
```

如果尚未上传 GitHub，也可以直接进入本地项目目录（请替换路径占位符）：

```powershell
cd "<你的项目目录>\local-word-to-xmind"
```

### 2. 创建虚拟环境并安装依赖

Windows PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

macOS / Linux：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 3. 准备 Ollama 模型

安装并启动 Ollama，然后下载默认模型：

```bash
ollama pull qwen3.5:4b
```

脚本在默认本地服务不可连接时会尝试执行 `ollama serve`。也可以手动启动：

```bash
ollama serve
```

### 4. 转换文档

使用仓库中的示例文件：

```powershell
python word_to_xmind.py --input "sample_requirements.docx" --output "sample_test_cases.xmind"
```

使用自己的文档：

```bash
python word_to_xmind.py --input "software_requirements.docx" --output "test_cases.xmind"
```

运行完成后，用 XMind 打开输出文件即可。

## Word 文档格式

推荐使用以下层级：

```text
标题 1：业务域或大章节
  标题 2：具体功能模块
    正文：功能描述、业务规则、正常流程、异常处理等
```

脚本会将同一个二级标题下的正文组成一个模型输入块。若没有二级标题，则回退到一级标题分块；若完全没有标题样式，则将全文作为一个输入块。

当前版本读取 Word 段落内容，不解析表格、图片、批注或修订记录。建议先把关键需求整理为普通段落。

## XMind 输出结构

```text
文档标题（中心主题）
└── 业务模块
    └── 测试对象（需求编号）
        └── 测试点
            └── 输入与步骤 + 预期结果
```

## 命令行参数

| 参数 | 说明 | 默认值 |
| --- | --- | --- |
| `--input`, `-i` | 输入 `.docx` 文件，必填 | 无 |
| `--output`, `-o` | 输出 `.xmind` 文件 | `test_cases.xmind` |
| `--model`, `-m` | Ollama 模型名称 | `qwen3.5:4b` |
| `--temperature`, `-t` | 模型温度 | `0.1` |
| `--timeout` | 单次模型调用超时，单位为秒 | `600` |
| `--ollama-host` | Ollama 服务地址 | `OLLAMA_HOST` 或 `http://127.0.0.1:11434` |
| `--no-auto-start-ollama` | 禁止脚本自动启动本机 Ollama | 未启用 |
| `--startup-timeout` | 等待 Ollama 启动的秒数 | `30` |
| `--log-file` | 指定日志文件 | `logs/word_to_xmind_时间.log` |
| `--log-level` | 控制台日志级别 | `INFO` |

查看完整帮助：

```bash
python word_to_xmind.py --help
```

切换其他已安装的 Ollama 模型：

```bash
python word_to_xmind.py -i "sample_requirements.docx" -o "test_cases.xmind" -m "模型名称"
```

连接自定义服务地址：

```bash
python word_to_xmind.py -i "sample_requirements.docx" --ollama-host "http://127.0.0.1:11434"
```

> 如果 `--ollama-host` 或 `OLLAMA_HOST` 指向其他计算机，文档内容将发送到该地址，不再属于严格意义上的单机离线处理。

## 项目结构

```text
.
├── word_to_xmind.py      # 主程序
├── requirements.txt      # Python 依赖
├── sample_requirements.docx  # 示例需求文档
├── sample_test_cases.xmind   # 示例输出
├── project.md            # 原始需求说明
├── README.md             # 中文说明
└── README_EN.md          # English documentation
```

## 日志与隐私

程序默认在 `logs/` 中创建日志。日志文件会记录 DEBUG 级别信息，其中可能包含模型输出和需求相关内容，因此已通过 `.gitignore` 排除，不建议公开提交。

本项目无需 API Key。请仍在发布前检查示例 Word 文档的正文、作者、公司名称和文档属性，避免意外公开敏感信息。

## 常见问题

### 无法连接 Ollama

确认 Ollama 已安装并运行：

```bash
ollama serve
```

如果使用自定义端口，请通过 `--ollama-host` 或环境变量 `OLLAMA_HOST` 指定地址。

### 找不到模型

先下载默认模型，或通过 `--model` 使用已安装的模型：

```bash
ollama list
ollama pull qwen3.5:4b
```

### 模型返回的 JSON 无法解析

程序会自动清理和修复一次。如果某个模块仍失败，会记录错误并跳过该模块。可以尝试降低 `--temperature`、使用指令遵循能力更强的模型，或缩短对应模块的需求文本。

### XMind 文件无法打开

程序会自动补齐并验证 XMind 8 包结构。如果仍然失败，请保留日志，并确认安装了当前版本的依赖：

```bash
pip install -r requirements.txt --upgrade
```

## 参与贡献

欢迎提交 Issue 和 Pull Request。提交代码前，请确保：

- 不包含真实业务需求、账号、日志或其他敏感数据。
- 新功能不破坏默认的本地离线处理方式。
- 对 Word 解析、JSON 容错或 XMind 输出的变更附带可复现的测试样例。

## 开源许可证

仓库目前尚未包含许可证。正式公开前，请根据你的使用目标添加 `LICENSE` 文件；宽松开源项目通常可选择 MIT 或 Apache-2.0。
