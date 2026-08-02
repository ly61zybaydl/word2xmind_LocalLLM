# Local Word to XMind

[中文](README_CN.md) | [English](README.md)

A fully local Python tool that reads structured Word requirement documents, uses a local Ollama model to generate software test cases, and exports them as an XMind 8-compatible mind map.

> By default, document content is sent only to `http://127.0.0.1:11434`; no external LLM API is used. An internet connection is still required for the initial Ollama installation and model download.

## Features

- Detects `Heading 1` and `Heading 2` styles and splits the document by business module to avoid oversized model inputs.
- Generates test points using equivalence partitioning, boundary-value analysis, state transitions, exception flows, permission checks, and data validation.
- Runs `qwen3.5:4b` locally through Ollama with a default temperature of `0.1`.
- Constrains responses with a JSON Schema and handles common problems such as Markdown fences and missing closing brackets, including one model-assisted repair attempt.
- Builds and validates an XMind 8-compatible `.xmind` package.
- Supports automatic startup of a local Ollama service, custom hosts, timeouts, and log levels.

## Pipeline

```text
Word (.docx)
    ↓ Structure-aware parsing and chunking
Local Ollama model
    ↓ Structured JSON generation and repair
Test-case hierarchy mapping
    ↓
XMind (.xmind)
```

## Requirements

- Python 3.10 or later
- [Ollama](https://ollama.com/)
- Enough RAM or VRAM for the selected model
- XMind, only for viewing or editing the generated file

## Quick Start

### 1. Get the code

```bash
git clone https://github.com/<your-username>/local-word-to-xmind.git
cd local-word-to-xmind
```

Before publishing the repository, you can work directly from the local folder (replace the path placeholder):

```powershell
cd "<path-to-your-project>\local-word-to-xmind"
```

### 2. Create a virtual environment and install dependencies

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

macOS / Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Prepare the Ollama model

Install and start Ollama, then download the default model:

```bash
ollama pull qwen3.5:4b
```

If the default local endpoint is unavailable, the script attempts to run `ollama serve`. You can also start it manually:

```bash
ollama serve
```

### 4. Convert a document

Run the bundled example:

```powershell
python word_to_xmind.py --input "sample_requirements.docx" --output "sample_test_cases.xmind"
```

Run it with your own document:

```bash
python word_to_xmind.py --input "software-requirements.docx" --output "test-cases.xmind"
```

Open the resulting file in XMind when the command completes.

## Word Document Format

The recommended structure is:

```text
Heading 1: business domain or major section
  Heading 2: individual feature module
    Body: description, rules, normal flow, exception handling, and so on
```

All body paragraphs under the same second-level heading become one model input chunk. If no second-level headings exist, the tool falls back to first-level headings. If the document has no recognized heading styles, the whole document becomes one chunk.

The current version reads Word paragraphs only. It does not parse tables, images, comments, or tracked changes, so important requirements should be converted to regular paragraphs first.

## XMind Hierarchy

```text
Document title (central topic)
└── Business module
    └── Test object (requirement ID)
        └── Test point
            └── Inputs and steps + expected result
```

## Command-Line Options

| Option | Description | Default |
| --- | --- | --- |
| `--input`, `-i` | Input `.docx` file; required | None |
| `--output`, `-o` | Output `.xmind` file | `test_cases.xmind` |
| `--model`, `-m` | Ollama model name | `qwen3.5:4b` |
| `--temperature`, `-t` | Model temperature | `0.1` |
| `--timeout` | Timeout for one model call, in seconds | `600` |
| `--ollama-host` | Ollama service URL | `OLLAMA_HOST` or `http://127.0.0.1:11434` |
| `--no-auto-start-ollama` | Disable automatic local Ollama startup | Disabled |
| `--startup-timeout` | Seconds to wait for Ollama startup | `30` |
| `--log-file` | Custom log path | `logs/word_to_xmind_timestamp.log` |
| `--log-level` | Console log level | `INFO` |

View the complete CLI help:

```bash
python word_to_xmind.py --help
```

Use another locally installed Ollama model:

```bash
python word_to_xmind.py -i "requirements.docx" -o "test-cases.xmind" -m "model-name"
```

Connect to a custom Ollama endpoint:

```bash
python word_to_xmind.py -i "requirements.docx" --ollama-host "http://127.0.0.1:11434"
```

> If `--ollama-host` or `OLLAMA_HOST` points to another machine, document content is sent to that endpoint and processing is no longer strictly local to one computer.

## Project Layout

```text
.
├── word_to_xmind.py      # Main program
├── requirements.txt      # Python dependencies
├── sample_requirements.docx  # Example requirement document
├── sample_test_cases.xmind   # Example output
├── project.md            # Original project requirements
├── README.md             # Chinese documentation
└── README_EN.md          # English documentation
```

## Logging and Privacy

The program creates log files under `logs/` by default. These files contain DEBUG-level information that may include model output and requirement-related content. They are excluded through `.gitignore` and should not be committed publicly.

This project does not require an API key. Before publishing, still inspect the example Word document's body, author, company names, and document properties to avoid exposing sensitive information.

## Troubleshooting

### Ollama is unreachable

Make sure Ollama is installed and running:

```bash
ollama serve
```

For a custom port, set `--ollama-host` or the `OLLAMA_HOST` environment variable.

### The model is missing

Download the default model or select one that is already installed:

```bash
ollama list
ollama pull qwen3.5:4b
```

### The model returns invalid JSON

The program automatically cleans the response and attempts one repair. If a module still fails, it is logged and skipped. Try lowering `--temperature`, selecting a model with stronger instruction-following ability, or shortening the affected requirement section.

### XMind cannot open the generated file

The program adds and validates the required XMind 8 package files. If the issue persists, keep the logs and upgrade the current dependencies:

```bash
pip install -r requirements.txt --upgrade
```

## Contributing

Issues and pull requests are welcome. Before submitting code, please ensure that:

- No real business requirements, credentials, logs, or other sensitive data are included.
- New features preserve local-only processing as the default behavior.
- Changes to Word parsing, JSON recovery, or XMind generation include a reproducible test fixture.

## License

This project is licensed under the [MIT License](LICENSE).
