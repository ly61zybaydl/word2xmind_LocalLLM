# -*- coding: utf-8 -*-
"""
Word 需求转 XMind 测试导图工具

依赖安装：
    pip install python-docx ollama xmind

运行示例：
    python word_to_xmind.py --input "software_requirements.docx" --output "test_cases.xmind"
"""

from __future__ import annotations

import argparse
from datetime import datetime
import json
import logging
import os
import platform
import re
import subprocess
import sys
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from urllib.parse import urlparse


DEFAULT_MODEL = "qwen3.5:4b"
DEFAULT_OUTPUT = "test_cases.xmind"
DEFAULT_OLLAMA_HOST = "http://127.0.0.1:11434"

OLLAMA_OUTPUT_SCHEMA: dict[str, Any] = {
    "type": "object",
    "properties": {
        "module_name": {"type": "string"},
        "test_objects": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "object_name": {"type": "string"},
                    "req_id": {"type": "string"},
                    "test_points": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "point_name": {"type": "string"},
                                "inputs_and_steps": {"type": "string"},
                                "expected_result": {"type": "string"},
                            },
                            "required": ["point_name", "inputs_and_steps", "expected_result"],
                        },
                    },
                },
                "required": ["object_name", "req_id", "test_points"],
            },
        },
    },
    "required": ["module_name", "test_objects"],
}


@dataclass
class RequirementChunk:
    """按 Word 二级标题聚合出的一个业务需求文本块。"""

    module_name: str
    requirement_text: str
    heading1: str = ""


def _normalize_style_name(style_name: str) -> str:
    """统一样式名称，兼容英文 Heading 和中文标题样式。"""

    return re.sub(r"\s+", "", style_name.strip().lower())


def _is_heading_style(style_name: str, level: int) -> bool:
    """判断段落样式是否为指定级别标题。"""

    normalized = _normalize_style_name(style_name)
    return normalized in {f"heading{level}", f"标题{level}"}


def _get_paragraph_style_name(paragraph: Any) -> str:
    """安全读取段落样式名称，避免异常文档样式为空时报错。"""

    style = getattr(paragraph, "style", None)
    return getattr(style, "name", "") or ""


def extract_requirement_chunks(docx_path: Path) -> tuple[str, list[RequirementChunk]]:
    """
    阶段 1：Word 文档结构化读取与分块。

    核心策略：
    1. 使用 python-docx 的 Document 对象读取 Word，而不是直接抽纯文本。
    2. 识别 Heading 1 / 标题 1 作为大章节，Heading 2 / 标题 2 作为业务模块。
    3. 将同一个二级标题下的正文段落拼接成独立 chunk，避免一次性喂给本地模型导致上下文过长。
    """

    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError("缺少依赖 python-docx，请先执行：pip install python-docx") from exc

    document = Document(str(docx_path))
    core_title = (document.core_properties.title or "").strip()
    document_title = core_title or docx_path.stem
    title_from_heading = False

    chunks: list[RequirementChunk] = []
    current_heading1 = ""
    current_heading2 = ""
    current_lines: list[str] = []

    # 当遇到新的二级标题或文档结束时，将前一个模块落盘成 chunk。
    def flush_current_chunk() -> None:
        nonlocal current_heading2, current_lines
        body = "\n".join(line for line in current_lines if line.strip()).strip()
        if current_heading2 and body:
            module_title = current_heading2.strip()
            chunks.append(
                RequirementChunk(
                    module_name=module_title,
                    requirement_text=body,
                    heading1=current_heading1,
                )
            )
        current_lines = []

    # 兜底分组：如果文档没有二级标题，按一级标题聚合；再没有标题时按全文聚合。
    fallback_heading1_chunks: list[RequirementChunk] = []
    fallback_all_lines: list[str] = []
    fallback_h1_title = ""
    fallback_h1_lines: list[str] = []

    def flush_fallback_h1() -> None:
        nonlocal fallback_h1_lines
        body = "\n".join(line for line in fallback_h1_lines if line.strip()).strip()
        if fallback_h1_title and body:
            fallback_heading1_chunks.append(
                RequirementChunk(module_name=fallback_h1_title, requirement_text=body)
            )
        fallback_h1_lines = []

    for paragraph in document.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue

        style_name = _get_paragraph_style_name(paragraph)

        if _is_heading_style(style_name, 1):
            flush_current_chunk()
            flush_fallback_h1()
            current_heading1 = text
            current_heading2 = ""
            fallback_h1_title = text
            if not core_title and not title_from_heading:
                document_title = text
                title_from_heading = True
            continue

        if _is_heading_style(style_name, 2):
            flush_current_chunk()
            current_heading2 = text
            current_lines = []

            # 二级标题也放入一级标题兜底文本，便于无二级标题场景外的完整回退。
            if fallback_h1_title:
                fallback_h1_lines.append(text)
            fallback_all_lines.append(text)
            continue

        fallback_all_lines.append(text)
        if fallback_h1_title:
            fallback_h1_lines.append(text)

        if current_heading2:
            current_lines.append(text)

    flush_current_chunk()
    flush_fallback_h1()

    if chunks:
        return document_title, chunks

    if fallback_heading1_chunks:
        logging.warning("未识别到二级标题，将按一级标题生成需求分块。")
        return document_title, fallback_heading1_chunks

    fallback_body = "\n".join(fallback_all_lines).strip()
    if fallback_body:
        logging.warning("未识别到标题样式，将全文作为单个需求分块处理。")
        return document_title, [RequirementChunk(module_name=document_title, requirement_text=fallback_body)]

    return document_title, []


def build_prompt(chunk: RequirementChunk) -> str:
    """
    阶段 2：构建 Prompt 模板。

    这里强制要求模型只返回 JSON 字符串，且明确 JSON 字段结构，降低本地模型输出解释性文字的概率。
    """

    module_path = f"{chunk.heading1} / {chunk.module_name}" if chunk.heading1 else chunk.module_name
    schema_text = json.dumps(OLLAMA_OUTPUT_SCHEMA, ensure_ascii=False, indent=2)
    return f"""
你是一名资深软件测试工程师，正在根据软件需求规格说明书生成测试用例思维导图数据。

请基于以下需求文本，使用等价类划分、边界值分析、状态转移、异常流程、权限与数据校验等测试设计方法，挖掘完整且可执行的测试点。

当前模块标题：{module_path}

需求文本：
{chunk.requirement_text}

输出要求：
1. 只能输出一个纯 JSON 对象字符串。
2. 不允许输出 Markdown 代码块，不允许包含 ```json 或 ```。
3. 不允许输出解释、寒暄、推理过程或任何 JSON 之外的文字。
4. JSON 字符串必须可被 Python json.loads 直接解析。
5. 若需求编号不存在，req_id 使用空字符串。
6. 字符串内部如果需要使用双引号，必须转义为 \\\"，避免破坏 JSON。

必须符合以下 JSON Schema：
{schema_text}

请严格使用以下 JSON 结构：
{{
  "module_name": "{chunk.module_name}",
  "test_objects": [
    {{
      "object_name": "测试对象名称",
      "req_id": "需求编号（若有）",
      "test_points": [
        {{
          "point_name": "测试点（验证点）",
          "inputs_and_steps": "前置条件、输入及操作步骤",
          "expected_result": "预期结果与验证标准"
        }}
      ]
    }}
  ]
}}
""".strip()


def build_json_repair_prompt(raw_text: str, fallback_module_name: str) -> str:
    """构建 JSON 修复 Prompt，只要求模型修正语法，不重新发散生成测试点。"""

    schema_text = json.dumps(OLLAMA_OUTPUT_SCHEMA, ensure_ascii=False, indent=2)
    return f"""
你是一名 JSON 格式修复助手。

下面的文本本应是一个测试用例导图 JSON，但可能存在少逗号、多余字符、Markdown 包裹、字符串引号错误或结构不完整的问题。

请只做 JSON 语法修复和必要的字段补齐，不要新增解释，不要输出 Markdown，不要改变原有业务含义。

如果 module_name 缺失，使用：{fallback_module_name}
如果 req_id 缺失，使用空字符串。
字符串内部如果需要使用双引号，必须转义为 \\\"，避免破坏 JSON。

必须符合以下 JSON Schema：
{schema_text}

原始文本：
{raw_text}

请只输出符合以下结构的纯 JSON 对象：
{{
  "module_name": "{fallback_module_name}",
  "test_objects": [
    {{
      "object_name": "测试对象名称",
      "req_id": "",
      "test_points": [
        {{
          "point_name": "测试点",
          "inputs_and_steps": "输入与步骤",
          "expected_result": "预期结果"
        }}
      ]
    }}
  ]
}}
""".strip()


def _normalize_ollama_host(host: str) -> str:
    """将 Ollama 服务地址统一成 http://host:port 形式，便于 Python 客户端连接。"""

    stripped = host.strip()
    if not stripped:
        return DEFAULT_OLLAMA_HOST
    if re.match(r"^https?://", stripped, flags=re.IGNORECASE):
        return stripped
    return f"http://{stripped}"


def _effective_ollama_host(host: str | None) -> str:
    """优先使用命令行传入地址，其次使用环境变量 OLLAMA_HOST，最后使用 Ollama 默认地址。"""

    return _normalize_ollama_host(host or os.environ.get("OLLAMA_HOST", DEFAULT_OLLAMA_HOST))


def _is_local_ollama_host(host: str) -> bool:
    """判断 Ollama 地址是否指向本机；只有本机服务才适合由脚本自动启动。"""

    parsed = urlparse(host)
    hostname = (parsed.hostname or "").lower()
    return hostname in {"127.0.0.1", "localhost", "::1", "0.0.0.0"}


def _ollama_host_for_env(host: str) -> str:
    """把 URL 形式的 host 转成 ollama serve 更常用的 OLLAMA_HOST 值。"""

    parsed = urlparse(host)
    return parsed.netloc or host.replace("http://", "").replace("https://", "")


def _create_ollama_client(host: str | None, timeout: float) -> Any:
    """创建 Ollama Python 客户端，并在缺少依赖时给出明确安装提示。"""

    try:
        import ollama
    except ImportError as exc:
        raise RuntimeError("缺少依赖 ollama，请先执行：pip install ollama") from exc

    return ollama.Client(host=_effective_ollama_host(host), timeout=timeout)


def _can_connect_ollama(host: str | None, timeout: float = 5.0) -> bool:
    """通过 list 接口探测 Ollama 服务是否已经可连接。"""

    try:
        client = _create_ollama_client(host, timeout=timeout)
        client.list()
        return True
    except Exception as exc:
        logging.debug("Ollama 服务探测失败：%s", exc)
        return False


def _read_field(value: Any, name: str, default: Any = None) -> Any:
    """兼容 dict、Pydantic 对象和普通对象的字段读取。"""

    if isinstance(value, dict):
        return value.get(name, default)
    return getattr(value, name, default)


def _response_debug_preview(response: Any, limit: int = 1200) -> str:
    """生成响应对象预览，便于排查 200 OK 但内容为空的问题。"""

    if hasattr(response, "model_dump"):
        try:
            text = json.dumps(response.model_dump(), ensure_ascii=False, default=str)
        except Exception:
            text = repr(response)
    else:
        text = repr(response)

    return text if len(text) <= limit else text[:limit] + "...[truncated]"


def _extract_ollama_content(response: Any) -> str:
    """从不同版本的 Ollama 响应对象中提取模型正文。"""

    message = _read_field(response, "message", {})
    content = _read_field(message, "content", "")
    if content:
        return str(content)

    # generate 接口常见字段；这里作为兜底，兼容意外响应形态。
    legacy_response = _read_field(response, "response", "")
    if legacy_response:
        return str(legacy_response)

    thinking = _read_field(message, "thinking", "")
    if thinking:
        logging.error("Ollama 返回了 thinking 字段但正文 content 为空；脚本会使用 think=False 抑制思考模式。")

    logging.error("Ollama 原始响应预览：%s", _response_debug_preview(response))
    return ""


def _start_ollama_serve(host: str) -> subprocess.Popen[Any]:
    """
    在后台启动 ollama serve。

    Windows 下使用 CREATE_NO_WINDOW 避免弹出额外窗口；stdout/stderr 丢弃，主流程通过连接探测判断启动结果。
    """

    env = os.environ.copy()
    env["OLLAMA_HOST"] = _ollama_host_for_env(host)
    creationflags = subprocess.CREATE_NO_WINDOW if platform.system() == "Windows" else 0

    try:
        return subprocess.Popen(
            ["ollama", "serve"],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            env=env,
            creationflags=creationflags,
        )
    except FileNotFoundError as exc:
        raise RuntimeError("找不到 ollama 命令，请确认 Ollama CLI 已安装并加入 PATH。") from exc


def ensure_ollama_service(
    host: str | None,
    auto_start: bool = True,
    startup_timeout: float = 30.0,
) -> None:
    """
    确保本地 Ollama 服务可用。

    如果服务未启动且地址指向本机，则自动后台启动 ollama serve。模型不需要单独 run；
    后续 chat 请求会自动加载已下载的模型。
    """

    effective_host = _effective_ollama_host(host)
    if _can_connect_ollama(effective_host):
        logging.info("已连接 Ollama 服务：%s", effective_host)
        return

    if not auto_start:
        raise RuntimeError(f"Ollama 服务不可连接：{effective_host}。请先手动启动 ollama serve。")

    if not _is_local_ollama_host(effective_host):
        raise RuntimeError(
            f"Ollama 服务不可连接：{effective_host}。该地址不是本机地址，脚本不会自动启动远程服务。"
        )

    logging.warning("Ollama 服务不可连接，正在后台启动：ollama serve（%s）", effective_host)
    process = _start_ollama_serve(effective_host)

    deadline = time.perf_counter() + startup_timeout
    while time.perf_counter() < deadline:
        time.sleep(1)
        if _can_connect_ollama(effective_host):
            logging.info("Ollama 服务已启动并可连接：%s", effective_host)
            return
        if process.poll() is not None:
            break

    raise RuntimeError(
        f"已尝试自动启动 Ollama，但在 {startup_timeout:.0f} 秒内仍无法连接：{effective_host}。"
    )


def call_ollama(
    prompt: str,
    model: str = DEFAULT_MODEL,
    temperature: float = 0.1,
    timeout: float = 600.0,
    host: str | None = None,
    response_format: Any = "json",
) -> str:
    """
    阶段 3：调用本地 Ollama 推理。

    使用 ollama Python 库连接本地服务；format="json" 会请求模型以 JSON 模式输出。
    注意：这里不调用任何外部网络 API，只访问本机 Ollama 服务。
    """

    # host 可来自命令行参数、环境变量 OLLAMA_HOST 或默认端口；传入时可连接自定义端口或远程内网服务。
    client = _create_ollama_client(host, timeout=timeout)
    response = client.chat(
        model=model,
        messages=[{"role": "user", "content": prompt}],
        format=response_format,
        think=False,
        options={"temperature": temperature},
        stream=False,
    )

    # ollama 包不同版本的响应对象形态略有差异，这里统一做兼容提取。
    content = _extract_ollama_content(response)

    if not content:
        raise RuntimeError("Ollama 返回内容为空，请检查服务状态和模型输出。")

    return str(content)


def _json_error_context(text: str, error: json.JSONDecodeError, radius: int = 120) -> str:
    """截取 JSON 错误位置附近文本，便于定位模型漏逗号或引号的问题。"""

    start = max(error.pos - radius, 0)
    end = min(error.pos + radius, len(text))
    return text[start:end].replace("\n", "\\n")


def _extract_balanced_json_objects(text: str) -> list[str]:
    """从混杂文本中提取花括号平衡的 JSON 对象候选，兼容 Markdown 或解释性前后缀。"""

    objects: list[str] = []
    start: int | None = None
    depth = 0
    in_string = False
    escaped = False

    for index, char in enumerate(text):
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue

        if char == '"':
            in_string = True
            continue

        if char == "{":
            if depth == 0:
                start = index
            depth += 1
            continue

        if char == "}" and depth:
            depth -= 1
            if depth == 0 and start is not None:
                objects.append(text[start : index + 1].strip())
                start = None

    return objects


def _complete_unclosed_json(text: str) -> str:
    """补齐模型输出末尾缺失的 JSON 括号，并去掉闭合括号前多余逗号。"""

    stack: list[str] = []
    in_string = False
    escaped = False
    completed = text.strip()

    for char in completed:
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue

        if char == '"':
            in_string = True
        elif char == "{":
            stack.append("}")
        elif char == "[":
            stack.append("]")
        elif char in {"}", "]"}:
            if stack and stack[-1] == char:
                stack.pop()

    if in_string:
        completed += '"'

    if stack:
        completed += "".join(reversed(stack))

    return re.sub(r",\s*([}\]])", r"\1", completed)


def parse_model_json(raw_text: str, log_failure: bool = True) -> dict[str, Any] | None:
    """
    阶段 4：JSON 解析与异常处理。

    先直接 json.loads；若模型输出了 Markdown 代码块，则剥离 ```json / ``` 后重试。
    如果仍失败，再尝试截取首个 JSON 对象范围。最终失败则返回 None，由主流程跳过该模块。
    """

    candidates: list[str] = [raw_text.strip()]

    cleaned = re.sub(r"^\s*```(?:json)?\s*", "", raw_text.strip(), flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```\s*$", "", cleaned).strip()
    if cleaned not in candidates:
        candidates.append(cleaned)

    for base_candidate in list(candidates):
        json_start = base_candidate.find("{")
        if json_start >= 0:
            from_first_object = base_candidate[json_start:].strip()
            if from_first_object not in candidates:
                candidates.append(from_first_object)

            completed = _complete_unclosed_json(from_first_object)
            if completed not in candidates:
                candidates.append(completed)

    for extracted in _extract_balanced_json_objects(raw_text):
        if extracted not in candidates:
            candidates.append(extracted)

    last_error: Exception | None = None
    last_candidate = ""
    for candidate in candidates:
        try:
            parsed = json.loads(candidate)
        except json.JSONDecodeError as exc:
            last_error = exc
            last_candidate = candidate
            continue

        if isinstance(parsed, dict):
            return parsed

        if log_failure:
            logging.error("模型 JSON 顶层不是对象，实际类型：%s", type(parsed).__name__)
        return None

    if log_failure:
        logging.error("JSON 解析失败。错误：%s", last_error)
        if isinstance(last_error, json.JSONDecodeError) and last_candidate:
            logging.error("错误位置附近内容：%s", _json_error_context(last_candidate, last_error))
    return None


def repair_model_json(
    raw_text: str,
    fallback_module_name: str,
    model: str,
    temperature: float,
    timeout: float,
    host: str | None,
) -> dict[str, Any] | None:
    """当模型第一次返回非法 JSON 时，使用本地模型做一次 JSON 语法修复。"""

    logging.warning("模型返回的 JSON 解析失败，正在调用本地模型进行一次 JSON 修复。")
    logging.debug("待修复 JSON 原始文本（模块：%s）：\n%s", fallback_module_name, raw_text)
    repair_prompt = build_json_repair_prompt(raw_text, fallback_module_name)
    repaired_text = call_ollama(
        repair_prompt,
        model=model,
        temperature=min(temperature, 0.1),
        timeout=timeout,
        host=host,
        response_format="json",
    )
    logging.debug("JSON 修复输出（模块：%s）：\n%s", fallback_module_name, repaired_text)
    return parse_model_json(repaired_text, log_failure=True)


def _as_text(value: Any, default: str = "") -> str:
    """将模型返回的任意值安全转换为 XMind 节点文本。"""

    if value is None:
        return default
    text = str(value).strip()
    return text or default


def _as_list(value: Any) -> list[Any]:
    """模型偶尔会把列表字段返回成非列表，这里统一做保护。"""

    return value if isinstance(value, list) else []


def normalize_module_json(module_json: dict[str, Any], fallback_module_name: str) -> dict[str, Any]:
    """
    对模型 JSON 做轻量归一化，避免缺字段导致生成 XMind 时报错。

    这里不改变业务含义，只补默认值、过滤非对象条目。
    """

    normalized_objects: list[dict[str, Any]] = []
    for test_object in _as_list(module_json.get("test_objects")):
        if not isinstance(test_object, dict):
            continue

        normalized_points: list[dict[str, str]] = []
        for point in _as_list(test_object.get("test_points")):
            if not isinstance(point, dict):
                continue
            normalized_points.append(
                {
                    "point_name": _as_text(point.get("point_name"), "未命名测试点"),
                    "inputs_and_steps": _as_text(point.get("inputs_and_steps"), "未提供输入与步骤"),
                    "expected_result": _as_text(point.get("expected_result"), "未提供预期结果"),
                }
            )

        normalized_objects.append(
            {
                "object_name": _as_text(test_object.get("object_name"), "未命名测试对象"),
                "req_id": _as_text(test_object.get("req_id")),
                "test_points": normalized_points,
            }
        )

    return {
        "module_name": _as_text(module_json.get("module_name"), fallback_module_name),
        "test_objects": normalized_objects,
    }


def _xmind_meta_xml() -> str:
    """生成 XMind 8 兼容的 meta.xml。"""

    now = datetime.now().astimezone().isoformat(timespec="seconds")
    return f"""<?xml version="1.0" encoding="UTF-8"?>
<xmap-meta xmlns="urn:xmind:xmap:xmlns:meta:2.0" version="2.0">
  <Creator>
    <Name>word_to_xmind.py</Name>
    <Version>1.0</Version>
  </Creator>
  <Created>{now}</Created>
  <Modified>{now}</Modified>
</xmap-meta>
"""


def _xmind_manifest_xml() -> str:
    """生成 XMind 8 兼容的 META-INF/manifest.xml。"""

    return """<?xml version="1.0" encoding="UTF-8"?>
<manifest xmlns="urn:xmind:xmap:xmlns:manifest:1.0">
  <file-entry full-path="/" media-type="application/vnd.xmind.workbook"/>
  <file-entry full-path="content.xml" media-type="text/xml"/>
  <file-entry full-path="styles.xml" media-type="text/xml"/>
  <file-entry full-path="comments.xml" media-type="text/xml"/>
  <file-entry full-path="meta.xml" media-type="text/xml"/>
</manifest>
"""


def make_xmind8_compatible(output_path: Path) -> None:
    """
    补齐 XMind 8 兼容包结构。

    xmind 1.2.0 生成的 zip 往往只有 content.xml/styles.xml/comments.xml，
    部分新版 XMind 客户端会把缺少 manifest/meta 的文件提示为损坏。
    """

    if not zipfile.is_zipfile(output_path):
        raise RuntimeError(f"生成的 XMind 文件不是合法 zip 包：{output_path}")

    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    if tmp_path.exists():
        tmp_path.unlink()

    with zipfile.ZipFile(output_path, "r") as source_zip:
        existing_names = source_zip.namelist()
        payloads = {name: source_zip.read(name) for name in existing_names}

    required_names = {"content.xml", "styles.xml", "comments.xml"}
    missing_names = sorted(required_names - set(payloads))
    if missing_names:
        raise RuntimeError(f"生成的 XMind 文件缺少必要内容：{', '.join(missing_names)}")

    payloads["meta.xml"] = payloads.get("meta.xml", _xmind_meta_xml().encode("utf-8"))
    payloads["META-INF/manifest.xml"] = _xmind_manifest_xml().encode("utf-8")

    try:
        with zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
            for name in ["content.xml", "styles.xml", "comments.xml", "meta.xml", "META-INF/manifest.xml"]:
                target_zip.writestr(name, payloads[name])

            # 保留将来可能存在的附件、缩略图、标记等资源，避免补包时误删。
            for name, data in payloads.items():
                if name not in {"content.xml", "styles.xml", "comments.xml", "meta.xml", "META-INF/manifest.xml"}:
                    target_zip.writestr(name, data)

        tmp_path.replace(output_path)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()


def validate_xmind_package(output_path: Path) -> None:
    """验证生成的 XMind 包结构和关键 XML，提前发现会导致客户端打不开的问题。"""

    import xml.etree.ElementTree as ET

    if not zipfile.is_zipfile(output_path):
        raise RuntimeError(f"XMind 文件不是合法 zip 包：{output_path}")

    with zipfile.ZipFile(output_path, "r") as xmind_zip:
        names = set(xmind_zip.namelist())
        required_names = {"content.xml", "styles.xml", "comments.xml", "meta.xml", "META-INF/manifest.xml"}
        missing_names = sorted(required_names - names)
        if missing_names:
            raise RuntimeError(f"XMind 文件缺少必要条目：{', '.join(missing_names)}")

        for xml_name in required_names:
            ET.fromstring(xmind_zip.read(xml_name))

    logging.info("XMind 包结构验证通过：%s", output_path)


def configure_logging(log_file: Path | None, log_level: str = "INFO") -> Path:
    """
    配置控制台日志和文件日志。

    控制台默认只显示 INFO 及以上；日志文件记录 DEBUG 级别，包含模型原始输出和修复过程。
    """

    if log_file is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = Path.cwd() / "logs" / f"word_to_xmind_{timestamp}.log"

    log_file.parent.mkdir(parents=True, exist_ok=True)

    root_logger = logging.getLogger()
    root_logger.handlers.clear()
    root_logger.setLevel(logging.DEBUG)

    console_handler = logging.StreamHandler()
    console_handler.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    console_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(
        logging.Formatter("%(asctime)s %(levelname)s [%(name)s] %(message)s")
    )

    root_logger.addHandler(console_handler)
    root_logger.addHandler(file_handler)

    logging.getLogger("httpx").setLevel(logging.WARNING)
    logging.getLogger("httpcore").setLevel(logging.WARNING)
    logging.info("日志文件：%s", log_file)
    return log_file


def build_xmind_file(document_title: str, modules: list[dict[str, Any]], output_path: Path) -> None:
    """
    阶段 5：生成 XMind 文件。

    映射关系：
    - 中心主题：Word 文档标题
    - 一级节点：module_name
    - 二级节点：object_name
    - 三级节点：point_name
    - 四级节点：inputs_and_steps + expected_result
    """

    try:
        import xmind
    except ImportError as exc:
        raise RuntimeError("缺少依赖 xmind，请先执行：pip install xmind") from exc

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 生成新工作簿，避免同名旧文件中的历史节点残留。
    if output_path.exists():
        output_path.unlink()

    workbook = xmind.load(str(output_path))
    sheet = workbook.getPrimarySheet()
    sheet.setTitle("测试用例导图")

    root_topic = sheet.getRootTopic()
    root_topic.setTitle(_as_text(document_title, "测试用例导图"))

    for module in modules:
        module_topic = root_topic.addSubTopic()
        module_topic.setTitle(_as_text(module.get("module_name"), "未命名模块"))

        for test_object in _as_list(module.get("test_objects")):
            if not isinstance(test_object, dict):
                continue

            object_title = _as_text(test_object.get("object_name"), "未命名测试对象")
            req_id = _as_text(test_object.get("req_id"))
            if req_id:
                object_title = f"{object_title}（{req_id}）"

            object_topic = module_topic.addSubTopic()
            object_topic.setTitle(object_title)

            for point in _as_list(test_object.get("test_points")):
                if not isinstance(point, dict):
                    continue

                point_topic = object_topic.addSubTopic()
                point_topic.setTitle(_as_text(point.get("point_name"), "未命名测试点"))

                leaf_topic = point_topic.addSubTopic()
                leaf_topic.setTitle(
                    "输入与步骤："
                    + _as_text(point.get("inputs_and_steps"), "未提供")
                    + "\n预期结果："
                    + _as_text(point.get("expected_result"), "未提供")
                )

    xmind.save(workbook, path=str(output_path))
    make_xmind8_compatible(output_path)
    validate_xmind_package(output_path)


def parse_args(argv: list[str]) -> argparse.Namespace:
    """解析命令行参数。"""

    parser = argparse.ArgumentParser(description="将 Word 需求文档转换为 XMind 测试用例导图。")
    parser.add_argument("--input", "-i", required=True, help="输入 Word .docx 文件路径。")
    parser.add_argument(
        "--output",
        "-o",
        default=DEFAULT_OUTPUT,
        help=f"输出 .xmind 文件路径，默认：{DEFAULT_OUTPUT}",
    )
    parser.add_argument("--model", "-m", default=DEFAULT_MODEL, help=f"Ollama 模型名，默认：{DEFAULT_MODEL}")
    parser.add_argument(
        "--temperature",
        "-t",
        type=float,
        default=0.1,
        help="Ollama temperature 参数，默认：0.1",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=600.0,
        help="单次 Ollama 调用超时时间（秒），默认：600",
    )
    parser.add_argument(
        "--ollama-host",
        default=None,
        help=f"Ollama 服务地址，例如：{DEFAULT_OLLAMA_HOST}。不传则优先使用 OLLAMA_HOST，否则使用默认地址。",
    )
    parser.add_argument(
        "--no-auto-start-ollama",
        action="store_false",
        dest="auto_start_ollama",
        help="服务不可连接时不自动后台启动 ollama serve。",
    )
    parser.add_argument(
        "--startup-timeout",
        type=float,
        default=30.0,
        help="自动启动 Ollama 后等待服务可连接的时间（秒），默认：30",
    )
    parser.add_argument(
        "--log-file",
        default=None,
        help="日志文件路径。不传则写入当前目录 logs/word_to_xmind_时间.log。",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="控制台日志级别，默认：INFO。日志文件始终记录 DEBUG 细节。",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """主流程：串联五个阶段，完成 Word -> Ollama JSON -> XMind。"""

    args = parse_args(argv or sys.argv[1:])
    log_file = Path(args.log_file).expanduser().resolve() if args.log_file else None
    configure_logging(log_file, args.log_level)

    input_path = Path(args.input).expanduser().resolve()
    output_path = Path(args.output).expanduser().resolve()

    if not input_path.exists():
        logging.error("输入文件不存在：%s", input_path)
        return 1

    if input_path.suffix.lower() != ".docx":
        logging.error("输入文件必须是 .docx 格式：%s", input_path)
        return 1

    try:
        document_title, chunks = extract_requirement_chunks(input_path)
        if not chunks:
            logging.error("未从 Word 文档中提取到任何需求文本块。")
            return 1

        logging.info("文档标题：%s", document_title)
        logging.info("共提取到 %d 个需求分块。", len(chunks))

        ensure_ollama_service(
            args.ollama_host,
            auto_start=args.auto_start_ollama,
            startup_timeout=args.startup_timeout,
        )

        module_results: list[dict[str, Any]] = []
        for index, chunk in enumerate(chunks, start=1):
            logging.info("正在处理第 %d/%d 个模块：%s", index, len(chunks), chunk.module_name)

            prompt = build_prompt(chunk)
            started_at = time.perf_counter()
            logging.info("正在调用本地 Ollama 模型 %s，请等待模型返回。", args.model)
            raw_response = call_ollama(
                prompt,
                model=args.model,
                temperature=args.temperature,
                timeout=args.timeout,
                host=args.ollama_host,
            )
            elapsed = time.perf_counter() - started_at
            logging.info("Ollama 已返回，第 %d 个模块耗时 %.1f 秒。", index, elapsed)
            logging.debug("模型原始输出（模块：%s）：\n%s", chunk.module_name, raw_response)
            parsed_json = parse_model_json(raw_response, log_failure=False)
            if parsed_json is None:
                parsed_json = repair_model_json(
                    raw_response,
                    fallback_module_name=chunk.module_name,
                    model=args.model,
                    temperature=args.temperature,
                    timeout=args.timeout,
                    host=args.ollama_host,
                )
            if parsed_json is None:
                logging.error("当前模块修复后仍不是有效 JSON，将跳过：%s", chunk.module_name)
                continue

            module_results.append(normalize_module_json(parsed_json, chunk.module_name))

        if not module_results:
            logging.error("所有模块均未生成有效 JSON，无法生成 XMind 文件。")
            return 1

        build_xmind_file(document_title, module_results, output_path)
        logging.info("XMind 文件已生成：%s", output_path)
        return 0

    except RuntimeError as exc:
        logging.error("%s", exc)
        return 1
    except Exception:
        logging.exception("处理过程中发生未预期错误。")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
