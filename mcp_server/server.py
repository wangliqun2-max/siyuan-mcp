"""
server.py - MCP Server for Tender Parameter Extraction
Exposes 3 MCP tools via SSE transport + a simple file-upload web page.

Architecture:
  GET  /            → simple HTML upload page
  POST /upload      → receive PDF/DOCX, parse, return session_id
  GET  /sse         → MCP SSE endpoint (BISHENG connects here)
  POST /messages/   → MCP message endpoint (used internally by SSE protocol)
"""

import os
import sys
import json
import uuid
import tempfile
import asyncio
from pathlib import Path

# ── Project root on path so backend.services can be imported ──────────────────
ROOT_DIR = Path(__file__).parent.parent
sys.path.insert(0, str(ROOT_DIR))

from dotenv import load_dotenv
load_dotenv(ROOT_DIR / ".env")

# ── Backend service imports ────────────────────────────────────────────────────
from backend.services.doc_parser import (
    parse_document,
    get_section_text,
    keyword_density_find_section,
    find_section_in_headings,
    get_doc_summary_for_llm,
    SITE_CONDITIONS_HEADING_KEYWORDS,
)
from backend.services.llm_extractor import (
    find_section_from_headings,
    find_product_section,
    extract_parameters,
    get_param_list_from_template,
)
from backend.services.excel_handler import read_template_params

TEMPLATE_PATH = str(ROOT_DIR / os.getenv("EXCEL_TEMPLATE_PATH", "Checklist模板.xlsx"))
ALLOWED_EXTENSIONS = {".pdf", ".docx", ".doc"}

# ── FastMCP Server ─────────────────────────────────────────────────────────────
from fastmcp import FastMCP

mcp = FastMCP(
    "标书参数抽取",
    instructions=(
        "这是一个电力工程招标文件参数抽取工具，专用于变压器技术参数提取。\n"
        "使用流程：\n"
        "1. 用户先通过上传页面上传PDF/DOCX，获取 session_id\n"
        "2. 调用 locate_section 定位变压器章节\n"
        "3. 确认页码范围后调用 extract_params 抽取参数\n"
        "4. 调用 get_results_summary 获取 Markdown 格式汇总表"
    ),
)


# ── Helper ─────────────────────────────────────────────────────────────────────
def _load_session(session_id: str):
    """Load chunks and headings for a session. Raises FileNotFoundError if expired."""
    chunks_path = os.path.join(tempfile.gettempdir(), f"{session_id}_chunks.json")
    headings_path = os.path.join(tempfile.gettempdir(), f"{session_id}_headings.json")
    if not os.path.exists(chunks_path):
        raise FileNotFoundError(f"Session '{session_id}' 不存在或已过期，请重新上传文件")
    with open(chunks_path, "r", encoding="utf-8") as f:
        chunks = json.load(f)
    headings = []
    if os.path.exists(headings_path):
        with open(headings_path, "r", encoding="utf-8") as f:
            headings = json.load(f)
    return chunks, headings


def _merge(base: dict, new_r: dict) -> dict:
    """Merge extraction results: found=True always wins."""
    for k, v in new_r.items():
        if k not in base or (
            isinstance(v, dict) and v.get("found") and not base.get(k, {}).get("found")
        ):
            base[k] = v
    return base


# ── MCP Tool 1: locate_section ─────────────────────────────────────────────────
@mcp.tool()
def locate_section(session_id: str, product_type: str = "power_transformer") -> dict:
    """
    从已上传的招标文件中定位目标产品的技术规范章节。

    Args:
        session_id: 上传文件后获得的会话ID
        product_type: 产品类型，默认 "power_transformer"（主变压器）

    Returns:
        {section_title, start_page, end_page, confidence, notes, total_pages}
    """
    try:
        chunks, headings = _load_session(session_id)
    except FileNotFoundError as e:
        return {"error": str(e)}

    # Use physical_page for user-facing total — falls back to chunk page for PDFs
    total_pages = max((c.get("physical_page", c.get("page", 0)) for c in chunks), default=0)
    section_result = None

    # Step 1: Heading-based LLM detection
    if headings:
        try:
            result = find_section_from_headings(headings, product_type=product_type)
            if result.get("start_page") is not None:
                section_result = result
        except Exception as e:
            print(f"[locate_section] heading detection error: {e}")

    # Step 2: Keyword-density fallback
    if section_result is None:
        result = keyword_density_find_section(chunks)
        if result.get("start_page") is not None:
            section_result = result

    # Step 3: Full-doc LLM (last resort)
    if section_result is None:
        doc_summary = get_doc_summary_for_llm(chunks, max_chars=40000)
        try:
            section_result = find_product_section(doc_summary, product_type)
        except Exception as e:
            section_result = {
                "section_title": "", "start_page": None,
                "end_page": None, "confidence": 0.0, "notes": str(e),
            }

    section_result["total_pages"] = total_pages
    return section_result


# ── MCP Tool 2: extract_params ─────────────────────────────────────────────────
@mcp.tool()
def extract_params(session_id: str, start_page: int, end_page: int) -> dict:
    """
    从已定位的章节中抽取变压器技术参数。

    Args:
        session_id: 上传文件后获得的会话ID
        start_page: 章节起始页（来自 locate_section 的 start_page）
        end_page:   章节终止页（来自 locate_section 的 end_page）

    Returns:
        {success, stats: {total, found, not_found}, extracted: {param: {value, unit, source_text, found}}}
    """
    try:
        chunks, headings = _load_session(session_id)
    except FileNotFoundError as e:
        return {"error": str(e)}

    # Load parameter list from Excel template
    template_info = read_template_params(TEMPLATE_PATH)
    all_params = template_info["params"]
    param_list = get_param_list_from_template(template_info)
    sections_ctx = template_info.get("sections", [])

    CORE_SECTIONS = {2, 3, 4}
    site_params = [p["name"] for p in all_params if p.get("section_num") == 1]
    core_params = [p["name"] for p in all_params if p.get("section_num") in CORE_SECTIONS]
    extra_params = [
        p["name"] for p in all_params
        if p.get("section_num", 0) not in CORE_SECTIONS and p.get("section_num", 0) != 1
    ]

    core_set, extra_set, site_set = set(core_params), set(extra_params), set(site_params)
    site_sections = [s for s in sections_ctx if any(p in site_set for p in s.get("params", []))]
    core_sections = [s for s in sections_ctx if any(p in core_set for p in s.get("params", []))]
    extra_sections = [s for s in sections_ctx if any(p in extra_set for p in s.get("params", []))]

    extracted_all: dict = {}

    # Core electrical params
    if core_params:
        text = get_section_text(chunks, start_page, end_page)
        if text.strip():
            result = extract_parameters(text, core_params, sections_context=core_sections, max_chars=200_000)
            _merge(extracted_all, result)

    # Extra params
    if extra_params:
        text = get_section_text(chunks, start_page, end_page)
        if text.strip():
            result = extract_parameters(text, extra_params, sections_context=extra_sections, max_chars=200_000)
            _merge(extracted_all, result)

    # Site/environmental params (different section in document)
    if site_params:
        site_sec = find_section_in_headings(headings, SITE_CONDITIONS_HEADING_KEYWORDS)
        if site_sec.get("start_page"):
            text = get_section_text(chunks, site_sec["start_page"], site_sec["end_page"])
        else:
            text = get_section_text(chunks, 1, 50)
        if text.strip():
            result = extract_parameters(text, site_params, sections_context=site_sections)
            _merge(extracted_all, result)

    # Reorder to match template order
    extracted_ordered = {
        name: extracted_all.get(
            name, {"value": None, "unit": "", "source_text": "", "found": False}
        )
        for name in param_list
    }

    found_count = sum(1 for v in extracted_ordered.values() if isinstance(v, dict) and v.get("found"))

    # Persist results for get_results_summary
    results_path = os.path.join(tempfile.gettempdir(), f"{session_id}_results.json")
    with open(results_path, "w", encoding="utf-8") as f:
        json.dump(extracted_ordered, f, ensure_ascii=False)

    return {
        "success": True,
        "stats": {
            "total": len(param_list),
            "found": found_count,
            "not_found": len(param_list) - found_count,
        },
        "extracted": extracted_ordered,
    }


# ── MCP Tool 3: get_results_summary ───────────────────────────────────────────
@mcp.tool()
def get_results_summary(session_id: str) -> str:
    """
    获取参数抽取结果的 Markdown 格式汇总表（在 extract_params 之后调用）。

    Args:
        session_id: 会话ID

    Returns:
        人类可读的 Markdown 表格字符串
    """
    results_path = os.path.join(tempfile.gettempdir(), f"{session_id}_results.json")
    if not os.path.exists(results_path):
        return "❌ 未找到抽取结果，请先调用 extract_params 工具"

    with open(results_path, "r", encoding="utf-8") as f:
        extracted = json.load(f)

    found_items = [(k, v) for k, v in extracted.items() if isinstance(v, dict) and v.get("found")]
    not_found = [k for k, v in extracted.items() if isinstance(v, dict) and not v.get("found")]

    lines = [
        f"## 变压器参数抽取结果\n",
        f"**共找到 {len(found_items)} / {len(extracted)} 个参数**\n",
        "| 参数名 | 值 | 单位 | 原文摘要 |",
        "|--------|-----|------|---------|",
    ]
    for name, info in found_items:
        val = info.get("value") or ""
        unit = info.get("unit") or ""
        src = (info.get("source_text") or "")[:40]
        lines.append(f"| {name} | {val} | {unit} | {src} |")

    if not_found:
        lines.append(f"\n**未找到的参数（{len(not_found)} 个）：** " + "、".join(not_found[:20]))
        if len(not_found) > 20:
            lines.append(f"… 等共 {len(not_found)} 个")

    return "\n".join(lines)


# ── File Upload Web Page ───────────────────────────────────────────────────────
UPLOAD_HTML = """<!DOCTYPE html>
<html lang="zh">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>标书参数抽取 · 文件上传</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { font-family: 'PingFang SC', Arial, sans-serif; background: #f0f4ff; min-height: 100vh; display: flex; align-items: center; justify-content: center; }
    .card { background: #fff; border-radius: 16px; padding: 40px; max-width: 560px; width: 90%; box-shadow: 0 8px 32px rgba(0,0,0,0.08); }
    h1 { font-size: 22px; color: #1e3a8a; margin-bottom: 6px; }
    p.sub { color: #64748b; font-size: 14px; margin-bottom: 28px; }
    .drop-zone { border: 2px dashed #93c5fd; border-radius: 12px; padding: 36px 20px; text-align: center; background: #eff6ff; transition: background 0.2s; cursor: pointer; }
    .drop-zone:hover { background: #dbeafe; }
    .drop-zone p { color: #3b82f6; margin-bottom: 14px; font-size: 15px; }
    input[type=file] { display: none; }
    label.choose { display: inline-block; background: #3b82f6; color: #fff; padding: 10px 22px; border-radius: 8px; cursor: pointer; font-size: 14px; }
    label.choose:hover { background: #2563eb; }
    button#submit-btn { width: 100%; margin-top: 18px; background: #1e3a8a; color: #fff; border: none; padding: 14px; border-radius: 10px; font-size: 16px; cursor: pointer; transition: background 0.2s; }
    button#submit-btn:hover { background: #1e40af; }
    button#submit-btn:disabled { background: #94a3b8; cursor: not-allowed; }
    #selected-files { margin-top: 10px; font-size: 13px; color: #475569; min-height: 20px; }
    #result { margin-top: 24px; display: none; }
    .success { background: #f0fdf4; border: 1px solid #86efac; border-radius: 10px; padding: 20px; }
    .session-id { font-size: 28px; font-weight: bold; color: #16a34a; font-family: monospace; margin: 10px 0; letter-spacing: 2px; }
    .copy-btn { background: #dcfce7; border: 1px solid #86efac; color: #15803d; padding: 6px 14px; border-radius: 6px; cursor: pointer; font-size: 13px; }
    .error { background: #fef2f2; border: 1px solid #fca5a5; border-radius: 10px; padding: 16px; color: #b91c1c; }
    #loading { display: none; text-align: center; color: #64748b; padding: 16px; }
    .spinner { display: inline-block; width: 20px; height: 20px; border: 3px solid #cbd5e1; border-top-color: #3b82f6; border-radius: 50%; animation: spin 0.8s linear infinite; margin-right: 8px; vertical-align: middle; }
    @keyframes spin { to { transform: rotate(360deg); } }
    .step { background: #f8fafc; border-radius: 8px; padding: 12px 16px; margin-top: 14px; font-size: 13px; color: #475569; line-height: 1.6; }
    .step strong { color: #1e3a8a; }
  </style>
</head>
<body>
<div class="card">
  <h1>📄 标书参数抽取工具</h1>
  <p class="sub">上传招标文件，获取 session_id，然后在 BISHENG 对话框中输入即可开始智能抽取</p>

  <form id="uploadForm" enctype="multipart/form-data">
    <div class="drop-zone" id="dropZone">
      <p>支持 PDF / DOCX / DOC，可多文件上传</p>
      <label class="choose" for="fileInput">选择文件</label>
      <input type="file" id="fileInput" name="files" accept=".pdf,.docx,.doc" multiple>
    </div>
    <div id="selected-files">未选择文件</div>
    <button type="submit" id="submit-btn" disabled>🚀 上传并解析</button>
  </form>

  <div id="loading"><span class="spinner"></span>正在解析文件，请稍候（通常需要 5-15 秒）…</div>

  <div id="result"></div>

  <div class="step">
    <strong>使用流程：</strong><br>
    1. 上传招标文件 → 获得 session_id<br>
    2. 打开 BISHENG，输入：<em>「我的 session_id 是 [xxx]，请帮我抽取变压器参数」</em>
  </div>
</div>

<script>
const fileInput = document.getElementById('fileInput');
const submitBtn = document.getElementById('submit-btn');
const selectedFiles = document.getElementById('selected-files');

fileInput.onchange = () => {
  const names = Array.from(fileInput.files).map(f => f.name);
  selectedFiles.textContent = names.length ? names.join(', ') : '未选择文件';
  submitBtn.disabled = names.length === 0;
};

document.getElementById('uploadForm').onsubmit = async (e) => {
  e.preventDefault();
  submitBtn.disabled = true;
  document.getElementById('loading').style.display = 'block';
  document.getElementById('result').style.display = 'none';

  const formData = new FormData(e.target);
  try {
    const res = await fetch('/upload', { method: 'POST', body: formData });
    const data = await res.json();
    document.getElementById('loading').style.display = 'none';
    const div = document.getElementById('result');
    div.style.display = 'block';
    if (data.session_id) {
      div.innerHTML = `<div class="success">
        <p>✅ 上传成功！请将以下 <strong>session_id</strong> 发送给 BISHENG Agent：</p>
        <div class="session-id" id="sid">${data.session_id}</div>
        <button class="copy-btn" onclick="navigator.clipboard.writeText('${data.session_id}').then(()=>this.textContent='✓ 已复制')">复制</button>
        <p style="margin-top:14px;font-size:13px;color:#475569">
          文件：${data.filenames.join(', ')}<br>总页数：${data.total_pages}
        </p>
      </div>`;
    } else {
      div.innerHTML = `<div class="error">❌ ${data.error || '上传失败'}</div>`;
      submitBtn.disabled = false;
    }
  } catch (err) {
    document.getElementById('loading').style.display = 'none';
    document.getElementById('result').style.display = 'block';
    document.getElementById('result').innerHTML = `<div class="error">❌ 网络错误：${err}</div>`;
    submitBtn.disabled = false;
  }
};
</script>
</body>
</html>
"""


# ── Upload HTTP Handler ────────────────────────────────────────────────────────
from starlette.requests import Request
from starlette.responses import HTMLResponse, JSONResponse


async def upload_page(request: Request):
    return HTMLResponse(UPLOAD_HTML)


async def upload_file(request: Request):
    """Receive PDF/DOCX files, parse them, return session_id."""
    form = await request.form()
    files = form.getlist("files")

    if not files:
        return JSONResponse({"error": "未提供文件"}, status_code=400)

    all_chunks, all_headings = [], []
    page_offset = 0
    filenames = []

    for upload in files:
        if not getattr(upload, "filename", None):
            continue
        ext = Path(upload.filename).suffix.lower()
        if ext not in ALLOWED_EXTENSIONS:
            return JSONResponse(
                {"error": f"不支持的文件类型: {upload.filename}，请上传 PDF 或 Word 文件"},
                status_code=400,
            )

        contents = await upload.read()
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
        tmp.write(contents)
        tmp.close()

        try:
            doc_result = await asyncio.to_thread(parse_document, tmp.name)
        finally:
            os.unlink(tmp.name)

        file_chunks = doc_result.get("chunks", [])
        file_headings = doc_result.get("headings", [])
        # Use actual physical page count from app.xml when available (DOCX),
        # otherwise fall back to max physical_page in chunks (PDF)
        file_physical_total = doc_result.get(
            "total_physical_pages",
            max((c.get("physical_page", c["page"]) for c in file_chunks), default=1)
        )

        if not file_chunks:
            continue

        for chunk in file_chunks:
            chunk["global_page"] = chunk["page"] + page_offset
            chunk["source_file"] = upload.filename
            # Also offset physical_page so multi-file searches work correctly
            if "physical_page" in chunk:
                chunk["physical_page"] += page_offset
        for h in file_headings:
            if "page" in h:
                h["page"] += page_offset
            if "physical_page" in h:
                h["physical_page"] += page_offset
            if "chunk_page" in h:
                h["chunk_page"] += page_offset

        page_offset += file_physical_total
        filenames.append(upload.filename)
        all_chunks.extend(file_chunks)
        all_headings.extend(file_headings)

    if not all_chunks:
        return JSONResponse(
            {"error": "无法从文件中提取文本，请检查是否为扫描版 PDF"},
            status_code=400,
        )

    for chunk in all_chunks:
        chunk["page"] = chunk["global_page"]

    total_pages = max((c.get("physical_page", c["page"]) for c in all_chunks), default=0)
    session_id = uuid.uuid4().hex

    with open(os.path.join(tempfile.gettempdir(), f"{session_id}_chunks.json"), "w", encoding="utf-8") as f:
        json.dump(all_chunks, f, ensure_ascii=False)
    with open(os.path.join(tempfile.gettempdir(), f"{session_id}_headings.json"), "w", encoding="utf-8") as f:
        json.dump(all_headings, f, ensure_ascii=False)

    return JSONResponse({"session_id": session_id, "filenames": filenames, "total_pages": total_pages})


# ── Compose Final ASGI App ─────────────────────────────────────────────────────
from starlette.applications import Starlette
from starlette.routing import Route, Mount
from starlette.middleware import Middleware
from starlette.middleware.cors import CORSMiddleware
import uvicorn


def create_app():
    # Get FastMCP's ASGI app (SSE transport: exposes /sse and /messages/)
    mcp_asgi = mcp.http_app(transport="sse")

    combined = Starlette(
        routes=[
            Route("/", upload_page, methods=["GET"]),
            Route("/upload", upload_file, methods=["POST"]),
            Mount("/", app=mcp_asgi),   # handles /sse, /messages/, etc.
        ],
        middleware=[
            Middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"]),
        ],
    )
    return combined


app = create_app()

if __name__ == "__main__":
    port = int(os.getenv("PORT", 8088))
    print(f"🚀 MCP Server starting on http://0.0.0.0:{port}")
    print(f"   Upload page : http://localhost:{port}/")
    print(f"   MCP SSE     : http://localhost:{port}/sse")
    uvicorn.run(app, host="0.0.0.0", port=port, log_level="info")
