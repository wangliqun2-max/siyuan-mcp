"""
app.py - Flask backend for Siyuan Tender Parameter Extraction Tool
"""
import os
import json
import tempfile
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from dotenv import load_dotenv
import io

# Load environment variables from .env in parent directory
load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))

from services.doc_parser import (parse_document, get_doc_summary_for_llm,
                                  get_section_text, keyword_find_section,
                                  keyword_density_find_section,
                                  find_section_in_headings,
                                  SITE_CONDITIONS_HEADING_KEYWORDS)
from services.llm_extractor import (
    find_product_section, find_section_from_headings,
    extract_parameters, get_param_list_from_template, save_title_to_map,
)
from services.excel_handler import read_template_params, write_results_to_excel

app = Flask(__name__)
CORS(app)

# Flask 3.x sort_keys defaults to True, which destroys our carefully ordered
# param dict. Disable it so jsonify preserves dict insertion order.
app.json.sort_keys = False

# Template path
TEMPLATE_PATH = os.path.join(
    os.path.dirname(__file__), '..', 
    os.getenv("EXCEL_TEMPLATE_PATH", "Checklist模板.xlsx")
)

ALLOWED_EXTENSIONS = {'.pdf', '.docx', '.doc'}


def allowed_file(filename: str) -> bool:
    return os.path.splitext(filename)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({"status": "ok", "message": "Siyuan POC backend running"})


@app.route('/api/template-params', methods=['GET'])
def get_template_params():
    """Return the list of parameters from the Excel template."""
    try:
        template_info = read_template_params(TEMPLATE_PATH)
        param_names = [p["name"] for p in template_info["params"]]
        return jsonify({
            "success": True,
            "sheet_name": template_info["sheet_name"],
            "param_count": len(param_names),
            "params": param_names
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500


@app.route('/api/locate-section', methods=['POST'])
def locate_section():
    """
    Step 1: Upload document, AI locates the transformer section.
    
    Form data:
        file: PDF or Word document
        product_type: e.g. "transformer" (optional, defaults to transformer)
    
    Returns:
        {
            "success": bool,
            "session_id": str,       # temp file reference for step 2
            "section": {
                "section_title": str,
                "start_page": int,
                "end_page": int,
                "confidence": float,
                "notes": str
            },
            "total_pages": int
        }
    """
    # Accept both 'files' (multi) and legacy 'file' (single) field names
    uploaded = request.files.getlist('files') or request.files.getlist('file')
    if not uploaded or all(f.filename == '' for f in uploaded):
        return jsonify({"success": False, "error": "No file provided"}), 400

    invalid = [f.filename for f in uploaded if not allowed_file(f.filename)]
    if invalid:
        return jsonify({
            "success": False,
            "error": f"Unsupported file type(s): {', '.join(invalid)}. Please upload PDF or Word (.docx) files."
        }), 400

    product_type = request.form.get('product_type', 'transformer')

    try:
        all_chunks: list[dict] = []
        all_headings: list[dict] = []
        file_map: dict[int, str] = {}   # global_page → original filename
        file_page_ranges: list[dict] = []  # for response info
        page_offset = 0

        for file in uploaded:
            ext = os.path.splitext(file.filename)[1].lower()
            tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=ext, dir=tempfile.gettempdir())
            file.save(tmp_file.name)
            tmp_file.close()

            try:
                doc_result = parse_document(tmp_file.name)
            finally:
                os.unlink(tmp_file.name)

            file_chunks   = doc_result.get("chunks", [])
            file_headings = doc_result.get("headings", [])
            # Use actual physical page count when available (DOCX reads from app.xml)
            file_physical_total = doc_result.get(
                "total_physical_pages",
                max((c.get("physical_page", c["page"]) for c in file_chunks), default=1)
            )

            if not file_chunks:
                print(f"[Upload] Warning: no text extracted from {file.filename}")
                continue

            # Re-number pages with offset; tag each chunk with source file
            first_global = page_offset + 1
            for chunk in file_chunks:
                g_page = chunk["page"] + page_offset
                chunk["global_page"] = g_page
                chunk["source_file"] = file.filename
                file_map[g_page] = file.filename
            for h in file_headings:
                if "page" in h:
                    h["page"] = h["page"] + page_offset

            last_page_in_file = file_physical_total
            last_global = last_page_in_file + page_offset
            file_page_ranges.append({
                "filename": file.filename,
                "start_page": first_global,
                "end_page": last_global,
                "pages": last_page_in_file
            })
            print(f"[Upload] {file.filename}: {file_physical_total} physical pages → global {first_global}-{last_global}")

            all_chunks.extend(file_chunks)
            all_headings.extend(file_headings)
            page_offset = last_global  # next file starts after this one

        if not all_chunks:
            return jsonify({"success": False, "error": "Could not extract text from any of the uploaded files. Are they scanned/image PDFs?"}), 400

        total_pages = sum(fr["pages"] for fr in file_page_ranges)

        # Use global_page as the "page" field that the rest of the pipeline uses
        for chunk in all_chunks:
            chunk["page"] = chunk["global_page"]
        section_result = None

        # ==== Step 1: Heading-based LLM detection =============================
        # Works for BOTH DOCX (Word styles) and PDF (regex-extracted headings).
        # end_page = next_heading_page - 1, so chapter boundaries are exact.
        # e.g. "14  Power Transformers" starts p307, next "15  Earthing" at p334
        #  → section is p307-p333.
        if all_headings:
            print(f"[Section] Found {len(all_headings)} headings. Running LLM heading detection...")
            try:
                section_result = find_section_from_headings(all_headings, product_type="power_transformer")
                if section_result.get("start_page") is not None:
                    print(f"[Section] ✓ Heading match: '{section_result['section_title']}' "
                          f"pages {section_result['start_page']}-{section_result['end_page']} "
                          f"(confidence={section_result.get('confidence', 0):.2f})")
                else:
                    print("[Section] Heading LLM returned no match.")
                    section_result = None
            except Exception as e:
                print(f"[Section] Heading detection error: {e}")
                section_result = None
        else:
            print("[Section] No headings extracted from document.")

        # ==== Step 2: Keyword-density fallback (no LLM cost) ==================
        # Used when no headings were extracted (e.g. heavily image-based PDF).
        # Does NOT use chapter headings for boundary detection, so end_page may
        # be slightly imprecise (+/- a few pages).
        if section_result is None:
            print(f"[Section] Falling back to keyword density detection across {len(all_chunks)} pages...")
            density_result = keyword_density_find_section(all_chunks)
            if density_result.get("start_page") is not None:
                section_result = density_result
                print(f"[Section] ✓ Density match: '{section_result['section_title']}' "
                      f"pages {section_result['start_page']}-{section_result['end_page']}")

        # ==== Step 3: Full-doc LLM summary (last resort) ======================
        if section_result is None:
            print("[Section] Last resort: full-doc LLM summary detection...")
            TRANSFORMER_KEYWORDS = ["POWER TRANSFORMER", "132/33KV", "MAIN TRANSFORMER"]
            doc_summary = get_doc_summary_for_llm(
                all_chunks, max_chars=40000, product_keywords=TRANSFORMER_KEYWORDS
            )
            try:
                section_result = find_product_section(doc_summary, product_type)
            except Exception as e:
                print(f"[Section] LLM fallback error: {e}")
                section_result = {"section_title": "", "start_page": None,
                                   "end_page": None, "confidence": 0.0, "notes": str(e)}

        # Store session
        import uuid
        session_id = uuid.uuid4().hex
        session_file = os.path.join(tempfile.gettempdir(), f"{session_id}_chunks.json")
        with open(session_file, 'w', encoding='utf-8') as f:
            json.dump(all_chunks, f, ensure_ascii=False)

        headings_file = os.path.join(tempfile.gettempdir(), f"{session_id}_headings.json")
        with open(headings_file, 'w', encoding='utf-8') as f:
            json.dump(all_headings, f, ensure_ascii=False)

        file_map_file = os.path.join(tempfile.gettempdir(), f"{session_id}_filemap.json")
        with open(file_map_file, 'w', encoding='utf-8') as f:
            json.dump(file_map, f, ensure_ascii=False)

        return jsonify({
            "success": True,
            "session_id": session_id,
            "section": section_result,
            "total_pages": total_pages,
            "filenames": [f["filename"] for f in file_page_ranges],
            "file_page_ranges": file_page_ranges,
            "filename": file_page_ranges[0]["filename"] if file_page_ranges else ""
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"success": False, "error": f"Error processing document: {str(e)}"}), 500


@app.route('/api/confirm-section', methods=['POST'])
def confirm_section():
    """
    Called when the user confirms a section finding (clicks 确认章节).
    Saves the confirmed section title to the title map YAML for future use.

    JSON body:
        {
            "section_title": str,   # confirmed title to save
            "product_type": str     # e.g. "power_transformer"
        }
    """
    data = request.get_json(silent=True) or {}
    title = (data.get("section_title") or "").strip()
    product_type = (data.get("product_type") or "power_transformer").strip()

    if not title:
        return jsonify({"success": False, "error": "No section_title provided"}), 400

    updated = save_title_to_map(product_type, title)
    return jsonify({
        "success": True,
        "saved": updated,
        "message": f"Title {'saved to map' if updated else 'already known'}: {title}"
    })


@app.route('/api/extract-params', methods=['POST'])
def extract_params():
    """
    Step 2: Extract parameters from the located section.
    
    JSON body:
        {
            "session_id": str,
            "start_page": int,
            "end_page": int,
            "product_type": str
        }
    
    Returns:
        {
            "success": bool,
            "extracted": {
                "param_name": {
                    "value": str,
                    "unit": str,
                    "source_text": str,
                    "found": bool
                }
            },
            "stats": {
                "total": int,
                "found": int,
                "not_found": int
            }
        }
    """
    data = request.get_json()
    session_id = data.get('session_id')
    start_page = int(data.get('start_page', 1))
    end_page = int(data.get('end_page', 999))

    # Load cached chunks
    session_file = os.path.join(tempfile.gettempdir(), f"{session_id}_chunks.json")
    if not os.path.exists(session_file):
        return jsonify({"success": False, "error": "Session expired. Please re-upload the document."}), 404
    
    try:
        with open(session_file, 'r', encoding='utf-8') as f:
            chunks = json.load(f)

        # Get parameter list from template (preserves section_num per param)
        template_info = read_template_params(TEMPLATE_PATH)
        all_params = template_info["params"]
        param_list = get_param_list_from_template(template_info)
        sections_ctx = template_info.get("sections", [])  # for LLM context & frontend grouping

        # ── Split params by section ─────────────────────────────────────────
        site_params  = [p["name"] for p in all_params if p.get("section_num") == 1]
        equip_params = [p["name"] for p in all_params if p.get("section_num", 1) != 1]
        site_sections  = [s for s in sections_ctx if s["params"] and s["params"][0] in site_params]
        equip_sections = [s for s in sections_ctx if s["params"] and s["params"][0] in equip_params]

        print(f"[Extract] site_params={len(site_params)} equip_params={len(equip_params)}")

        extracted_all: dict = {}

        # Total pages in session
        total_pages = max(c["page"] for c in chunks) if chunks else end_page

        # NOTE: Do NOT extend end_page beyond the confirmed section boundary.
        # The heading-based locator finds pages 307-334 (Power Transformer chapter).
        # Extending into pages 335+ would include Chapter 15 (Earthing Transformer),
        # a different product, which contaminates the extraction and causes the LLM
        # to return found=false for cooling/transport/oil/tank params (sections 9-12).
        print(f"[Extract] Using transformer section: pages {start_page}-{end_page} "
              f"(of {total_pages} total)")


        CORE_SECTIONS = {2, 3, 4}
        core_params  = [p["name"] for p in all_params if p.get("section_num") in CORE_SECTIONS]
        extra_params = [p["name"] for p in all_params if p.get("section_num", 0) not in CORE_SECTIONS
                        and p.get("section_num", 0) != 1]
        core_set  = set(core_params)
        extra_set = set(extra_params)
        core_sections  = [s for s in equip_sections if any(p in core_set  for p in s["params"])]
        extra_sections = [s for s in equip_sections if any(p in extra_set for p in s["params"])]

        # Group chunks by source_file so each file is searched independently.
        # This avoids the char-limit window missing Part-3 content when Part-2 is prepended.
        from collections import defaultdict
        file_chunks_map: dict = defaultdict(list)
        for c in chunks:
            file_chunks_map[c.get("source_file", "__single__")].append(c)

        # Merge helper: found=true wins over found=false across files
        def _merge(base: dict, new_r: dict) -> dict:
            for k, v in new_r.items():
                if k not in base or (isinstance(v, dict) and v.get("found") and not base.get(k, {}).get("found")):
                    base[k] = v
            return base

        # ── 2a: Core electrical params ─────────────────────────────────────
        if core_params:
            for src_file, fchunks in file_chunks_map.items():
                core_text = get_section_text(fchunks, start_page, end_page)
                if not core_text.strip():
                    print(f"[Core] '{src_file}': no text in pages {start_page}-{end_page}, skip")
                    continue
                print(f"[Core] '{src_file}': pages {start_page}-{end_page}, {len(core_text)} chars")
                core_result = extract_parameters(core_text, core_params,
                                                 sections_context=core_sections, max_chars=200_000)
                found_core = [k for k,v in core_result.items() if isinstance(v,dict) and v.get("found")]
                print(f"[Core] found {len(found_core)}: {found_core[:10]}")
                _merge(extracted_all, core_result)

        # ── 2b: Extra params (sections 7-18) — single call ───────────────────
        # source_text is now capped at 50 chars in the prompt, so per-param output
        # is ~30 tokens. 65 params × 30 = ~1950 output tokens → safe within 16K.
        if extra_params:
            for src_file, fchunks in file_chunks_map.items():
                extra_text = get_section_text(fchunks, start_page, end_page)
                if not extra_text.strip():
                    print(f"[Extra] '{src_file}': no text in pages {start_page}-{end_page}, skip")
                    continue
                print(f"[Extra] '{src_file}': pages {start_page}-{end_page}, "
                      f"{len(extra_text)} chars, {len(extra_params)} params")
                extra_result = extract_parameters(extra_text, extra_params,
                                                  sections_context=extra_sections,
                                                  max_chars=200_000)
                found_extra = [k for k, v in extra_result.items()
                               if isinstance(v, dict) and v.get("found")]
                print(f"[Extra] found {len(found_extra)}/{len(extra_params)}: {found_extra[:10]}")
                _merge(extracted_all, extra_result)


        # ── Site params → heading-based section lookup (zero LLM cost) ──────
        # Load session headings, keyword-search for "Service Conditions" /
        # "Site Conditions" / "Ambient Conditions" etc., then use its page range.
        # end_page = next same-or-higher-level heading's page (no -1 per spec).
        # Falls back to pages 1-50 only when no heading keyword matches.
        if site_params:
            headings_file = os.path.join(tempfile.gettempdir(), f"{session_id}_headings.json")
            saved_headings = []
            if os.path.exists(headings_file):
                with open(headings_file, 'r', encoding='utf-8') as f:
                    saved_headings = json.load(f)

            site_section = find_section_in_headings(saved_headings,
                                                     SITE_CONDITIONS_HEADING_KEYWORDS)
            if site_section.get("start_page"):
                site_start = site_section["start_page"]
                site_end   = site_section["end_page"]
                print(f"[Site] Heading match: '{site_section['section_title']}' "
                      f"pages {site_start}-{site_end}")
                site_text = get_section_text(chunks, site_start, site_end)
            else:
                print("[Site] No matching heading — fallback to pages 1-50")
                site_text = get_section_text(chunks, 1, 50)

            if site_text.strip():
                site_result = extract_parameters(site_text, site_params,
                                                 sections_context=site_sections)
                found_site = [k for k, v in site_result.items()
                              if isinstance(v, dict) and v.get("found")]
                print(f"[Site] found {len(found_site)}/{len(site_params)}: {found_site}")
                _merge(extracted_all, site_result)  # use merge so found=True always wins


        # ── Reorder to match Excel template order ────────────────────────────
        extracted_ordered = {}
        for param_name in param_list:
            if param_name in extracted_all:
                extracted_ordered[param_name] = extracted_all[param_name]
            else:
                extracted_ordered[param_name] = {"value": None, "unit": "", "source_text": "", "found": False}

        found_count = sum(1 for v in extracted_ordered.values() if isinstance(v, dict) and v.get("found"))

        # ── Debug: print breakdown of found vs not-found per section ────────
        print(f"\n[Extract] RESULT SUMMARY: {found_count}/{len(param_list)} found")
        found_names = [k for k, v in extracted_ordered.items() if isinstance(v, dict) and v.get("found")]
        not_found_names = [k for k, v in extracted_ordered.items() if isinstance(v, dict) and not v.get("found")]
        print(f"[Extract] FOUND ({len(found_names)}): {found_names[:20]}")
        if len(found_names) > 20:
            print(f"[Extract]   ... and {len(found_names)-20} more")
        print(f"[Extract] NOT-FOUND ({len(not_found_names)}): {not_found_names[:10]}")
        if len(not_found_names) > 10:
            print(f"[Extract]   ... and {len(not_found_names)-10} more")

        # Build all_chunks dict for the right-panel document viewer
        # Key = page str, value = {text, source_file}
        all_chunks_dict = {
            str(c.get("page", i)): {
                "text": c.get("text", ""),
                "source_file": c.get("source_file", "")
            }
            for i, c in enumerate(chunks)
        }

        return jsonify({
            "success": True,
            "extracted": extracted_ordered,
            "sections": sections_ctx,
            "all_chunks": all_chunks_dict,
            "stats": {
                "total": len(param_list),
                "found": found_count,
                "not_found": len(param_list) - found_count,
            },
        })

    except Exception as e:
        return jsonify({"success": False, "error": f"Extraction error: {str(e)}"}), 500


@app.route('/api/export-csv', methods=['POST'])
@app.route('/api/export-excel', methods=['POST'])   # keep old URL working
def export_csv():
    """
    Generate and return a CSV of extracted parameters.
    Columns: 参数名, 找到, 値, 单位, 来源原文
    UTF-8 BOM so Windows Excel opens Chinese characters correctly.
    """
    import csv as _csv
    import io as _io
    data = request.get_json()
    extracted = data.get('extracted', {})

    buf = _io.StringIO()
    writer = _csv.writer(buf, quoting=_csv.QUOTE_ALL)
    writer.writerow(['参数名', '找到', '値', '单位', '来源原文'])
    for param_name, info in extracted.items():
        if not isinstance(info, dict):
            continue
        writer.writerow([
            param_name,
            '是' if info.get('found') else '否',
            info.get('value') or '',
            info.get('unit') or '',
            (info.get('source_text') or '').replace('\n', ' '),
        ])

    # UTF-8 BOM prefix
    csv_bytes = ('\ufeff' + buf.getvalue()).encode('utf-8')
    today = __import__('datetime').date.today().isoformat()
    return send_file(
        io.BytesIO(csv_bytes),
        mimetype='text/csv; charset=utf-8',
        as_attachment=True,
        download_name=f'变压器参数提取结果_{today}.csv'
    )



if __name__ == '__main__':
    print("🚀 Siyuan POC Backend starting...")
    print(f"📄 Excel template: {TEMPLATE_PATH}")
    print(f"🤖 LLM Model: {os.getenv('LLM_MODEL', 'openai/gpt-4o')}")
    app.run(debug=True, host='0.0.0.0', port=5001)
