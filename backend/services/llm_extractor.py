"""
llm_extractor.py - Two-step LLM extraction using GPT-4o via OneRouter

Step 1: find_section_from_headings - Use document headings + few-shot title map to
                                     locate the target product chapter (generalizable)
Step 2: find_product_section        - Fallback: locate section from doc summary
Step 3: extract_parameters          - Extract specific technical specs from the section

Uses direct HTTP requests for maximum compatibility.
"""
import json
import os
import yaml
import requests
from dotenv import load_dotenv
from pathlib import Path

# Load .env relative to this file's location (backend/services/ -> ../../.env)
_env_path = os.path.join(os.path.dirname(__file__), '..', '..', '.env')
load_dotenv(_env_path, override=True)


def _get_config():
    """每次请求时动态读取配置，避免模块导入时序问题。"""
    return {
        "base_url": os.getenv("OPENROUTER_BASE_URL", "https://openrouter.ai/api/v1").strip(),
        "api_key": os.getenv("OPENROUTER_API_KEY", "").strip(),
        "model": os.getenv("LLM_MODEL", "openai/gpt-4o").strip(),
    }


def _call_llm(messages: list[dict], temperature: float = 0.0) -> dict:
    """
    Call OpenRouter LLM API via direct HTTP requests.
    Returns parsed JSON dict from the model's response.
    Retries up to 3 times on transient network errors (e.g. ConnectionResetError 10054).
    """
    import time

    cfg = _get_config()
    api_key = cfg["api_key"]
    base_url = cfg["base_url"]
    model = cfg["model"]

    if not api_key:
        raise ValueError("OPENROUTER_API_KEY is not set. Check your .env file.")

    print(f"[LLM] Using key: {api_key[:8]}... model: {model}")  # 调试用

    headers = {
        "Authorization": f"Bearer {api_key}",
        "HTTP-Referer": "https://siyuan-poc.local",
        "X-Title": "Siyuan Tender Analyzer",
        "Content-Type": "application/json"
    }

    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": 16000,           # prevent truncation when extracting 97+ params
        "response_format": {"type": "json_object"}
    }

    MAX_RETRIES = 3
    RETRY_DELAYS = [5, 15, 30]  # seconds between retries

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.post(
                f"{base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=180,
                proxies={}  # 绕过系统代理，直连 OpenRouter
            )
            break  # success — exit retry loop
        except (requests.exceptions.ConnectionError,
                requests.exceptions.ChunkedEncodingError,
                ConnectionResetError) as e:
            last_error = e
            if attempt < MAX_RETRIES:
                wait = RETRY_DELAYS[attempt - 1]
                print(f"[LLM] ⚠️ Network error (attempt {attempt}/{MAX_RETRIES}): {e}. "
                      f"Retrying in {wait}s...")
                time.sleep(wait)
            else:
                raise RuntimeError(f"LLM request failed after {MAX_RETRIES} attempts: {e}") from e

    if response.status_code != 200:
        raise RuntimeError(
            f"OpenRouter API error {response.status_code}: {response.text[:500]}"
        )

    data = response.json()

    if "error" in data:
        raise RuntimeError(f"OpenRouter returned error: {data['error']}")

    choice = data["choices"][0]
    finish_reason = choice.get("finish_reason", "unknown")
    if finish_reason == "length":
        print(f"[LLM] ⚠️ RESPONSE TRUNCATED (finish_reason=length) – increase max_tokens or batch params")
    else:
        print(f"[LLM] finish_reason={finish_reason}")

    content = choice["message"]["content"]
    parsed = json.loads(content)
    print(f"[LLM] Response keys: {len(parsed)} params returned")
    return parsed



def find_product_section(doc_summary: str, product_type: str = "transformer") -> dict:
    """
    Step 1: Ask LLM to locate the section in the document that covers the given product.

    Args:
        doc_summary: Compact text with page numbers (from get_doc_summary_for_llm)
        product_type: Product to search for (e.g. "transformer")

    Returns:
        {
            "section_title": str,
            "start_page": int,
            "end_page": int,
            "confidence": float,
            "notes": str
        }
    """
    messages = [
        {
            "role": "system",
            "content": (
                "You are an expert in electrical power engineering tender documents. "
                "Your task is to locate specific product sections within substation construction "
                "tender documents. These tenders typically cover many products: transformers, "
                "circuit breakers, cables, switchgear, etc. "
                "Always respond with valid JSON only, no explanation text."
            )
        },
        {
            "role": "user",
            "content": f"""Below is a summary of an English electrical engineering tender document,
showing the first ~200 characters of each page.

Find the section that covers: **Power Transformer / 变压器**
(may also be called "Transformer", "Power Transformer", "Main Transformer", etc.)

Return a JSON object with this exact format:
{{
  "section_title": "exact section title as it appears in the document",
  "start_page": <first page number of transformer section, integer>,
  "end_page": <last page number of transformer section, integer>,
  "confidence": <0.0 to 1.0, how confident you are>,
  "notes": "any relevant observations"
}}

Document summary (with page numbers):
{doc_summary}"""
        }
    ]

    return _call_llm(messages, temperature=0.1)


# ─── Title map & heading-based section detection ──────────────────────────────

_TITLE_MAP_PATH = Path(__file__).parent.parent / "transformer_title_map.yaml"


def load_title_map() -> dict:
    """Load the product→section-title mapping YAML."""
    try:
        with open(_TITLE_MAP_PATH, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}
    except FileNotFoundError:
        return {}


def save_title_to_map(product_key: str, new_title: str) -> bool:
    """
    Append *new_title* to the known_titles list for *product_key* in the YAML
    mapping file, if it isn't already there.

    Returns True if the file was updated.
    """
    data = load_title_map()
    products = data.get("products", {})
    product = products.get(product_key, {})
    known = product.get("known_titles", [])

    if new_title in known:
        return False  # already known

    known.append(new_title)
    product["known_titles"] = known
    products[product_key] = product
    data["products"] = products

    with open(_TITLE_MAP_PATH, "w", encoding="utf-8") as f:
        yaml.dump(data, f, allow_unicode=True, default_flow_style=False)
    print(f"[TitleMap] Saved new title for '{product_key}': {new_title}")
    return True


def find_section_from_headings(
    headings: list[dict],
    product_type: str = "power_transformer",
) -> dict:
    """
    PRIMARY section detection method.

    Sends the document's heading list to the LLM together with few-shot
    examples from the YAML title map.  The LLM picks the heading that
    most likely corresponds to the target product's technical spec section.

    Args:
        headings: list of {"text": str, "chunk_page": int} from parse_document
        product_type: key in transformer_title_map.yaml (e.g. "power_transformer")

    Returns:
        {
            "section_title": str,
            "start_page": int,      # chunk_page of the matched heading
            "end_page": int,        # chunk_page of the *next* heading - 1
            "confidence": float,
            "notes": str,
        }
    """
    if not headings:
        return _empty_section("No headings found in document")

    # Build heading list string
    heading_lines = "\n".join(
        f"  [{h['chunk_page']}] {h['text']}" for h in headings
    )

    # Load few-shot examples from the title map
    title_map = load_title_map()
    product_info = title_map.get("products", {}).get(product_type, {})
    known_titles = product_info.get("known_titles", [])
    exclude_titles = product_info.get("exclude_titles", [])

    few_shot_block = ""
    if known_titles:
        examples = "\n".join(f'  \u2713 "{t}"' for t in known_titles[:10])
        few_shot_block += f"\n\u5df2\u77e5\u5339\u914d\u7ae0\u8282\u6807\u9898\u793a\u4f8b\uff08\u5386\u53f2\u9879\u76ee\u79ef\u7d2f\uff09\uff1a\n{examples}\n"
    if exclude_titles:
        excludes = "\n".join(f'  \u2717 "{t}"' for t in exclude_titles[:8])
        few_shot_block += f"\n\u4ee5\u4e0b\u6807\u9898\u7c7b\u578b\u5e94\u6392\u9664\uff08\u4e0d\u5339\u914d\uff09\uff1a\n{excludes}\n"

    _PRODUCT_DESCRIPTIONS = {
        "power_transformer": (
            "\u4e3b\u53d8\u538b\u5668\uff08Main/Power Transformer\uff09\u6280\u672f\u89c4\u683c\u7ae0\u8282\u3002"
            "\u6ce8\u610f\uff1a\u53ea\u5339\u914d\u4e3b\u53d8\u538b\u5668\u672c\u8eab\u7684\u89c4\u683c\u7ae0\u8282\uff0c"
            "\u6392\u9664\u8f85\u52a9\u53d8\u538b\u5668\u3001\u4eea\u7528\u53d8\u538b\u5668\uff0c"
            "\u4ee5\u53ca\u5176\u4ed6\u8bbe\u5907\u7ae0\u8282\u4e2d\u987a\u5e26\u63d0\u5230\u53d8\u538b\u5668\u7684\u90e8\u5206\u3002"
        ),
        "site_conditions": (
            "\u73b0\u573a\u8fd0\u884c\u6761\u4ef6/\u73af\u5883\u53c2\u6570\u7ae0\u8282\uff08Site Conditions / Environmental Conditions\uff09\u3002"
            "\u8fd9\u79cd\u7ae0\u8282\u901a\u5e38\u5728\u6807\u4e66\u5f00\u5934\uff0c\u63cf\u8ff0\u5b89\u88c5\u5730\u70b9\u7684\u6c14\u5019\u3001"
            "\u6d77\u62d4\u3001\u6e29\u5ea6\u3001\u6c61\u67d3\u7b49\u7ea7\u7b49\u73af\u5883\u53c2\u6570\u3002"
        ),
    }
    product_desc = _PRODUCT_DESCRIPTIONS.get(
        product_type,
        f"\u4e0e '{product_type}' \u76f8\u5173\u7684\u6280\u672f\u89c4\u683c\u7ae0\u8282\u3002"
    )

    messages = [
        {
            "role": "system",
            "content": (
                "\u4f60\u662f\u7535\u529b\u5de5\u7a0b\u6807\u4e66\u5206\u6790\u4e13\u5bb6\u3002"
                "\u7ed9\u5b9a\u6587\u6863\u7684\u7ae0\u8282\u6807\u9898\u5217\u8868\uff0c\u8bf7\u8bc6\u522b\u54ea\u4e2a\u7ae0\u8282\u5bf9\u5e94\u76ee\u6807\u4ea7\u54c1/\u4e3b\u9898\u7684\u6280\u672f\u89c4\u683c\u5185\u5bb9\u3002"
                "\u8bf7\u53ea\u8fd4\u56de\u5408\u6cd5\u7684 JSON\uff0c\u4e0d\u8981\u8f93\u51fa\u5176\u4ed6\u5185\u5bb9\u3002"
                "notes \u5b57\u6bb5\u5fc5\u987b\u7528\u4e2d\u6587\u586b\u5199\u3002"
            ),
        },
        {
            "role": "user",
            "content": f"""\u4ee5\u4e0b\u662f\u4ece\u82f1\u6587\u7535\u6c14\u5de5\u7a0b\u6807\u4e66\u4e2d\u63d0\u53d6\u7684\u7ae0\u8282\u6807\u9898\u5217\u8868\u3002
\u6bcf\u4e2a\u6807\u9898\u524d\u62ec\u53f7\u4e2d\u7684\u6570\u5b57\u662f\u5176\u5185\u90e8\u5206\u5757\u9875\u7801\u3002
{few_shot_block}
\u6587\u6863\u7ae0\u8282\u5217\u8868\uff1a
{heading_lines}

\u4efb\u52a1\uff1a\u8bf7\u627e\u51fa\u5bf9\u5e94\u4ee5\u4e0b\u5185\u5bb9\u7684\u7ae0\u8282\uff1a{product_desc}

\u8fd4\u56de JSON\uff1a
{{
  "matched_heading": "\u4ece\u4e0a\u65b9\u5217\u8868\u4e2d\u7cbe\u786e\u590d\u5236\u7684\u7ae0\u8282\u6807\u9898\uff0c\u82e5\u65e0\u5339\u914d\u5219\u4e3a null",
  "chunk_page": <\u5339\u914d\u7ae0\u8282\u7684\u6574\u6570\u9875\u7801\uff0c\u82e5\u65e0\u5219\u4e3a null>,
  "confidence": <0.0 \u5230 1.0 \u7684\u7f6e\u4fe1\u5ea6>,
  "notes": "\u7528\u4e2d\u6587\u7b80\u8981\u8bf4\u660e\u5339\u914d\u7406\u7531\u6216\u672a\u5339\u914d\u539f\u56e0"
}}""",
        },
    ]

    result = _call_llm(messages, temperature=0.1)

    matched = result.get("matched_heading")
    chunk_page = result.get("chunk_page")

    if not matched or chunk_page is None:
        return _empty_section(result.get("notes", "LLM could not match any heading"))

    # ── Determine end_page using SAME-OR-HIGHER-LEVEL heading boundary ───────
    matched_idx = next(
        (i for i, h in enumerate(headings)
         if h["chunk_page"] == chunk_page and matched.lower() in h["text"].lower()),
        None,
    )
    if matched_idx is None:
        # fallback: match by page only
        matched_idx = next(
            (i for i, h in enumerate(headings) if h["chunk_page"] == chunk_page),
            None,
        )

    matched_level = headings[matched_idx].get("level", 2) if matched_idx is not None else 2

    # Use physical_page for user-facing output (falls back to chunk_page for PDFs)
    start_physical = chunk_page
    if matched_idx is not None:
        start_physical = headings[matched_idx].get("physical_page", chunk_page)

    end_physical = None
    if matched_idx is not None:
        for h in headings[matched_idx + 1:]:
            h_level = h.get("level", 2)
            h_phys = h.get("physical_page", h["chunk_page"])
            if h_level <= matched_level and h_phys > start_physical:
                end_physical = h_phys
                break

    if end_physical is None:
        end_physical = start_physical + 40

    # Sanity guard
    if end_physical < start_physical:
        end_physical = start_physical + 40

    # Minimum span guard: tech spec sections must span >= 10 physical pages
    MIN_PHYS_SPAN = 10
    if (end_physical - start_physical) < MIN_PHYS_SPAN:
        extended = None
        if matched_idx is not None:
            for h in headings[matched_idx + 1:]:
                h_phys = h.get("physical_page", h["chunk_page"])
                if h_phys >= start_physical + MIN_PHYS_SPAN:
                    extended = h_phys
                    break
        if extended is None:
            extended = start_physical + 60
        print(f"[HeadingSection] Narrow span ({end_physical - start_physical} pages), "
              f"extending end_page {end_physical} → {extended}")
        end_physical = extended

    print(f"[HeadingSection] '{matched}' → physical pages {start_physical}-{end_physical} "
          f"(matched_level=L{matched_level})")

    return {
        "section_title": matched,
        "start_page": start_physical,
        "end_page": end_physical,
        "confidence": result.get("confidence", 0.8),
        "notes": result.get("notes", ""),
    }



def _empty_section(notes: str = "") -> dict:
    return {
        "section_title": "",
        "start_page": None,
        "end_page": None,
        "confidence": 0.0,
        "notes": notes,
    }


def extract_parameters(
    section_text: str,
    param_list: list[str],
    sections_context: list[dict] | None = None,
    max_chars: int = 200_000,
) -> dict:
    """
    Step 2: Extract specific technical parameters from the located section text.

    Args:
        section_text: Full text of the section
        param_list: List of parameter names from the Excel checklist
        sections_context: Optional [{"title": "2.基本参数", "params": [...]}] for LLM context
        max_chars: Maximum chars to send to LLM (default 200K — fits GPT-4o 128K context)

    Returns:
        {"param_name": {"value", "unit", "source_text", "found"}}
    """
    if len(section_text) > max_chars:
        section_text = section_text[:max_chars] + "\n\n[... section truncated ...]"

    # Build param list - use flat format proven to work.
    # If sections_context provided, add section name as a parenthetical note on the
    # first item of each section so LLM has context without changing response format.
    if sections_context:
        param_set = set(param_list)
        name_to_section: dict[str, str] = {}
        for sec in sections_context:
            for i, p in enumerate(sec.get("params", [])):
                if p in param_set:
                    if i == 0:  # first param of section → annotate with section title
                        name_to_section[p] = sec["title"]
        lines = []
        for p in param_list:
            if p in name_to_section:
                lines.append(f"\n### {name_to_section[p]}")
            lines.append(f"- {p}")
        param_list_str = "\n".join(lines).strip()
    else:
        param_list_str = "\n".join(f"- {p}" for p in param_list)

    messages = [
        {
            "role": "system",
            "content": (
                "You are a transformer technical specifications expert who is BILINGUAL in Chinese and English. "
                "You extract technical parameter values from English tender documents for POWER TRANSFORMERS "
                "(大型电力变压器, typically 110kV/220kV/500kV, rated capacity 40MVA+). "
                "IMPORTANT: The section text may contain multiple transformer types. "
                "Extract values ONLY for the POWER TRANSFORMER (not earthing transformer, "
                "not auxiliary transformer, not instrument transformer). "
                "Parameter names are labeled in Chinese. Use your bilingual knowledge to match them to "
                "their English equivalents in the document. "
                "Extract values accurately. If a parameter is genuinely not mentioned, set 'found' to false. "
                "Always respond with valid JSON only."
            )
        },
        {
            "role": "user",
            "content": f"""Extract the following technical parameters from this transformer specification text.

IMPORTANT NOTES:
1. Parameter names are in CHINESE; the document is in ENGLISH. Complete bilingual reference:
   Site conditions (S1):
     '运行场所'=installation site/location, '最高温度'=maximum ambient temperature,
     '年平均温度'=annual average temperature, '最大风速'=maximum wind speed,
     '污染程度'=pollution level/creepage class, '海拔高度'=altitude/elevation,
     '最低温度'=minimum temperature, '月平均温度'=monthly average temperature,
     '最大日照强度'=maximum solar radiation intensity, '环境腐蚀等级'=corrosion category.
   Basic parameters (S2):
     '执行标准'=applicable standard, '频率'=frequency, '噪声水平'=noise level/sound pressure,
     '额定容量'=rated power/capacity, '联结组别'=vector group, '开关类型'=tap changer type,
     '相数'=number of phases, '冷却方式'=cooling method, '额定电压'=rated voltage,
     '绝缘水平'=insulation level, '系统短路电流'=system short-circuit current,
     '短路持续时间'=short-circuit duration, '局部放电'=partial discharge,
     '阻抗电压'=impedance voltage/leakage reactance, '空载损耗'=no-load loss/iron loss,
     '负载损耗'=load loss/copper loss, '效率'=efficiency,
     '顶层油温升'=top oil temperature rise, '绕组平均温升'=winding average temperature rise,
     '绕组热点温升'=winding hot-spot temperature rise, '铁心温升'=core temperature rise.
   Iron core (S3):
     '型式 [3.1]'=core type (core/shell), '磁通密度'=flux density,
     '激磁电流'=magnetizing/excitation current, '铁心装配'=core assembly requirements,
     '铁心绝缘耐热等级'=core insulation thermal class, '谐波要求'=harmonic requirement.
   Winding (S4):
     '电流密度'=current density, '绝缘纸类型'=insulation paper type.
   Protection accessories (S7):
     '绕组温度计'=winding thermometer/temperature gauge, '油温计'=oil thermometer,
     '气体继电器'=buchholz relay, '油位计(OLTC)'=OLTC oil level indicator,
     '突发压力继电器 [7.5]'=sudden pressure relay/SPR (main tank),
     '呼吸器(本体)'=main tank breather/silica gel breather,
     '压力释放阀'=pressure relief valve/PRV,
     '突发压力继电器 [7.8]'=sudden pressure relay/SPR (OLTC),
     '油位计(本体)'=main tank oil level indicator, '呼吸器(OLTC)'=OLTC breather.
   Online monitoring (S8):
     '光纤测温'=fiber optic temperature monitoring,
     '气体在线监测'=dissolved gas analysis/DGA online monitor,
     '充氮灭火'=nitrogen injection/fire suppression,
     '套管在线监测'=bushing online monitoring, '局放监测'=partial discharge monitor,
     '胶囊泄露监测'=bladder/diaphragm leak monitoring,
     '信号集中器'=signal concentrator/junction box,
     '综合类在线监测'=comprehensive online monitoring system,
     '变压器在线监测系统'=transformer online monitoring system.
   Cooling equipment (S9):
     '片散'=radiator/fin radiator/corrugated fin,
     '涂漆(热镀锌)'=painting/hot-dip galvanizing,
     '片散中间片涂漆'=intermediate fin painting,
     '冷却装置挂本体/分体'=cooling unit attached-to-tank/separate,
     '风扇'=fan/forced air cooling, '油泵'=oil pump/forced oil,
     '冷却器'=cooler/heat exchanger.
   Transport (S10):
     '运输方式'=transport method/mode, '变压器运输类型'=transformer transport type,
     '冲撞记录仪'=impact recorder/shock recorder, 'GPS'=GPS tracking,
     '运输尺寸及重量限制'=transport dimension and weight limit,
     '安装尺寸及重量限制'=installation dimension and weight limit.
   Insulating oil (S11):
     '油类型'=oil type/insulating oil type, '油厂家'=oil manufacturer/supplier,
     '备用油'=spare oil/reserve oil, '油运输方式'=oil transport method.
   Tank (S12):
     '油箱压力强度'=tank pressure strength/pressure test,
     '油箱真空强度'=tank vacuum strength/vacuum test,
     '防腐设计分类/涂层规格'=corrosion protection/anti-corrosion coating specification,
     '箱盖固定方式'=tank cover/lid fixing method,
     '焊缝测试'=weld test/seam test, '紧固件'=fasteners/bolts,
     '密封件'=sealing/gasket, '小车'=trolley/rollers/skids,
     '减震垫'=anti-vibration pad/shock absorber pad,
     '是否油箱防爆要求'=tank explosion-proof requirement,
     '电缆罩'=cable cover/cable box,
     '本体储油柜'=main tank conservator, '开关储油柜'=OLTC conservator.
   Conservator (S13):
     '型式 [13.1]'=conservator type (open/sealed/bladder),
     '主油箱吸湿器型式要求'=main tank dehydrating breather type.
   Terminal box (S14):
     '外壳材质'=enclosure material, '辅助电源要求'=auxiliary power supply,
     '是否带漏电保护'=residual current protection/earth leakage protection,
     '二次电缆'=secondary/control cable, '端子要求'=terminal requirements.
   Other accessories (S15):
     '中性点接地装置要求'=neutral grounding device,
     '中性点接地电阻要求'=neutral grounding resistance,
     '铁心接地电流监测'=core earth current monitoring,
     '智能组件柜'=intelligent terminal unit/IED cabinet,
     '防跌落系统'=anti-falling/anti-seismic system,
     '热虹吸过滤器'=thermosiphon filter,
     '火灾探测器'=fire detector.
   Misc (S16-18):
     '备品备件'=spare parts, '试验要求'=test requirements/type tests,
     '电流互感器'=current transformer/CT.
2. Suffixes like [3.1] or [13.1] in parameter names are Excel checklist item numbers for disambiguation.
   IGNORE THESE SUFFIXES when searching - search for the base technical term only.
   e.g., '型式 [3.1]' → search for core TYPE; '型式 [13.1]' → search for conservator TYPE.
3. The document uses a numbered table format: ItemNum | Description | Unit | RequiredValue
   e.g. "4.2 | Shell or core | Core" means core type = Core.
   "5.1 | First stage | ONAN" means first cooling stage = ONAN.
4. If a field says "Should be Proposed By Manufacturer/Tenderer" or is blank, set found=false.
5. The ### headings below are section labels for context only — do NOT include them as JSON keys.

Parameters to extract:
{param_list_str}

For each "-" bullet parameter, return:
- "value": extracted value (string), or null if not found
- "unit": unit of measurement (e.g. "kV", "MVA", "%", "°C"), or "" if none
- "source_text": first 50 characters of the source sentence/row (truncate after 50 chars), or "" if not found
- "found": true if value was found, false otherwise

Return a JSON object where EACH KEY is the EXACT PARAMETER NAME as listed after "-" bullets above
(including any [X.X] suffix — copy the key name exactly as written).

Transformer section text:
{section_text}"""
        }
    ]

    result = _call_llm(messages, temperature=0.0)

    # Truncation detection: if LLM returned fewer keys than expected, some params were silently dropped
    if len(result) < len(param_list):
        missing = len(param_list) - len(result)
        print(f"[extract_parameters] ⚠️  TRUNCATION DETECTED: sent {len(param_list)} params, "
              f"got {len(result)} back ({missing} missing). "
              f"Consider splitting into smaller batches or reducing source_text length.")
    return result


def get_param_list_from_template(template_info: dict) -> list[str]:
    """Extract parameter names from Excel template info."""
    return [p["name"] for p in template_info["params"]]
