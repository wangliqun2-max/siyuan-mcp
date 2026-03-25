"""
doc_parser.py - Parse PDF and Word documents into page-based text chunks.

Key features:
- Extracts both paragraph text AND table content (crucial for spec documents)
- Tracks headings by Word paragraph style for structural navigation
- Provides heading-based and keyword-based section finding utilities
"""
import pdfplumber
from docx import Document
import os


# ─── Heading style name prefixes to recognise ────────────────────────────────
_HEADING_STYLE_PREFIXES = ("heading", "title", "h1", "h2", "h3")


def _is_heading_style(style_name: str) -> bool:
    """Return True if the paragraph style looks like a heading."""
    name_lower = style_name.lower()
    return any(name_lower.startswith(p) for p in _HEADING_STYLE_PREFIXES)


# ─── Public API ───────────────────────────────────────────────────────────────

def parse_document(file_path: str) -> dict:
    """
    Parse a PDF or Word document.

    Returns:
        {
            "chunks":   [{"page": int, "text": str}, ...],
            "headings": [{"text": str, "chunk_page": int}, ...],  # DOCX only
        }
    Callers that previously expected only a list of chunks can access
    result["chunks"] for backward compatibility.
    """
    ext = os.path.splitext(file_path)[1].lower()

    if ext == ".pdf":
        return _parse_pdf(file_path)   # now returns {chunks, headings}
    elif ext in (".docx", ".doc"):
        return _parse_docx(file_path)
    else:
        raise ValueError(
            f"Unsupported file type: {ext}. Please upload PDF or Word (.docx) files."
        )


def get_doc_summary_for_llm(
    chunks: list[dict],
    max_chars: int = 40000,
    product_keywords: list[str] | None = None,
) -> str:
    """
    Build a compact summary of the document for LLM section-finding.

    Chunks that match any of *product_keywords* are marked with ★ and shown
    with a longer preview (500 chars), giving the LLM clear signals.
    """
    kw_upper = [k.upper() for k in (product_keywords or [])]
    lines: list[str] = []
    total_chars = 0

    for chunk in chunks:
        text_upper = chunk["text"].upper()
        is_match = kw_upper and any(kw in text_upper for kw in kw_upper)

        if is_match:
            preview = chunk["text"][:500].replace("\n", " ")
            line = f"[Page {chunk['page']} ★MATCH] {preview}"
        else:
            preview = chunk["text"][:120].replace("\n", " ")
            line = f"[Page {chunk['page']}] {preview}"

        lines.append(line)
        total_chars += len(line)

        if total_chars >= max_chars:
            lines.append("... [document continues]")
            break

    return "\n".join(lines)


def keyword_find_section(
    chunks: list[dict],
    product_keywords: list[str],
    context_chunks: int = 2,
) -> dict:
    """
    Pure-Python keyword-based section finder (fast, zero LLM cost).

    Scans ALL chunks' full text and returns the page range of the longest
    contiguous run of matching chunks.
    """
    match_pages: list[int] = []
    for chunk in chunks:
        text_upper = chunk["text"].upper()
        if any(kw.upper() in text_upper for kw in product_keywords):
            match_pages.append(chunk["page"])

    if not match_pages:
        return {
            "start_page": None, "end_page": None,
            "section_title": "", "confidence": 0.0,
        }

    # Find the longest contiguous run
    runs: list[list[int]] = []
    run = [match_pages[0]]
    for p in match_pages[1:]:
        if p <= run[-1] + context_chunks + 1:
            run.append(p)
        else:
            runs.append(run)
            run = [p]
    runs.append(run)
    best_run = max(runs, key=len)

    # Try to extract a human-readable section title from the first matching chunk
    first_match_chunk = next(c for c in chunks if c["page"] == best_run[0])
    section_title = _extract_title_from_chunk(first_match_chunk["text"], product_keywords)

    return {
        "start_page": best_run[0],
        "end_page": best_run[-1],
        "section_title": section_title or "Power Transformer / 变压器",
        "confidence": min(0.9, 0.5 + len(best_run) * 0.05),
    }


# Spec-specific keywords for Power Transformer technical requirement chapters.
# Covers all checklist sections: electrical params (2-4), protection (7-8),
# cooling (9), transport (10), oil (11), tank (12), conservation (13),
# terminal box (14), accessories (15-18).
_TRANSFORMER_SPEC_KEYWORDS = [
    # Core electrical specs (sections 2-4)
    "winding", "impedance", "no-load loss", "load loss", "no load loss",
    "magnetizing", "flux density", "hotspot", "hot spot", "temperature rise",
    "kraft paper", "silicon steel", "copper conductor", "winding resistance",
    "short circuit impedance", "partial discharge",
    # Protection & monitoring (sections 7-8)
    "buchholz", "conservator", "silica gel", "sudden pressure",
    "dissolved gas", "impact recorder",
    # Cooling & tank (sections 9-12)
    "oltc", "tap changer", "bushing", "radiator", "insulating oil",
    "tank pressure", "core loss", "pressure relief",
    "neutral earthing", "neutral grounding",
    # Unique transformer chapter terms
    "oil preservation", "oil conservator", "corrugated tank",
    "thermometer pocket", "oil thermometer", "winding thermometer",
    "nitrogen injection", "sudden pressure relay",
    "on-load tap", "voltage ratio", "vector group",
    "temperature rise test", "impulse test", "dielectric test",
    "routine test",
]

# A page scores if it contains this many DISTINCT spec keywords
_MIN_SPEC_SCORE = 2

# Skip the first N pages so scope-of-work / general sections are excluded
# before we start looking for the dense spec chapter.
_PAGE_SKIP = 100


def keyword_density_find_section(
    chunks: list[dict],
    spec_keywords: list[str] | None = None,
) -> dict:
    """
    Find the product technical-specification section using keyword density.

    Algorithm — "peak + expand":
    1. Skip the first _PAGE_SKIP pages (scope / general sections).
    2. Score every remaining page by distinct spec keyword count.
    3. Find the highest-scoring page (the "peak") — almost certainly inside
       the spec chapter.
    4. Expand outward from the peak, stopping when score drops to 0 (empty
       of all spec keywords) — that marks the chapter boundary.

    Returns the same dict shape as find_product_section.
    """
    kws = [k.lower() for k in (spec_keywords or _TRANSFORMER_SPEC_KEYWORDS)]

    # Score every page
    page_score: dict[int, int] = {}
    for chunk in chunks:
        text_lower = chunk["text"].lower()
        page_score[chunk["page"]] = sum(1 for kw in kws if kw in text_lower)

    all_pages = sorted(page_score.keys())
    if not all_pages:
        return {"start_page": None, "end_page": None,
                "section_title": "", "confidence": 0.0}

    # ── Step 1: skip first _PAGE_SKIP pages ─────────────────────────────────
    candidate_pages = [p for p in all_pages if p > _PAGE_SKIP]
    if not candidate_pages:
        candidate_pages = all_pages   # very short document fallback

    # ── Step 2: find the peak page ───────────────────────────────────────────
    peak_page = max(candidate_pages, key=lambda p: page_score[p])
    peak_score = page_score[peak_page]

    if peak_score < _MIN_SPEC_SCORE:
        print(f"[DensityFinder] Peak score {peak_score} too low — no spec section found")
        return {"start_page": None, "end_page": None,
                "section_title": "", "confidence": 0.0}

    print(f"[DensityFinder] Peak: page {peak_page} score={peak_score}")

    # ── Step 3: expand left until a score-0 page is hit ─────────────────────
    page_set = set(all_pages)
    start_p = peak_page
    p = peak_page - 1
    while p >= min(all_pages) and p in page_set:
        if page_score[p] == 0:
            break
        start_p = p
        p -= 1

    # ── Step 4: expand right until a score-0 page is hit ────────────────────
    end_p = peak_page
    p = peak_page + 1
    while p <= max(all_pages) and p in page_set:
        if page_score[p] == 0:
            break
        end_p = p
        p += 1

    section_pages = [p for p in all_pages if start_p <= p <= end_p]
    total_score = sum(page_score[p] for p in section_pages)
    avg_score = total_score / len(section_pages) if section_pages else 0
    confidence = min(0.95, 0.55 + avg_score * 0.04)

    # ── Extract a title from the first page of the section ───────────────────
    first_chunk = next((c for c in chunks if c["page"] == start_p), None)
    title = ""
    if first_chunk:
        for line in first_chunk["text"].split("\n"):
            line = line.strip()
            if line and len(line) < 120 and any(
                kw in line.lower()
                for kw in ["transformer", "power transformer", "14.", "chapter 14"]
            ):
                title = line
                break

    print(f"[DensityFinder] Section: pages {start_p}-{end_p} "
          f"({len(section_pages)} pages, avg_score={avg_score:.1f}, "
          f"confidence={confidence:.2f})")

    return {
        "start_page": start_p,
        "end_page": end_p,
        "section_title": title or "Power Transformer Technical Specifications",
        "confidence": confidence,
        "notes": (f"Peak-expand: peak p={peak_page} score={peak_score}, "
                  f"avg {avg_score:.1f} spec terms/page"),
    }




def get_section_text(chunks: list[dict], start_page: int, end_page: int) -> str:
    """Extract and concatenate text from a specific chunk-page range."""
    section_chunks = [
        c for c in chunks
        if start_page <= c["page"] <= end_page
    ]
    return "\n\n".join(c["text"] for c in section_chunks)


# Keywords that identify "site / service / ambient conditions" headings.
# Checked case-insensitively against heading text.
SITE_CONDITIONS_HEADING_KEYWORDS = [
    "service condition", "site condition", "ambient condition",
    "environmental condition", "climate condition",
    "operating condition", "installation site",
    "climatic condition", "weather condition",
]


def find_section_in_headings(
    headings: list[dict],
    heading_keywords: list[str],
) -> dict:
    """
    Find a section in the headings list by keyword-matching on heading text,
    then compute end_page as the next same-or-higher-level heading's page.

    This is a ZERO-LLM-COST lookup: it does a simple substring search.
    Suitable for well-defined section labels like 'Service Conditions'.

    Args:
        headings:         List of {text, chunk_page, level} dicts from PDF/DOCX.
        heading_keywords: Lower-case substrings to match against heading text.

    Returns:
        {start_page, end_page, section_title, confidence}
        or {start_page: None, ...} if no heading matched.
    """
    matched_idx = None
    for i, h in enumerate(headings):
        text_lower = h["text"].lower()
        if any(kw in text_lower for kw in heading_keywords):
            matched_idx = i
            break

    if matched_idx is None:
        return {"start_page": None, "end_page": None,
                "section_title": "", "confidence": 0.0}

    h_matched = headings[matched_idx]
    start_page = h_matched["chunk_page"]
    matched_level = h_matched.get("level", 2)

    # Find next heading of same-or-higher level → that page is end_page.
    end_page = None
    for h in headings[matched_idx + 1:]:
        if h.get("level", 2) <= matched_level and h["chunk_page"] > start_page:
            end_page = h["chunk_page"]   # inclusive: no -1
            break
    if end_page is None:
        end_page = start_page + 20  # short section: 20-page upper bound

    print(f"[SectionLookup] '{h_matched['text']}' → pages {start_page}-{end_page}")
    return {
        "start_page": start_page,
        "end_page": end_page,
        "section_title": h_matched["text"],
        "confidence": 0.85,
    }




# ─── Internal helpers ─────────────────────────────────────────────────────────

import re as _re

# Matches numbered section headings as they appear in EPC tender PDFs.
# Format: "14 Power Transformers" or "14.8 Transformer Design"
# Constraints:
#   - Number: 1-2 digits, optionally followed by ONE decimal group (.digits)
#     This naturally excludes deep sub-sections like "15.17.9.1 Something"
#   - Separator: 1+ whitespace chars (actual PDFs use single space)
#   - Title: starts with capital letter, letters/spaces/punctuation only
#     The negative lookahead (?!\d) stops it matching table entries like "14 300kV"
_PDF_HEADING_RE = _re.compile(
    r'^\s*(\d{1,2}(?:\.\d{1,2})?)\s+'
    r'(?!\d)([A-Z][A-Za-z][A-Za-z /\-,().&]{2,75})\s*$'
)

# Level-1 heading: single integer like "14", "15"
_PDF_L1_HEADING_RE = _re.compile(r'^\d{1,2}$')


def _extract_headings_from_pdf(chunks: list[dict]) -> list[dict]:
    """
    Scan each PDF page for lines that look like numbered section headings.

    Returns a list of {"text": str, "chunk_page": int, "level": int} dicts
    in the same format that _parse_docx produces for headings, so that the
    existing find_section_from_headings() pipeline can use them without
    modification.

    Strategy:
    - For each page, check every line against _PDF_HEADING_RE.
    - Keep only level-1 headings ("14 Power Transformers") and top-level
      level-2 headings ("14.8 ...") to keep the list manageable for the LLM.
    - Deduplicate headings that appear on consecutive pages (running headers).
    """
    headings: list[dict] = []
    seen_texts: set[str] = set()

    for chunk in chunks:
        page = chunk["page"]
        for raw_line in chunk["text"].split("\n"):
            line = raw_line.strip()
            if not line or len(line) > 100:
                continue
            m = _PDF_HEADING_RE.match(line)
            if not m:
                continue
            number_part = m.group(1)   # e.g. "14" or "14.8"
            title_part = m.group(2).strip()
            full_heading = f"{number_part}  {title_part}"

            # Determine level (1 = chapter, 2 = sub-section)
            level = 1 if _PDF_L1_HEADING_RE.match(number_part) else 2

            # Include level-1 headings always; level-2 only up to first decimal
            if level == 2 and number_part.count(".") > 1:
                continue  # skip 14.8.1 and deeper

            # Deduplicate (running page headers repeat the section title)
            key = full_heading.lower()
            if key in seen_texts:
                continue
            seen_texts.add(key)

            headings.append({
                "text": full_heading,
                "chunk_page": page,
                "level": level,
            })

    print(f"[PDFHeadings] Extracted {len(headings)} headings from PDF")
    if headings[:5]:
        for h in headings[:5]:
            print(f"  [{h['chunk_page']}] {h['text']}")
    return headings


def _extract_title_from_chunk(text: str, keywords: list[str]) -> str:
    """Find the first short line in *text* that contains any keyword."""
    for line in text.split("\n"):
        stripped = line.strip()
        if stripped and len(stripped) < 120:
            if any(kw.upper() in stripped.upper() for kw in keywords):
                return stripped
    return ""


def _parse_pdf(file_path: str) -> dict:
    """Parse PDF using pdfplumber, one chunk per page.

    Returns {"chunks": [...], "headings": [...]} — same shape as _parse_docx.
    Headings are extracted by scanning each page for numbered-section lines.
    """
    chunks: list[dict] = []
    with pdfplumber.open(file_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and text.strip():
                chunks.append({"page": page_num, "text": text.strip()})
    headings = _extract_headings_from_pdf(chunks)
    return {"chunks": chunks, "headings": headings}



def _parse_docx(file_path: str) -> dict:
    """
    Parse a Word document.

    - Iterates body elements in document order, capturing paragraphs AND tables.
    - Detects headings by paragraph style (Heading 1/2/3, Title, …).
    - Groups content into pseudo-pages of ~3 000 chars each.

    Returns {"chunks": [...], "headings": [...]}.
    """
    doc = Document(file_path)
    chunks: list[dict] = []
    headings: list[dict] = []

    current_text: list[str] = []
    current_chars = 0
    page_num = 1
    PAGE_SIZE = 3000

    def flush() -> None:
        nonlocal page_num, current_text, current_chars
        if current_text:
            chunks.append({"page": page_num, "text": "\n".join(current_text)})
            page_num += 1
            current_text = []
            current_chars = 0

    def add_text(text: str) -> None:
        nonlocal current_chars
        text = text.strip()
        if not text:
            return
        current_text.append(text)
        current_chars += len(text)
        if current_chars >= PAGE_SIZE:
            flush()

    # Iterate body elements in document order
    for child in doc.element.body:
        raw_tag = child.tag
        tag = raw_tag.split("}")[-1] if "}" in raw_tag else raw_tag

        if tag == "p":
            from docx.text.paragraph import Paragraph
            para = Paragraph(child, doc)
            text = para.text.strip()
            if not text:
                continue

            # Detect heading before flushing so page_num is accurate
            try:
                style_name = para.style.name if para.style else ""
            except Exception:
                style_name = ""

            if _is_heading_style(style_name):
                # Force a new chunk so the heading starts cleanly
                flush()
                headings.append({"text": text, "chunk_page": page_num})

            add_text(text)

        elif tag == "tbl":
            from docx.table import Table
            table = Table(child, doc)
            for row in table.rows:
                seen: set[int] = set()
                row_cells: list[str] = []
                for cell in row.cells:
                    cell_id = id(cell._tc)
                    if cell_id in seen:
                        continue
                    seen.add(cell_id)
                    cell_text = cell.text.strip()
                    if cell_text:
                        row_cells.append(cell_text)
                if row_cells:
                    add_text(" | ".join(row_cells))

    flush()  # flush any remaining text

    return {"chunks": chunks, "headings": headings}
