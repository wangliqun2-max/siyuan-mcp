"""检查提取的文本里是否包含运行环境、第8、第15小节相关关键词"""
import sys
sys.path.insert(0, 'backend')
from services.doc_parser import parse_document, get_section_text

r = parse_document('Section V Part- B2 - Standard AIS Rev1.docx')
text = get_section_text(r['chunks'], 84, 125)
print(f"提取文本总字符数: {len(text)}\n")

checks = {
    "=== 第1节：运行环境 ===": ['altitude', 'ambient temperature', 'pollution', 'humidity', 'rainfall', 'wind speed', 'seismic', 'solar radiation'],
    "=== 第8节相关 ===":      ['buchholz', 'sudden pressure', 'winding temperature', 'oil temperature', 'gas relay', 'temperature relay'],
    "=== 第15节相关 ===":     ['nameplate', 'name plate', 'terminal marking', 'accessories', 'thermometer', 'oil level', 'pressure relief'],
}

for section, keywords in checks.items():
    print(section)
    for kw in keywords:
        found = kw.lower() in text.lower()
        status = "✓ Found" if found else "✗ NOT FOUND"
        print(f"  {status:12s}  {kw}")
    print()
