"""
快速诊断：打印DOCX文件的所有标题，帮助分析locate_section失败原因
用法：python debug_headings.py <docx文件路径>
"""
import sys
import os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'backend'))

from services.doc_parser import parse_document

if len(sys.argv) < 2:
    print("用法: python debug_headings.py <docx文件路径>")
    sys.exit(1)

doc_path = sys.argv[1]
print(f"解析文件: {doc_path}")
result = parse_document(doc_path)

headings = result.get("headings", [])
print(f"\n共找到 {len(headings)} 个标题\n")
print(f"{'chunk页':>7}  {'物理页':>6}  {'级别':>4}  标题")
print("-" * 80)

for h in headings:
    chunk = h.get('chunk_page', h.get('page', '?'))
    phys = h.get('physical_page', '?')
    level = h.get('level', '?')
    text = h.get('text', '')
    marker = " <<<" if 'TRANSFORMER' in text.upper() else ""
    print(f"{chunk:>7}  {phys:>6}  L{level:>3}  {text[:55]}{marker}")
