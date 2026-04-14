# -*- coding: utf-8 -*-
"""生成 questions/压强题目.docx（最小 OOXML，仅依赖标准库）。"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "questions")
OUT_PATH = os.path.join(OUT_DIR, "压强题目.docx")

ITEMS = [
    ("马德堡半球实验主要说明了什么？", ["大气压强的存在", "液体压强与深度有关", "浮力产生的原因", "分子间存在引力"], "A", "半球实验是经典的大气压存在证明。"),
    ("标准大气压约为多少？", ["1.0×10⁵ Pa", "1.0×10³ Pa", "1.0×10⁷ Pa", "1 Pa"], "A", "约 1.013×10⁵ Pa，常近似为 10⁵ Pa。"),
    ("用吸管喝饮料，主要利用了？", ["大气压", "液体压强", "重力", "摩擦力"], "A", "吸气使管内气压降低，外界大气压把液体压入管内。"),
    ("海拔升高，大气压强通常如何变化？", ["减小", "增大", "不变", "先增后减"], "A", "空气稀薄，气压随高度升高而降低。"),
    ("托里拆利实验中，玻璃管内水银柱高度主要取决于？", ["外界大气压", "玻璃管粗细", "水银槽形状", "管是否倾斜"], "A", "水银柱产生的压强与外界大气压平衡。"),
    ("高压锅内食物更容易煮熟，主要原因是？", ["锅内气压高，沸点升高", "锅内气压低，沸点降低", "导热更快", "水量更多"], "A", "气压升高，水的沸点升高。"),
    ("钢笔吸墨水时，挤压橡皮管是为了？", ["排出空气，利用大气压压入墨水", "增大摩擦力", "过滤杂质", "降温"], "A", "先减小内部气压，再靠大气压把墨水压入。"),
    ("离心式水泵抽水高度有限，主要受限于？", ["当地大气压能支撑的水柱高度", "电机功率", "水管颜色", "水的密度变化"], "A", "理论上大气压最多支撑约10.3m水柱（标准大气压）。"),
    ("下列现象中与大气压无关的是？", ["拦河坝上窄下宽", "吸盘吸附在玻璃上", "覆杯实验", "针筒抽取药液"], "A", "拦河坝形状主要考虑液体压强随深度增大。"),
    ("做托里拆利实验时，若玻璃管顶端开孔，水银柱会？", ["下降直至与槽内液面相平", "继续升高", "不变", "喷出管外"], "A", "顶端开孔后管内外都与大气相通，不再形成真空支撑。"),
]


def w_p(text):
    t = escape(text)
    return (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
    )


def build_document_xml():
    parts = [
        w_p("以下为大气压强选择题（共10题）。格式：题干 → A/B/C/D → 正确答案 → 解析。"),
    ]
    letters = "ABCD"
    for i, (q, opts, ans, exp) in enumerate(ITEMS, 1):
        parts.append(w_p(f"{i}. {q}"))
        for j, opt in enumerate(opts):
            parts.append(w_p(f"{letters[j]}. {opt}"))
        parts.append(w_p(f"正确答案：{ans}"))
        parts.append(w_p(f"解析：{exp}"))
        parts.append(w_p(""))
    body_inner = "\n".join(parts)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {body_inner}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>"""


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>"""

RELS_ROOT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr/></w:rPrDefault></w:docDefaults>
</w:styles>"""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    doc_xml = build_document_xml()
    with zipfile.ZipFile(OUT_PATH, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", RELS_ROOT)
        z.writestr("word/document.xml", doc_xml)
        z.writestr("word/_rels/document.xml.rels", DOC_RELS)
        z.writestr("word/styles.xml", STYLES_XML)
    print("Wrote", OUT_PATH)


if __name__ == "__main__":
    main()
