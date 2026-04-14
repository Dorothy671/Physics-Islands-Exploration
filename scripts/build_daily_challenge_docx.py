# -*- coding: utf-8 -*-
"""生成 questions/每日挑战题目.docx（与压强题库相同 OOXML 结构）。"""
import os
import zipfile
from xml.sax.saxutils import escape

OUT_DIR = os.path.join(os.path.dirname(__file__), "..", "questions")
OUT_PATH = os.path.join(OUT_DIR, "每日挑战题目.docx")

ITEMS = [
    ("在国际单位制中，力的单位是？", ["牛顿", "千克", "米每秒", "帕斯卡"], "A", "牛顿（N）是力的单位。"),
    ("光在同种均匀介质中沿什么路径传播？", ["折线", "直线", "抛物线", "螺旋线"], "B", "同种均匀介质中光沿直线传播。"),
    ("声音的产生主要是因为？", ["温度变化", "物体振动", "颜色变化", "体积膨胀"], "B", "声音由物体振动产生。"),
    ("下列环境中，声音最难传播的是？", ["水中", "钢铁中", "真空中", "空气中"], "C", "真空不能传声。"),
    ("部分电路欧姆定律可写作？", ["I=U/R", "I=UR", "U=IR²", "R=U²/I"], "A", "I=U/R 为常见形式。"),
    ("惯性的大小主要取决于物体的？", ["速度", "质量", "温度", "形状"], "B", "惯性大小只与质量有关。"),
    ("在液体密度一定时，深度越大，液体压强？", ["越小", "不变", "越大", "为零"], "C", "p=ρgh，h 增大则 p 增大。"),
    ("平面镜所成的像与物体到镜面的距离关系是？", ["像更远", "相等", "物更远", "随机"], "B", "像距等于物距。"),
    ("两电阻串联后的总电阻 R 与各电阻关系为？", ["R=R1·R2", "R=R1+R2", "R=R1/R2", "R=|R1-R2|"], "B", "串联总电阻等于各电阻之和。"),
    ("凸透镜对平行于主光轴的入射光具有？", ["发散作用", "会聚作用", "全反射", "无偏折"], "B", "凸透镜对光有会聚作用。"),
]


def w_p(text):
    t = escape(text)
    return (
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'
    )


def build_document_xml():
    parts = [w_p("以下为初中物理每日挑战选择题（共10题）。题干 → A-D → 正确答案 → 解析。")]
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
