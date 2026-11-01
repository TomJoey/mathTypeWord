"""Replace MathType / Equation Editor 3.0 OLE equations in a .docx by native Word equations (OMML).

usage: python tools_convert_docx.py <input.docx> <output.docx>

* The input file is only read.  The output is a new file.
* Each OLE equation is parsed from its MTEF data (mtef.py), converted to MathML (mtef2mathml.py)
  and then to OMML with the stylesheet that ships with Office (MML2OMML.XSL).
* Equations that cannot be converted are left untouched (reported at the end).
"""
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET

from lxml import etree

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))
import mtef  # noqa: E402
import mtef2mathml  # noqa: E402
import olefile  # noqa: E402

XSL_CANDIDATES = [
    r"C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\MML2OMML.XSL",
]
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
O = "urn:schemas-microsoft-com:office:office"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def mathml_from_bin(data):
    ole = olefile.OleFileIO(data)
    try:
        stream = ole.openstream("Equation Native").read()
    finally:
        ole.close()
    parser, nodes = mtef.parse_equation_native(stream)
    math, _ = mtef2mathml.to_mathml(parser, nodes)
    return ET.tostring(math, encoding="unicode")


def convert(src, dst):
    xsl_path = next(p for p in XSL_CANDIDATES if os.path.exists(p))
    xslt = etree.XSLT(etree.parse(xsl_path))
    zin = zipfile.ZipFile(src)
    doc = etree.fromstring(zin.read("word/document.xml"))
    relsRoot = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
    rid2target = {r.get("Id"): r.get("Target") for r in relsRoot}
    converted = failed = 0
    problems = []
    for obj in list(doc.iter("{%s}object" % W)):
        ole = obj.find("{%s}OLEObject" % O)
        if ole is None or not (ole.get("ProgID") or "").startswith("Equation."):
            continue
        try:
            target = rid2target[ole.get("{%s}id" % R)]
            data = zin.read("word/" + target)
            omml = xslt(etree.fromstring(mathml_from_bin(data).encode("utf8"))).getroot()
            if omml is None or etree.QName(omml).localname != "oMath":
                raise ValueError("unexpected XSLT result")
        except Exception as e:  # leave this equation as it is
            failed += 1
            problems.append((obj.sourceline, repr(e)))
            continue
        run = obj.getparent()
        parent = run.getparent()
        others = [c for c in run if c is not obj and etree.QName(c).localname != "rPr"]
        if not others:
            parent.replace(run, omml)
        else:
            run.remove(obj)
            parent.insert(parent.index(run) + 1, omml)
        converted += 1

    # drop relationships (and parts) that nothing references any more
    newDoc = etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)
    used = set(re.findall(rb'r:(?:id|embed|link|pict)="([^"]+)"', newDoc))
    removedParts = set()
    for rel in list(relsRoot):
        if rel.get("Id").encode() in used:
            continue
        if rel.get("TargetMode") == "External":
            continue
        if rel.get("Type", "").endswith(("/oleObject", "/image")):
            relsRoot.remove(rel)
            removedParts.add("word/" + rel.get("Target"))
    # keep a part if another .rels file still points at it
    for name in zin.namelist():
        if name.endswith(".rels") and name != "word/_rels/document.xml.rels":
            for t in re.findall(rb'Target="([^"]+)"', zin.read(name)):
                removedParts.discard("word/" + t.decode())
                removedParts.discard(t.decode().lstrip("/"))
    newRels = etree.tostring(relsRoot, xml_declaration=True, encoding="UTF-8", standalone=True)

    os.makedirs(os.path.dirname(os.path.abspath(dst)), exist_ok=True)
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in removedParts:
                continue
            if item.filename == "word/document.xml":
                zout.writestr(item, newDoc)
            elif item.filename == "word/_rels/document.xml.rels":
                zout.writestr(item, newRels)
            else:
                zout.writestr(item, zin.read(item.filename))
    return converted, failed, problems, len(removedParts)


if __name__ == "__main__":
    c, f, problems, removed = convert(sys.argv[1], sys.argv[2])
    print(f"converted {c} equations, {f} left as they were, removed {removed} unused parts -> {sys.argv[2]}")
    for line, err in problems[:10]:
        print("  not converted near line", line, err)
