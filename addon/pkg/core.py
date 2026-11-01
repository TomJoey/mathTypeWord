"""NVDA-independent core: Word flat-OPC XML (Range.WordOpenXML) -> MathML."""
import base64
import io
import re
import xml.etree.ElementTree as ET

from . import olefile
from .mtef import parse_equation_native
from .mtef2mathml import to_mathml

# the embedded OLE object of an inline shape, inside the flat OPC package
_EMBED_RE = re.compile(
    r'<pkg:part\s[^>]*pkg:name="/word/embeddings/[^"]+"[^>]*>'
    r"(?:(?!</pkg:part>).)*?<pkg:binaryData>(.*?)</pkg:binaryData>",
    re.S,
)


def ole_bytes_from_flat_opc(xml):
    m = _EMBED_RE.search(xml)
    if not m:
        raise LookupError("no embedded OLE object in range")
    return base64.b64decode(re.sub(r"\s+", "", m.group(1)))


def mathml_from_ole_bytes(data):
    ole = olefile.OleFileIO(io.BytesIO(data))
    try:
        stream = ole.openstream("Equation Native").read()
    finally:
        ole.close()
    parser, nodes = parse_equation_native(stream)
    math, _warnings = to_mathml(parser, nodes)
    return ET.tostring(math, encoding="unicode")


def mathml_from_flat_opc(xml):
    return mathml_from_ole_bytes(ole_bytes_from_flat_opc(xml))
