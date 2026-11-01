"""Assemble the NVDA add-on: python build_addon.py  ->  dist/mathTypeWord-<version>.nvda-addon"""
import os, re, shutil, zipfile, sys
VERSION = "0.4.1"
HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "mtef", "src")
import olefile
OLE = os.path.dirname(olefile.__file__)
out = os.path.join(HERE, "build"); shutil.rmtree(out, ignore_errors=True)
gp = os.path.join(out, "globalPlugins", "mathTypeWord"); os.makedirs(gp)
for f in ("__init__.py", "core.py", "diskindex.py", "uiabridge.py"):
    shutil.copy(os.path.join(HERE, "pkg", f), gp)
for f in ("mtef.py", "mtef2mathml.py"):
    s = open(os.path.join(SRC, f), encoding="utf8").read()
    s = s.replace("from mtef import", "from .mtef import")
    open(os.path.join(gp, f), "w", encoding="utf8").write(s)
shutil.copy(os.path.join(OLE, "olefile.py"), gp)
shutil.copytree(os.path.join(HERE, "doc"), os.path.join(out, "doc"))
shutil.copy(os.path.join(OLE, "LICENSE.txt"), os.path.join(gp, "olefile-LICENSE.txt"))
open(os.path.join(out, "manifest.ini"), "w", encoding="utf8").write(f'''name = mathTypeWord
summary = "MathType Equations for Word"
license = GPL v2
licenseURL = https://www.gnu.org/licenses/gpl-2.0.html
description = """Lets NVDA read MathType and Equation Editor 3.0 equations in Microsoft Word, which stopped working when NVDA became 64-bit (NVDA issue 20115). Equation data is taken from Word and converted to MathML in pure Python; NVDA's own Word support and MathCAT then speak and navigate it."""
author = "TomJoey"
url = https://github.com/TomJoey/mathTypeWord
version = {VERSION}
minimumNVDAVersion = 2026.1.0
lastTestedNVDAVersion = 2026.3.0
updateChannel = None
''')
os.makedirs(os.path.join(HERE, "dist"), exist_ok=True)
dst = os.path.join(HERE, "dist", f"mathTypeWord-{VERSION}.nvda-addon")
with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
    for root, _, files in os.walk(out):
        for f in files:
            p = os.path.join(root, f); z.write(p, os.path.relpath(p, out))
print(dst, os.path.getsize(dst), "bytes")
