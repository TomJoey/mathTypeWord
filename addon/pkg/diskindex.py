"""Fast path: index every equation of a *saved* .docx file once, in a background thread.

Getting one equation from Word (Range.WordOpenXML) costs ~200 ms, but asking Word "which inline
shape is this" costs ~10 ms.  So for a saved document we read the .docx ourselves, list its inline
shapes in document order, and convert all equations up front; at reading time only the position
of the equation among the document's inline shapes is needed.

Safety: the index is only used when
  * the number of inline shapes found in the file equals Word's InlineShapes.Count,
  * the shape at that position is an equation, and
  * its width and height equal Word's (so a mis-numbered shape is practically impossible).
Anything else returns None and the caller falls back to the slow, always-correct method.
"""
import re
import threading
import zipfile

from . import core

try:
	from logHandler import log
except ImportError:  # tests outside NVDA
	class _Log:
		def __getattr__(self, name):
			return lambda *a, **k: None

	log = _Log()

_SHAPE_RE = re.compile(
	r"<w:object\b.*?</w:object>|<w:drawing\b.*?</w:drawing>|<w:pict\b.*?</w:pict>",
	re.S,
)
_FALLBACK_RE = re.compile(r"<mc:Fallback\b.*?</mc:Fallback>", re.S)
_REL_RE = re.compile(r"<Relationship\b[^>]*>")
_STYLE_SIZE_RE = re.compile(r"width:\s*([\d.]+)pt;\s*height:\s*([\d.]+)pt")
_SIZE_TOLERANCE = 1.0  # points


class Entry:
	__slots__ = ("kind", "progid", "width", "height", "mathml")

	def __init__(self, kind, progid=None, width=None, height=None, mathml=None):
		self.kind, self.progid, self.width, self.height, self.mathml = kind, progid, width, height, mathml


class DocIndex:
	def __init__(self, entries):
		self.entries = entries
		self.countChecked = False
		self.bad = False


def _attr(tag, name):
	m = re.search(rf'\b{re.escape(name)}="([^"]*)"', tag)
	return m.group(1) if m else None


def build(path):
	"""Read the docx and return a list of Entry (one per inline shape, in document order)."""
	with zipfile.ZipFile(path) as z:
		doc = z.read("word/document.xml").decode("utf8", "ignore")
		rels = {}
		relsXml = z.read("word/_rels/document.xml.rels").decode("utf8", "ignore")
		for m in _REL_RE.finditer(relsXml):
			rid, target = _attr(m.group(0), "Id"), _attr(m.group(0), "Target")
			if rid and target:
				rels[rid] = target
		# Word shows mc:Choice only; counting the fallback as well would double count.
		doc = _FALLBACK_RE.sub("", doc)
		entries = []
		for m in _SHAPE_RE.finditer(doc):
			s = m.group(0)
			if s.startswith("<w:object"):
				ole = re.search(r"<o:OLEObject\b[^>]*>", s)
				size = _STYLE_SIZE_RE.search(s)
				width = float(size.group(1)) if size else None
				height = float(size.group(2)) if size else None
				progid = _attr(ole.group(0), "ProgID") if ole else None
				rid = _attr(ole.group(0), "r:id") if ole else None
				mathml = None
				if progid and progid.startswith("Equation.") and rid in rels and width is not None:
					try:
						target = rels[rid]
						name = target.lstrip("/") if target.startswith("/") else "word/" + target
						mathml = core.mathml_from_ole_bytes(z.read(name))
					except Exception:
						log.debugWarning(f"mathTypeWord: cannot convert {rid} from the file", exc_info=True)
				entries.append(Entry("ole", progid, width, height, mathml))
			elif s.startswith("<w:drawing"):
				if "<wp:inline" in s:
					entries.append(Entry("drawing"))
			elif "<v:imagedata" in s:
				entries.append(Entry("pict"))
	return entries


_lock = threading.Lock()
_indexes = {}  # (path, mtime) -> DocIndex | None (None = being built / failed)


def _start_build(key):
	def work():
		try:
			index = DocIndex(build(key[0]))
			log.info(
				f"mathTypeWord: indexed {sum(1 for e in index.entries if e.mathml)} equations "
				f"of {len(index.entries)} inline shapes in {key[0]}",
			)
		except Exception:
			log.debugWarning("mathTypeWord: could not index the document file", exc_info=True)
			index = DocIndex([])
			index.bad = True
		with _lock:
			_indexes[key] = index

	with _lock:
		if key in _indexes:
			return
		_indexes[key] = None
	threading.Thread(target=work, name="mathTypeWord-index", daemon=True).start()


def prepare(path, mtime, wait=False):
	"""Start indexing (idempotent).  With wait=True block until done (used by tests)."""
	key = (path, mtime)
	_start_build(key)
	if wait:
		while True:
			with _lock:
				if _indexes.get(key) is not None:
					return
			threading.Event().wait(0.05)


def forget():
	with _lock:
		_indexes.clear()


def lookup(doc, rangeObj, offset, path, mtime):
	"""MathML for the equation starting at `offset`, or None if the fast path cannot be trusted."""
	key = (path, mtime)
	with _lock:
		index = _indexes.get(key, False)
	if index is False:
		_start_build(key)
		return None
	if index is None or index.bad:
		return None
	try:
		if not index.countChecked:
			if doc.InlineShapes.Count != len(index.entries):
				log.debug("mathTypeWord: inline shape count differs from file; fast path disabled for this document")
				index.bad = True
				return None
			index.countChecked = True
		if rangeObj.StoryType != 1:  # main text story only
			return None
		position = doc.Range(0, offset).InlineShapes.Count
		if position >= len(index.entries):
			return None
		entry = index.entries[position]
		if entry.mathml is None:
			return None
		shape = rangeObj.InlineShapes.Item(1)
		if (
			abs(shape.Width - entry.width) > _SIZE_TOLERANCE
			or abs(shape.Height - entry.height) > _SIZE_TOLERANCE
		):
			log.debug("mathTypeWord: shape size differs from file; using slow path")
			return None
		return entry.mathml
	except Exception:
		log.debugWarning("mathTypeWord: fast path failed", exc_info=True)
		return None
