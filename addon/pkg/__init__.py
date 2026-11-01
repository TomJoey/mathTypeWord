# Math reading for MathType / Equation Editor 3.0 equations in Microsoft Word.
#
# Why this exists: NVDA reads a MathType equation in Word by asking MathType's 32-bit
# MT6.dll for MathML (source/mathType.py).  Since NVDA became 64-bit (2026.1) that DLL
# cannot be loaded (nvaccess/nvda#20115), so MathType equations are not read at all.
# This add-on replaces that single step: the equation's OLE data is taken from Word (or from
# the saved .docx file) without MathType, the MTEF data is converted to MathML in pure
# Python, and NVDA's own Word support / MathCAT do the rest (speech, braille, navigation).
#
# It also makes "Equation.3" objects (Equation Editor 3.0, used by old documents) count
# as math: NVDA only treats ProgIDs starting with "Equation.DSMT" as MathType.
#
# Works with Word read through the object model and (see uiabridge.py) through UI Automation.
# Required NVDA setting: Math > "Use native math support in Word and Outlook" must be OFF
# (when on, NVDA removes all math providers while Word is focused).
import os
import time
from collections import OrderedDict

import globalPluginHandler
from logHandler import log

from . import core, diskindex, uiabridge

_CACHE_SIZE = 1024
_cache = OrderedDict()


_loggedReasons = set()


def _whyNot(reason):
	if reason not in _loggedReasons:
		_loggedReasons.add(reason)
		log.debug(f"mathTypeWord: document is not used for the fast path/cache: {reason}")


def _documentKey(rangeObj):
	"""(document, path, mtime) while the document equals the file on disk, else None.

	An unsaved (modified) document is never cached or read from the file, so an edited
	equation is always re-read from Word.
	"""
	try:
		doc = rangeObj.Document
		if not doc.Saved:
			_whyNot("document has unsaved changes (Word reports Saved=False)")
			return None
		path = doc.FullName
		if not os.path.isabs(path) or not os.path.isfile(path):
			_whyNot(f"document path is not a local file: {path!r}")
			return None
		return doc, path, os.path.getmtime(path)
	except Exception as e:
		_whyNot(f"cannot read document properties: {e!r}")
		return None


def _getMathMlViaMtef(self, field):
	shapeOffset = int(field["shapeoffset"])
	rangeObj = self._rangeObj.Duplicate
	rangeObj.Start = shapeOffset
	return _mathMlFor(rangeObj, shapeOffset)


def _mathMlFor(rangeObj, shapeOffset):
	"""MathML of the equation that starts at `shapeOffset` (rangeObj starts there too)."""
	docKey = _documentKey(rangeObj)
	key = None
	if docKey is not None:
		doc, path, mtime = docKey
		key = (path, mtime, shapeOffset)
		if key in _cache:
			_cache.move_to_end(key)
			return _cache[key]
		mathMl = diskindex.lookup(doc, rangeObj, shapeOffset, path, mtime)
		if mathMl is not None:
			_remember(key, mathMl)
			return mathMl
	# Slow but always correct: ask Word for the equation's OLE data.
	start = time.perf_counter()
	shapes = rangeObj.InlineShapes
	try:
		shape = shapes.Item(1)  # Word collections are 1-based
	except Exception:
		shape = shapes[0]
	mathMl = core.mathml_from_flat_opc(shape.Range.WordOpenXML)
	log.debug(
		f"mathTypeWord: read equation at {shapeOffset} from Word in {(time.perf_counter() - start) * 1000:.0f} ms",
	)
	if key is not None:
		_remember(key, mathMl)
	return mathMl


def _remember(key, mathMl):
	_cache[key] = mathMl
	while len(_cache) > _CACHE_SIZE:
		_cache.popitem(last=False)


def _uiaNote(message, always=False):
	"""Troubleshooting detail; only visible with NVDA's log level set to Debug."""
	log.debug(f"mathTypeWord UIA: {message}")


def _uiaRangeContainsCaret(textInfo, caretInfo):
	own, caret = textInfo._rangeObj, caretInfo._rangeObj
	START, END = 0, 1  # TextPatternRangeEndpoint_Start / _End
	return (
		own.CompareEndpoints(START, caret, START) <= 0
		and own.CompareEndpoints(END, caret, END) >= 0
	)


def _addMathToUIACommands(textInfo, commands):
	"""Turn the empty placeholder fields UIA gives for MathType equations into math fields."""
	import controlTypes
	import textInfos

	def isPlaceholder(cmd):
		if not (isinstance(cmd, textInfos.FieldCommand) and cmd.command == "controlStart"):
			return False
		field = cmd.field
		role = field.get("role")
		name = field.get("name") or ""
		# line / paragraph reading: an empty static text per equation
		if role == controlTypes.Role.STATICTEXT:
			return not name and not field.get("content")
		# character reading (left / right arrow): Word gives a graphic named after the OLE object,
		# e.g. "MathType 6.0 Equation" or "Microsoft Equation 3.0"
		if role in (controlTypes.Role.GRAPHIC, controlTypes.Role.EMBEDDEDOBJECT):
			lowered = name.lower()
			return "mathtype" in lowered or "equation" in lowered
		return False

	placeholders = [cmd for cmd in commands if isPlaceholder(cmd)]
	if not placeholders:
		return
	obj = textInfo.obj
	selection = getattr(obj, "WinwordSelectionObject", None)
	if not selection:
		_uiaNote("Word object model not available, equations stay silent")
		return
	start = time.perf_counter()
	byCaret = _uiaRangeContainsCaret(textInfo, obj.makeTextInfo(textInfos.POSITION_CARET))
	try:
		if byCaret:
			ranges = uiabridge.resolve(selection, len(placeholders))
		else:
			# say-all / reading ahead: locate the range by its screen geometry, verified by its text
			ranges = uiabridge.resolveByGeometry(textInfo._rangeObj, selection, len(placeholders))
			_uiaNote(f"geometry match for {len(placeholders)} placeholder(s): {ranges}", always=True)
	except uiabridge.Skip as e:
		try:
			sel = selection.Range
			uiaText = textInfo._rangeObj.GetText(60)
			shapes = uiabridge.describeShapes(sel)
			detail = f"caret={sel.Start}-{sel.End} uiaText={uiaText!r} shapes={shapes}"
		except Exception as e2:
			detail = f"(no detail: {e2!r})"
		_uiaNote(f"not matched ({len(placeholders)} placeholders): {e}; {detail}", always=True)
		return
	doc = selection.Range.Document
	mathMls = []
	for first, last in ranges:
		r = doc.Range(first, last)
		mathMls.append(_mathMlFor(r, first))
	for cmd, mathMl in zip(placeholders, mathMls):
		cmd.field["role"] = controlTypes.Role.MATH
		cmd.field["mathml"] = mathMl
	_uiaNote(f"{len(mathMls)} equation(s) found in {(time.perf_counter() - start) * 1000:.0f} ms")


class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def __init__(self):
		super().__init__()
		from NVDAObjects.window import winword

		cls = winword.WordDocumentTextInfo
		self._cls = cls
		self._origGetMathMl = cls.getMathMl
		self._origNormalize = cls._normalizeControlField
		origGetMathMl = self._origGetMathMl
		origNormalize = self._origNormalize

		def getMathMl(textInfo, field):
			try:
				return _getMathMlViaMtef(textInfo, field)
			except Exception:
				log.warning("mathTypeWord: could not get MathML, trying NVDA's own method", exc_info=True)
			try:
				return origGetMathMl(textInfo, field)
			except Exception:
				# NVDA's callers only handle LookupError.
				raise LookupError("Couldn't get MathML for this equation")

		def normalizeControlField(textInfo, field):
			# Equation Editor 3.0 / MathType 3 objects use ProgID "Equation.3".
			if field.get("role") == "object":
				progid = field.get("progid") or ""
				if progid.startswith("Equation.3"):
					field["progid"] = "Equation.DSMT-compat." + progid
			return origNormalize(textInfo, field)

		self._uiaCls = None
		try:
			from NVDAObjects.UIA import wordDocument as uiaWord

			uiaCls = uiaWord.WordDocumentTextInfo
			origUiaGetText = uiaCls.getTextWithFields

			def uiaGetTextWithFields(textInfo, formatConfig=None):
				commands = origUiaGetText(textInfo, formatConfig)
				try:
					_addMathToUIACommands(textInfo, commands)
				except Exception:
					log.debugWarning("mathTypeWord: UIA equation support failed for this line", exc_info=True)
					_uiaNote("failed, see debug log")
				return commands

			self._uiaCls, self._origUiaGetText = uiaCls, origUiaGetText
			uiaCls.getTextWithFields = uiaGetTextWithFields
		except Exception:
			log.warning("mathTypeWord: could not add UIA support", exc_info=True)
		cls.getMathMl = getMathMl
		cls._normalizeControlField = normalizeControlField
		log.info("mathTypeWord: MathType equation support for Word installed")

	def terminate(self):
		try:
			if self._uiaCls is not None:
				self._uiaCls.getTextWithFields = self._origUiaGetText
		except Exception:
			log.debugWarning("mathTypeWord: could not restore the UIA method", exc_info=True)
		try:
			self._cls.getMathMl = self._origGetMathMl
			self._cls._normalizeControlField = self._origNormalize
		except Exception:
			log.debugWarning("mathTypeWord: could not restore NVDA methods", exc_info=True)
		_cache.clear()
		diskindex.forget()
		super().terminate()
