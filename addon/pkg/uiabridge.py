"""UIA support: give the equations of a MathType document a voice when Word is read through UIA.

In UIA mode Word does not put MathType (OLE) equations into the text; for each equation NVDA only
gets an empty static-text control field at the equation's position.  Word's UIA tree has no way to
tell which equation of the document it is (NVDA issue #13842), but the *object model* knows exactly
where the caret is.  So, for the text range being read (which must contain the caret), we line up
the empty fields with the equations of the same line / paragraph in the object model, in order.

Safety rules (anything unclear -> do nothing, the text is then read exactly as before):
  * the UIA range must contain the caret,
  * the number of placeholder fields must equal the number of equations found in the object model
    for exactly one candidate (a single line, or the whole paragraph), never for two different ones,
  * every candidate must be an OLE equation (ProgID "Equation.*").
"""
from logHandler import log

# Information(wdFirstCharacterLineNumber)
_WD_LINE = 10


class Skip(Exception):
	"""The placeholders cannot be matched with certainty."""


# OLEFormat.ProgID costs ~70 ms per call in Word, so remember the answer for a shape that is still
# the same OLE object at the same place: (document, start, end, document length) -> is an equation.
_progIdCache = {}
_CACHE_MAX = 4000


def _isEquation(shape, docName, docLength, start, end):
	if shape.Type != 1:  # wdInlineShapeEmbeddedOLEObject
		return False
	key = (docName, docLength, start, end)
	known = _progIdCache.get(key)
	if known is not None:
		return known
	try:
		isEquation = (shape.OLEFormat.ProgID or "").startswith("Equation.")
	except Exception:
		isEquation = False
	if len(_progIdCache) >= _CACHE_MAX:
		_progIdCache.clear()
	_progIdCache[key] = isEquation
	return isEquation


def _equationShapes(paragraphRange, docName, docLength):
	"""[(start, end, isEquation)] of the inline shapes of a paragraph."""
	shapes = paragraphRange.InlineShapes
	count = shapes.Count
	found = []
	for i in range(1, count + 1):
		shape = shapes.Item(i)
		r = shape.Range
		start, end = r.Start, r.End
		found.append((start, end, _isEquation(shape, docName, docLength, start, end)))
	return found


def resolve(selection, placeholders):
	"""Return [(start, end)] object-model ranges of the equations the placeholders stand for.

	`selection` is Word's Selection (the caret); `placeholders` is how many empty fields were found
	in the UIA text.  Raises Skip when the match is not certain.
	"""
	if placeholders <= 0:
		raise Skip("no placeholders")
	rng = selection.Range
	if rng.StoryType != 1:
		raise Skip("not the main text story")
	doc = rng.Document
	# The equation under the caret (caret on / just before an equation), for char-by-char reading.
	if placeholders == 1:
		here = doc.Range(rng.Start, rng.Start + 1)
		if here.InlineShapes.Count == 1:
			shape = here.InlineShapes.Item(1)
			if shape.Type == 1 and (shape.OLEFormat.ProgID or "").startswith("Equation."):
				r = shape.Range
				return [(r.Start, r.End)]
	para = rng.Paragraphs(1).Range
	shapes = _equationShapes(para, doc.FullName, doc.Content.End)
	if not shapes:
		raise Skip("no inline shapes in the paragraph")
	if placeholders == 1:
		# Character reading (left / right arrow): the caret sits exactly at the start of the
		# equation (an OLE equation is ~26 object-model characters wide, so the one-character
		# range test above cannot see it) and NVDA reads the character after the caret.
		atCaret = [(s, e) for s, e, isEq in shapes if s == rng.Start and isEq]
		if len(atCaret) == 1:
			return atCaret
	if any(not isEq for _, _, isEq in shapes):
		# a picture or something else shares the paragraph: UIA's placeholders may include it
		raise Skip("paragraph contains inline shapes that are not equations")
	n = len(shapes)
	if placeholders > n:
		raise Skip(f"{placeholders} placeholders but only {n} equations in the paragraph")
	if placeholders == n:
		# Every equation of the paragraph is a placeholder.  A single line holding exactly n
		# equations would be this very same set, so there is no other candidate and no need for
		# the (slow, ~65 ms each) per-line queries.
		return [(s, e) for s, e, _ in shapes]
	candidates = []
	# per-line groups: only needed when the paragraph has more equations than placeholders, or to
	# make sure a whole-paragraph match is not also a single-line match of another split
	lineOf = [doc.Range(s, s + 1).Information(_WD_LINE) for s, _, _ in shapes]
	groups = {}
	for (s, e, _), line in zip(shapes, lineOf):
		groups.setdefault(line, []).append((s, e))
	caretLine = rng.Information(_WD_LINE)
	nearby = []
	for line, members in groups.items():
		if abs(line - caretLine) <= 1 and len(members) == placeholders:
			nearby.append((line, tuple(members)))
	candidates.extend(m for _, m in nearby)
	distinct = set(candidates)
	if len(distinct) == 1:
		return list(distinct.pop())
	# several equally plausible matches: trust the caret's own line only when the caret is clearly
	# inside that line (the positions before and after it are on the same line, so UIA and the
	# object model cannot disagree about which line the caret is on)
	here = rng.Start
	if here > para.Start and here < para.End - 1:
		inside = (
			doc.Range(here - 1, here - 1).Information(_WD_LINE) == caretLine
			and doc.Range(here + 1, here + 1).Information(_WD_LINE) == caretLine
		)
		if inside:
			own = [m for line, m in nearby if line == caretLine]
			if len(own) == 1:
				return list(own[0])
	raise Skip(f"{len(distinct)} possible matches for {placeholders} placeholders")


def describeShapes(rng):
	"""For the diagnostic log: inline shapes of the paragraph of `rng` as (start, end, isEquation)."""
	doc = rng.Document
	return _equationShapes(rng.Paragraphs(1).Range, doc.FullName, doc.Content.End)


# ---------------------------------------------------------------------------------------------
# Ranges that do not contain the caret (say-all, reading ahead): locate them by geometry.
#
# The UIA range knows where it is on the screen; Word's ActiveWindow.RangeFromPoint turns a screen
# point back into an object-model position.  Geometry can be off (DPI scaling, scrolling, hidden
# text), so the result is only trusted when the *text* of the object-model range equals the text
# of the UIA range (ignoring white space and the equation placeholders) and the number of OLE
# equations in it equals the number of placeholders.  Anything else -> Skip (silence).
# ---------------------------------------------------------------------------------------------
_IGNORED = set("￼​‎‏﻿ ")


def _significant(text):
	return "".join(ch for ch in (text or "") if not ch.isspace() and ord(ch) >= 32 and ch not in _IGNORED)


def resolveByGeometry(uiaRange, selection, placeholders):
	"""Return [(start, end)] for the equations of `uiaRange` (a UIA text range anywhere on screen)."""
	if placeholders <= 0:
		raise Skip("no placeholders")
	rects = list(uiaRange.GetBoundingRectangles() or ())
	if len(rects) < 4:
		raise Skip("range has no bounding rectangles (off screen?)")
	first = rects[0:4]
	last = rects[-4:]
	if first[2] <= 2 or last[2] <= 2:
		raise Skip("degenerate bounding rectangle")
	window = selection.Application.ActiveWindow
	doc = selection.Range.Document
	if selection.Range.StoryType != 1:
		raise Skip("not the main text story")
	x1, y1 = int(first[0] + 2), int(first[1] + first[3] / 2)
	x2, y2 = int(last[0] + last[2] - 2), int(last[1] + last[3] / 2)
	a = window.RangeFromPoint(x1, y1).Start
	b = window.RangeFromPoint(x2, y2).End
	if b <= a:
		raise Skip(f"geometry gave an empty range {a}-{b}")
	if b - a > 5000:
		raise Skip(f"geometry gave a suspiciously long range {a}-{b}")
	omRange = doc.Range(a, b)
	uiaText = _significant(uiaRange.GetText(-1))
	omText = _significant(omRange.Text)
	# UIA text starts with the automatic list label ("1.", "A.") of the paragraph; the object
	# model's Range.Text does not contain it.
	startPara = doc.Range(a, a).Paragraphs(1).Range
	if startPara.Start == a:
		try:
			omText = _significant(doc.Range(a, a).ListFormat.ListString) + omText
		except Exception:
			pass
	if uiaText != omText:
		raise Skip(f"text mismatch for {a}-{b}: uia={uiaText[:40]!r} word={omText[:40]!r}")
	docName, docLength = doc.FullName, doc.Content.End
	shapes = omRange.InlineShapes
	found = []
	for i in range(1, shapes.Count + 1):
		shape = shapes.Item(i)
		r = shape.Range
		if not _isEquation(shape, docName, docLength, r.Start, r.End):
			raise Skip("the range contains an inline shape that is not an equation")
		found.append((r.Start, r.End))
	if len(found) != placeholders:
		raise Skip(f"{placeholders} placeholders but {len(found)} equations in {a}-{b}")
	return found
