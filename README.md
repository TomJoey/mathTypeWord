# MathType Equations for Word (NVDA add-on)

Lets [NVDA](https://www.nvaccess.org/) read and navigate **MathType** and **Equation Editor 3.0** equations in Microsoft Word without MathType's 32-bit DLL.

Since NVDA became 64-bit (2026.1) its built-in MathType support no longer works ([nvaccess/nvda#20115](https://github.com/nvaccess/nvda/issues/20115)). This add-on takes the equation data from Word (or from the saved `.docx`), converts the MTEF data to MathML in pure Python, and hands it to NVDA's math support (e.g. MathCAT) for speech, braille and navigation.

* Works with Word read through the object model **and** through UI Automation (line, character and say-all reading).
* Also handles old `Equation.3` objects (Equation Editor 3.0).
* When an equation cannot be matched with certainty it stays silent; it is never read wrongly.

See [`addon/doc/en/readme.html`](addon/doc/en/readme.html) ([中文](addon/doc/zh_CN/readme.html)) for requirements (Math > "Use native math support in Word and Outlook" must be **off**) and known limits.

## Build

```
pip install olefile
python addon/build_addon.py      # -> addon/dist/mathTypeWord-<version>.nvda-addon
```

## Layout

* `addon/` – the NVDA global plugin (`pkg/`), docs and the build script.
* `mtef/src/` – pure-Python MTEF v3/v5 parser and MathML converter.
* `mtef/tools_convert_docx.py` – optional stand-alone tool: replace MathType equations in a `.docx` by native Word equations (needs `lxml`, `olefile` and Office's `MML2OMML.XSL`).

## License

GPL v2 or later. `olefile` (bundled in the add-on) is BSD-licensed.
Not affiliated with Wiris or Design Science; MathType is a trademark of its owners.
