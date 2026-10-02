# VB6FastIni

A lightweight, dependency-free INI reader/writer for Visual Basic 6.

---

## Overview

VB6FastIni is a small reusable VB6 class that provides fast access to INI files.

Unlike the Windows API (`GetPrivateProfileString()`), the entire file is loaded into memory once, allowing subsequent reads and writes to occur without repeated disk access.

The design philosophy is simple:

- Small
- Readable
- Reliable
- Reusable
- Fast enough

Load once.

Work in memory.

Save once.

---

## Features

- Pure VB6
- No Windows API
- No external DLLs
- No dependencies
- Reads entire INI into memory
- Writes entire INI in one operation
- Preserves comments
- Preserves blank lines
- Preserves section ordering
- Case-insensitive section and key lookups

---

## Parser behavior

- Lines are separated by CRLF or LF. Saving modified content writes CRLF line endings; a bare CR is not treated as a line separator.
- Blank lines are preserved.
- A comment is a line whose first non-whitespace character is `;` or `#`. Inline comments are not recognized, so text after a value remains part of that value.
- After surrounding whitespace is removed, a section header must be at least three characters long, start with `[` and end with `]`. Whitespace inside the brackets is trimmed from the section name.
- A key/value line is split at its first `=`. Whitespace around the key is trimmed; the value is kept as written, including leading or trailing whitespace and any later `=` characters. A line containing `=` may have an empty key.
- Non-empty lines that match none of these forms are ignored for lookup and preserved unchanged as unparsed lines.
- Keys before the first section belong to the unnamed section (`""`).
- Section and key lookups ignore case. If duplicate keys occur, lookup, write, and delete operate on the first matching key in file order. Repeated section headers with the same name are treated as the same section for key lookup; `DeleteSection` removes only the first matching section block.

---

## File encoding

The reader and writer do not detect or convert file encodings. The class is intended for legacy ANSI INI files; UTF-8 and UTF-16 files, including files with a byte-order mark, are not supported as encoded formats. Non-ASCII characters may not round-trip correctly across different Windows system code pages.

---

## Numeric read behavior

- If a key is missing, `ReadInteger` and `ReadDouble` return the supplied default value (or `0` when no default is supplied).
- If a key exists but its value is empty or not numeric, the read raises error 13 (Type mismatch).
- `ReadInteger` accepts only whole-number values. Fractional values raise error 13; values outside the VB6 `Long` range raise error 6 (Overflow).
- `ReadDouble` converts values using VB6's `IsNumeric` and `CDbl` rules. Values outside the `Double` range raise an overflow error.
- Numeric text follows the host's VB6 numeric parsing rules, including its regional decimal and grouping separators.

---

## Planned Public Interface

```vb
LoadIni()

Save()

SaveAs()

ReadString()

WriteString()

ReadInteger()

WriteInteger()

ReadBoolean()

WriteBoolean()

DeleteKey()

DeleteSection()

KeyExists()

SectionExists()
```

---

## Status

Current version:

**0.1.0**

Under active development.

Not yet recommended for production use.

---

## Licence

MIT

---

## Author

Tim Owers

Developed with assistance from ChatGPT.
