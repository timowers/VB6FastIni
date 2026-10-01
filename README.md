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
