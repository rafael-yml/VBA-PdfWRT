# PdfWRT

> Render PDF pages to PNG images directly from VBA. No external tools, no shell execution, no admin rights, no installs.

A VBA class that renders PDF pages to high-quality PNG images using direct WinRT vtable calls to `Windows.Data.Pdf`, the same native PDF renderer used by Microsoft Edge. Works entirely within the VBA runtime with no subprocess spawning, making it suitable for locked-down corporate environments where shell execution and external tools are restricted or monitored.

---

## Motivation

The standard approaches for PDF-to-image conversion from VBA all have significant drawbacks in enterprise environments:

- **xpdf tools** (`pdftotext.exe`, `pdftopng.exe`) really awesome tools that I have use many many times in the past but require shell execution which is commonly blocked or flagged by SIEM/EDR solutions
- **WScript.Shell / Shell()** often restricted by Group Policy or application whitelisting, at least it was the case for me
- **PowerShell via WScript.Shell** PowerShell execution and especially `-EncodedCommand` are high-signal IOCs that trigger alerts in most modern SIEM setups
- **Word COM object** can open PDFs but produces poor rasterization quality and is extremely slow; IMPOSSIBLE export individual pages as images I tried so many methods, including passing files through PowerPoint
- **Adobe Acrobat COM** requires full Acrobat (paid), not just Reader

The goal was a solution that:
1. Produces sharp, high-fidelity output
2. Never spawns a subprocess or touches the shell
3. Requires nothing beyond what is already installed on any modern Windows machine
4. Does not trigger Defender, SIEM, or EDR alerts

---

## Inspiration

This project is architecturally based on [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR), which is a super cool VBA class that performs OCR by making direct WinRT vtable calls to `Windows.Media.Ocr.OcrEngine` via `DispCallFunc`. Studying that codebase revealed that the same pattern could be applied to `Windows.Data.Pdf` to achieve PDF rendering — effectively giving VBA access to the full Windows PDF rendering pipeline without any intermediary process. I was so impressed with it the day I found it.

The same Kernel32-free approach from my own fork of VBA-UWPOCR (Which I call VBA-WinOCR) [My Fork](https://github.com/rafael-yml/VBA-WinOCR) is used here, replacing `RtlMoveMemory` from `Kernel32.dll` with `memcpy` from `msvcrt.dll` to avoid triggering Windows Defender in corporate environments.

VBA-PdfWRT pairs naturally with VBA-WinOCR to form a complete PDF → PNG → text pipeline entirely in native VBA:

```
PDF
 └─► VBA-PdfWRT (Windows.Data.Pdf)       →  per-page PNG images
      └─► VBA-WinOCR (Windows.Media.Ocr) →  extracted text
```

---

## Usage

```vb
Dim pdfRenderer As New PdfWRT
pdfRenderer.RenderPDFToImages "C:\Docs\report.pdf", "C:\Output\pages", 2480
```

### Parameters

| Parameter | Type | Description |
|---|---|---|
| `pdfPath` | String | Full path to the input PDF file |
| `outputFolder` | String | Folder where PNG files will be saved (created if it doesn't exist) |
| `widthPx` | Long | Output width in pixels (optional, default 2480) |

### Resolution guide

| `widthPx` | DPI (A4) | Use case |
|---|---|---|
| 1240 | 150 dpi | Quick preview |
| 2480 | 300 dpi | Standard quality, OCR |
| 4960 | 600 dpi | High fidelity archival |

Output files are named `page_001.png`, `page_002.png`, etc.

---

## How it works

Modern Windows installs ship with `Windows.Data.Pdf`, a WinRT component that exposes a high-quality PDF renderer (the same one used by Microsoft Edge). Normally this API is only accessible from UWP apps or .NET code, but it can be called directly from VBA.

---

## Confirmed vtable offsets

Verified on Windows 10 21H2 (build 19044) and Windows 11 23H2 (build 22631).

| Interface | IID | Method | Vtable Offset |
|---|---|---|---|
| IPdfDocumentStatics | `{433A0B5F-C007-4788-90F2-08143D922599}` | `LoadFromStreamAsync` | 8 |
| IPdfDocument | — | `GetPage` | 6 |
| IPdfDocument | — | `get_PageCount` | 7 |
| IPdfPage | — | `RenderToStreamAsync` | 6 |
| IPdfPage | — | `RenderToStreamWithOptionsAsync` | 7 |
| IPdfPageRenderOptions | — | `set_DestinationWidth` | 9 |

The factory IID `{433A0B5F-C007-4788-90F2-08143D922599}` was determined at runtime via `IInspectable.GetIids` enumeration.

---

## DLL dependencies

All DLLs are present on every modern Windows installation — nothing needs to be installed or registered.

| DLL | Usage |
|---|---|
| `Combase.dll` | `RoGetActivationFactory`, `RoActivateInstance`, `WindowsCreateString`, `WindowsDeleteString` |
| `Shcore.dll` | `CreateRandomAccessStreamOnFile` |
| `oleAut32.dll` | `DispCallFunc` (vtable call dispatcher) |
| `ole32.dll` | `CLSIDFromString` |
| `msvcrt.dll` | `memcpy` (used instead of Kernel32 `RtlMoveMemory` to avoid Defender triggers) |

---

## Tested environments

| OS | Edition | Version | Build |
|---|---|---|---|
| Windows 10 | Enterprise | 21H2 | 19044 |
| Windows 11 | Enterprise | 23H2 | 22631 |

---

## License

MIT License — see [LICENSE](LICENSE) for details.

---

## Credits

- [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR): architectural foundation and vtable call pattern
- [rafael-yml/VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR): Kernel32-free revision using msvcrt.dll

Copyright © 2026, [rafae-yml](https://rafael-yml.lovable.app/)
