# PdfWRT

Render PDF pages to PNG images directly from VBA. No external tools, no shell execution, no admin rights, no installs.

Uses direct WinRT vtable calls to `Windows.Data.Pdf`, the same native PDF renderer used by Microsoft Edge, entirely within the VBA runtime.

---

## Motivation

Standard approaches for PDF-to-image conversion from VBA all carry significant drawbacks in enterprise environments:

- **xpdf tools** (`pdftotext.exe`, `pdftopng.exe`) require shell execution, which is commonly blocked or flagged by SIEM/EDR solutions
- **WScript.Shell / Shell()** often restricted by Group Policy or application whitelisting
- **PowerShell via WScript.Shell** `-EncodedCommand` is a high-signal IOC in most SIEM setups
- **Word COM object** poor rasterization quality, extremely slow, cannot export individual pages as images
- **Adobe Acrobat COM** requires full Acrobat (paid)

PdfWRT produces sharp, high-fidelity output, never spawns a subprocess, requires nothing beyond what is installed on any modern Windows machine, and does not trigger Defender, SIEM, or EDR alerts.

---

## Usage

### Render to files

```vb
Dim pdf As New PdfWRT
pdf.RenderPDFToImages "C:\Docs\report.pdf", "C:\Output\pages"
```

| Parameter | Type | Default | Description |
|---|---|---|---|
| `pdfPath` | String | | Full path to the input PDF file |
| `outputFolder` | String | | Folder where PNG files will be saved (created if it does not exist) |
| `widthPx` | Long | `DefaultWidth` | Output width in pixels; 0 or omitted uses `DefaultWidth` |

Output files are named `page_001.png`, `page_002.png`, etc.

### Render to memory (zero disk I/O)

```vb
Dim pdf As New PdfWRT
Dim oPages As Collection
Set oPages = pdf.RenderPDFToBytes("C:\Docs\report.pdf")
```

Returns a `Collection` of `Byte()` arrays, one per page in order. Renders to a GUID-named temp folder under `%TEMP%`, reads each PNG into a `Byte()` array, and deletes all files before returning. The caller sees no disk activity. Returns an empty `Collection` on failure.

### Resolution guide

| `widthPx` | Approx DPI (A4) | Notes |
|---|---|---|
| 1240 | ~150 dpi | Quick preview |
| 2480 | ~300 dpi | Standard quality, default |
| 4960 | ~600 dpi | High fidelity archival |

---

---

## Properties

#### `DefaultWidth` → `Long` (default `2480`)

Render width in pixels used when `widthPx` is not supplied. Higher values give better OCR accuracy at the cost of larger PNG files and slower rendering.

```vb
pdf.DefaultWidth = 3508  ' A4 at 300 dpi landscape
pdf.DefaultWidth = 1240  ' faster preview quality
```

## PDF -> PNG -> Text pipeline

PdfWRT pairs naturally with [VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR) to form a complete in-memory pipeline. No files are written or managed by the caller:

```vb
Dim pdf As New PdfWRT
Dim ocr As New WinOCR

Dim oPages As Collection
Set oPages = pdf.RenderPDFToBytes("C:\Docs\scan.pdf")

Dim page As Variant
Dim sText As String
For Each page In oPages
    Dim aBytes() As Byte
    aBytes = page
    sText = sText & ocr.BytesToText(aBytes) & vbLf
Next page
```

```
PDF
 └─► PdfWRT  (Windows.Data.Pdf)       ->  per-page Byte() arrays in memory
      └─► WinOCR (Windows.Media.Ocr)  ->  extracted text
```

Both components operate entirely within VBA with no subprocess spawning, shell access, or external tools.

---

## How it works

Modern Windows ships with `Windows.Data.Pdf`, a WinRT component exposing a high-quality PDF renderer. Normally accessible only from UWP or .NET, it can be called directly from VBA via `DispCallFunc` vtable dispatch, the same technique used in [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR).

`RenderPDFToBytes` calls `RenderPDFToImages` internally to a GUID-named temp folder, reads each page PNG back as a `Byte()` array via VBA `Open`/`Get`, deletes the files, and returns the collection. The temp-file bridge avoids a WinRT async callback marshalling issue that caused crashes in the direct in-memory path.

---

## Confirmed vtable offsets

Verified on Windows 10 21H2 (build 19044) and Windows 11 23H2 (build 22631).

| Interface | IID | Method | Vtable Offset |
|---|---|---|---|
| `IPdfDocumentStatics` | `{433A0B5F-C007-4788-90F2-08143D922599}` | `LoadFromStreamAsync` | 8 |
| `IPdfDocument` | | `GetPage` | 6 |
| `IPdfDocument` | | `get_PageCount` | 7 |
| `IPdfPage` | | `RenderToStreamAsync` | 6 |
| `IPdfPage` | | `RenderToStreamWithOptionsAsync` | 7 |
| `IPdfPageRenderOptions` | | `set_DestinationWidth` | 9 |

---

## DLL dependencies

All present on every modern Windows installation.

| DLL | Usage |
|---|---|
| `Combase.dll` | `RoGetActivationFactory`, `RoActivateInstance`, `WindowsCreateString`, `WindowsDeleteString` |
| `Shcore.dll` | `CreateRandomAccessStreamOnFile` |
| `oleAut32.dll` | `DispCallFunc` (vtable call dispatcher) |
| `ole32.dll` | `CLSIDFromString`, `CoCreateGuid` |

---

## Error codes

| Code | Meaning |
|---|---|
| 1 | Failed to initialise `Windows.Data.Pdf` WinRT factory |
| 2 | Input file not found |
| 3 | Failed to open PDF file stream |
| 4 | `LoadFromStreamAsync` returned null |
| 5 | `IAsyncInfo` QI failed on document load |
| 6 | `PdfDocument` null after load, file may be corrupt or password-protected |
| 9999 | Timeout waiting for async operation |

---

## Tested environments

| OS | Edition | Version | Build |
|---|---|---|---|
| Windows 10 | Enterprise | 21H2 | 19044 |
| Windows 11 | Enterprise | 23H2 | 22631 |

---

## Credits

- [DanysysTeam/VBA-UWPOCR](https://github.com/DanysysTeam/VBA-UWPOCR): architectural foundation and vtable call pattern
- [rafael-yml/VBA-WinOCR](https://github.com/rafael-yml/VBA-WinOCR): `msvcrt.dll` approach used here

---

## License

MIT License. See [LICENSE](LICENSE) for details.

Copyright © 2026, [rafael-yml](https://rafael-yml.lovable.app/)
