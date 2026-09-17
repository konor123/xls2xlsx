# XLS2XLSX

Windows GUI utility for converting legacy `.xls` files to `.xlsx`.

## Conversion order

1. If LibreOffice is already installed, the app uses LibreOffice Calc first.
2. If LibreOffice is unavailable or conversion fails, the embedded recovery path is used.

LibreOffice conversion is preferred because it opens the workbook as a spreadsheet document and saves it as XLSX, which gives the best chance of retaining formulas, merged cells, formatting, and workbook structure.

The fallback recovery path handles legacy BIFF/OLE XLS files and also detects several non-standard files distributed with an `.xls` extension, such as HTML, Excel 2003 XML, and delimited text. The fallback is recovery-oriented and cannot guarantee preservation of every advanced workbook feature.

## GUI

- English interface
- Add multiple `.xls` files
- Drag and drop
- Batch conversion
- Original `.xls` files are left unchanged
- Existing output names are preserved by using `_converted`, `_converted_2`, etc.
- Long status text is clipped instead of overflowing the GUI

## Download

Use the EXE from **Releases**, or download the workflow artifact from the **Actions** tab.

A prebuilt copy is also included under `dist/XLS2XLSX.exe`.

## Build

Requirements:

- Go 1.23+
- Python 3 only for rebuilding `payload.zip`

```powershell
python tools/make_payload.py
$env:GOOS="windows"
$env:GOARCH="amd64"
$env:CGO_ENABLED="0"
go build -trimpath -ldflags="-H=windowsgui -s -w" -o dist/XLS2XLSX.exe .
```

## Release

Pushing a tag such as `v0.1.0` builds the Windows x64 EXE and attaches it to a GitHub Release.

```powershell
git tag v0.1.0
git push origin v0.1.0
```

## Third-party component

The embedded XLSX writer uses XlsxWriter. Its license is included in `third_party/XlsxWriter_LICENSE.txt`.
