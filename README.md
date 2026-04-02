# vsdx2pdf [libreoffice has feature to convert to pdf. use this app only if u cant use libreoffice. also this is basic.so dont expect much]

`vsdx2pdf` is a standalone command-line converter for turning Visio `.vsdx` files into local PDF files without requiring Microsoft Visio, LibreOffice, cloud APIs, or any other runtime dependency.

The converter is implemented in Go using only the Go standard library. Released binaries are single-file executables for Windows, Linux, and macOS.

## What It Does

- Opens `.vsdx` packages directly
- Renders each Visio page as a PDF page
- Supports nested shapes and background pages
- Supports master-shape inheritance for common stencil-based diagrams
- Renders fills, strokes, text, rotation, flipping, and basic dash styles
- Handles common geometry rows including `MoveTo`, `LineTo`, `ArcTo`, `EllipticalArcTo`, and `Ellipse`

## Current Scope

The tool is aimed at practical diagram export rather than full Visio feature parity. It works well for common flowcharts, block diagrams, and shape/text-heavy documents. Advanced Visio features that are not yet fully covered are listed in [ABOUT.md](ABOUT.md).

## Quick Start

Build from source:

```powershell
go build -o dist/vsdx2pdf.exe ./cmd/vsdx2pdf
```

Convert a file:

```powershell
./dist/vsdx2pdf.exe .\diagram.vsdx
```

Write to a specific output path:

```powershell
./dist/vsdx2pdf.exe -o .\exports\diagram.pdf .\diagram.vsdx
```

List pages without converting:

```powershell
./dist/vsdx2pdf.exe -list-pages .\diagram.vsdx
```

## Releases

The repository includes a GitHub Actions workflow that builds release artifacts for:

- Windows `amd64`
- Linux `amd64`
- macOS `amd64`

Pushing a tag like `v1.0.0` triggers the workflow and publishes zipped binaries to the GitHub release for that tag.

## Documentation

- [ABOUT.md](ABOUT.md)
- [USAGE.md](USAGE.md)
